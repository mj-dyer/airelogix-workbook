"""
AireLogix Bio Engine — Autonomous borrower research and profile synthesis.
Fires as a FastAPI BackgroundTask immediately after deal submission.
Uses Claude with a fetch_url tool to autonomously research borrower profiles
from public web sources, then synthesizes a narrative for the lender portal.
"""

import os, re, json
import httpx
from datetime import datetime, timezone
from typing import Optional

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
BIO_MODEL = "claude-sonnet-4-20250514"
MAX_TOOL_CALLS = 12
MAX_PAGE_CHARS = 8000  # Truncate fetched pages to avoid token overflow

FETCH_URL_TOOL = {
    "name": "fetch_url",
    "description": (
        "Fetch and read the plain-text content of a public web page. "
        "Use this to read company websites, news articles, professional profiles, "
        "press releases, industry directories, and Google News RSS feeds. "
        "Returns up to 8,000 characters of stripped plain text."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "url": {
                "type": "string",
                "description": "The full URL to fetch (must begin with http:// or https://)"
            },
            "purpose": {
                "type": "string",
                "description": "Brief note on what you are looking for on this page"
            }
        },
        "required": ["url"]
    }
}


async def _fetch_url(url: str) -> str:
    """
    Fetch a URL and return stripped plain text (up to MAX_PAGE_CHARS).
    Never raises — returns an error string on failure.
    """
    try:
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        }
        async with httpx.AsyncClient(follow_redirects=True, timeout=20.0) as client:
            resp = await client.get(url, headers=headers)
            if resp.status_code != 200:
                return f"[Fetch failed: HTTP {resp.status_code} for {url}]"
            html = resp.text
    except Exception as e:
        return f"[Fetch failed: {e}]"

    # Strip scripts and styles first, then all tags
    text = re.sub(r"<script[^>]*?>[\s\S]*?</script>", " ", html, flags=re.I)
    text = re.sub(r"<style[^>]*?>[\s\S]*?</style>", " ", text, flags=re.I)
    text = re.sub(r"<!--[\s\S]*?-->", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    # Decode common HTML entities
    for ent, char in [("&amp;","&"),("&lt;","<"),("&gt;",">"),("&quot;",'"'),("&#39;","'"),("&nbsp;"," ")]:
        text = text.replace(ent, char)
    text = re.sub(r"&[a-zA-Z0-9#]+;", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    if len(text) > MAX_PAGE_CHARS:
        return text[:MAX_PAGE_CHARS] + " …[truncated]"
    return text


def _build_research_prompt(
    borrower_name: str,
    deal_type: str,
    profile_data: dict,
    analysis: dict,
) -> str:
    """Build the initial research prompt for Claude."""
    is_corp = deal_type in ("corporate", "private_company", "company")

    # Deal context for the record header and verification
    loan_amount = analysis.get("transaction", {}).get("loanAmount", 0)
    aircraft = analysis.get("aircraft", {})
    aircraft_desc = f"{aircraft.get('year','')} {aircraft.get('make','')} {aircraft.get('model','')}".strip()

    # Structured fields the borrower provided (empty dict if Profile step not yet in wizard)
    occupation      = profile_data.get("occupation", "")
    employer        = profile_data.get("employer", "")
    source_of_wealth = profile_data.get("sourceOfWealth", "")
    company_desc    = profile_data.get("companyDescription", "")
    industry        = profile_data.get("industry", "")
    founded         = profile_data.get("yearFounded", "")
    headcount       = profile_data.get("headcount", "")
    provided_urls   = profile_data.get("urls", [])     # URLs borrower supplied
    uploaded_text   = profile_data.get("uploadedText", "")  # Text from uploaded bio doc

    # Guarantors for cross-referencing
    guarantors = analysis.get("guarantors", [])
    guarantor_names = [g.get("name", "") for g in guarantors if g.get("name")]

    lines = [
        "You are a senior aviation finance credit analyst preparing a borrower background profile",
        "for inclusion in a credit memo. Your job is to research this borrower thoroughly using",
        "publicly available information, then produce a structured research report.",
        "",
        "═══════════════════════════════════════════",
        "DEAL RECORD",
        "═══════════════════════════════════════════",
        f"Borrower:       {borrower_name}",
        f"Deal type:      {'Corporate / Business Entity' if is_corp else 'Individual (High Net Worth)'}",
        f"Aircraft:       {aircraft_desc}",
        f"Loan amount:    ${loan_amount:,.0f}",
    ]

    if occupation:       lines.append(f"Stated occupation: {occupation}")
    if employer:         lines.append(f"Stated employer:   {employer}")
    if source_of_wealth: lines.append(f"Source of wealth:  {source_of_wealth}")
    if is_corp:
        if company_desc: lines.append(f"Company overview:  {company_desc}")
        if industry:     lines.append(f"Industry:          {industry}")
        if founded:      lines.append(f"Year founded:      {founded}")
        if headcount:    lines.append(f"Headcount:         {headcount}")
    if guarantor_names:
        lines.append(f"Guarantors/Principals: {', '.join(guarantor_names)}")
    if provided_urls:
        lines.append(f"Borrower-provided URLs: {', '.join(provided_urls)}")
    if uploaded_text:
        lines += ["", "UPLOADED BIO DOCUMENT:", uploaded_text[:3000], ""]

    lines += [
        "",
        "═══════════════════════════════════════════",
        "RESEARCH INSTRUCTIONS",
        "═══════════════════════════════════════════",
        "",
        "Use fetch_url to gather information. Approach:",
        "",
        "1. Start with any borrower-provided URLs.",
        "",
        "2. Search for the borrower/company using Google News RSS:",
        "   https://news.google.com/rss/search?q=NAME+TERMS&hl=en&gl=US&ceid=US:en",
        "   (Replace NAME+TERMS with URL-encoded borrower name and relevant keywords.)",
        "   Parse the RSS XML for headlines, dates, and article URLs, then fetch",
        "   the most relevant articles.",
        "",
        "3. For INDIVIDUALS: research professional background, business ownership,",
        "   board positions, philanthropic roles, any public filings or press.",
        "   Try the company/firm website if employer is known.",
        "",
        "4. For COMPANIES: find the company website, About/Leadership page,",
        "   news coverage, industry standing, key contract awards or milestones.",
        "   Also research named principals/guarantors individually.",
        "",
        "5. Try additional public sources where relevant:",
        "   - Company website (about, team, news pages)",
        "   - LinkedIn company page (may be restricted — note if so)",
        "   - Crunchbase, Bloomberg, industry publications",
        "   - BBB, state corporation filings if useful for verification",
        "",
        "═══════════════════════════════════════════",
        "VERIFICATION",
        "═══════════════════════════════════════════",
        "",
        "Cross-reference what you find publicly against the stated fields above.",
        "Flag substantive discrepancies only (headcount, revenue scale, ownership",
        "structure, current titles, operational status). Do NOT flag minor wording",
        "differences. Severity: INFO (minor/unverifiable) | WATCH (notable gap) |",
        "CRITICAL (direct factual conflict).",
        "",
        "═══════════════════════════════════════════",
        "OUTPUT FORMAT",
        "═══════════════════════════════════════════",
        "",
        "After completing your research, return ONLY a JSON object — no text before",
        "or after the JSON block.",
        "",
        "```json",
        "{",
        '  "narrative": "2-3 paragraphs. Professional, factual, analytical tone suited',
        '                for a bank credit memo underwriter. Paragraph 1: who the borrower',
        '                or company is and their background. Paragraph 2: business or',
        '                professional context, scale, market standing. Paragraph 3: any',
        '                relevant news, public record, or notable findings. For individuals,',
        '                include news coverage here.",',
        '  "sources": [',
        '    { "url": "https://...", "title": "Page or article title", "relevance": "What you found" }',
        '  ],',
        '  "newsItems": [',
        '    { "headline": "...", "date": "YYYY-MM or approximate", "url": "...", "summary": "1-2 sentences" }',
        '  ],',
        '  "verificationFlags": [',
        '    { "field": "...", "stated": "...", "found": "...", "severity": "INFO|WATCH|CRITICAL" }',
        '  ],',
        '  "researchQuality": "high|medium|low",',
        '  "researchNote": "One sentence on the depth and reliability of what was found."',
        "}",
        "```",
        "",
        "If no meaningful public information is found, write a narrative derived from",
        "the application data alone. Set researchQuality to 'low'.",
    ]

    return "\n".join(lines)


def _extract_json(text: str) -> Optional[dict]:
    """
    Try progressively looser strategies to extract a JSON object from Claude's response.
    Returns parsed dict or None.
    """
    # 1. Strip markdown code fences then try direct parse
    cleaned = re.sub(r"```[a-z]*\n?", "", text).replace("```", "").strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # 2. Find the outermost { ... } span and parse that
    start = text.find("{")
    if start != -1:
        depth = 0
        for i, ch in enumerate(text[start:], start):
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[start:i+1])
                    except json.JSONDecodeError:
                        break

    # 3. Try json.loads on each line (sometimes Claude emits one-liner JSON)
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("{") and line.endswith("}"):
            try:
                return json.loads(line)
            except json.JSONDecodeError:
                continue

    return None


async def _run_research_loop(prompt: str) -> Optional[dict]:
    """
    Agentic Claude loop: call API, execute fetch_url tool calls, repeat until done.
    Returns parsed JSON result dict, or None on failure.
    """
    if not ANTHROPIC_API_KEY:
        print("[bio_engine] No ANTHROPIC_API_KEY — skipping research")
        return None

    messages = [{"role": "user", "content": prompt}]
    tool_call_count = 0
    turn = 0

    async with httpx.AsyncClient(timeout=120.0) as client:
        while tool_call_count <= MAX_TOOL_CALLS:
            turn += 1
            payload = {
                "model": BIO_MODEL,
                "max_tokens": 4000,
                "tools": [FETCH_URL_TOOL],
                "messages": messages,
            }

            print(f"[bio_engine] Turn {turn}: calling Claude (tool_calls_so_far={tool_call_count}, msg_count={len(messages)})")

            resp = await client.post(
                ANTHROPIC_URL,
                headers={
                    "x-api-key": ANTHROPIC_API_KEY,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json",
                },
                json=payload,
            )

            if resp.status_code != 200:
                print(f"[bio_engine] Turn {turn}: Claude API error {resp.status_code}: {resp.text[:500]}")
                return None

            data = resp.json()
            stop_reason = data.get("stop_reason")
            content = data.get("content", [])
            content_types = [b.get("type") for b in content]
            print(f"[bio_engine] Turn {turn}: stop_reason={stop_reason}  content_types={content_types}")

            # Append assistant turn
            messages.append({"role": "assistant", "content": content})

            if stop_reason == "end_turn":
                # Extract the final text block and parse JSON
                for block in content:
                    if block.get("type") == "text":
                        text = block["text"].strip()
                        result = _extract_json(text)
                        if result is not None:
                            print(f"[bio_engine] Turn {turn}: JSON parsed OK  quality={result.get('researchQuality')}")
                            return result
                        print(f"[bio_engine] Turn {turn}: JSON parse failed — full response:\n{text}")
                        return None
                print(f"[bio_engine] Turn {turn}: end_turn but no text block in content")
                return None

            if stop_reason == "tool_use":
                tool_results = []
                for block in content:
                    if block.get("type") != "tool_use":
                        continue
                    tool_call_count += 1
                    tool_id   = block["id"]
                    tool_name = block["name"]
                    tool_input = block.get("input", {})

                    if tool_name == "fetch_url":
                        url     = tool_input.get("url", "")
                        purpose = tool_input.get("purpose", "")
                        print(f"[bio_engine] fetch_url ({tool_call_count}/{MAX_TOOL_CALLS}): {url[:80]} — {purpose}")
                        result_text = await _fetch_url(url)
                        fetch_ok = not result_text.startswith("[Fetch failed")
                        print(f"[bio_engine] fetch_url result: {'ok' if fetch_ok else 'FAILED'} ({len(result_text)} chars)")
                    else:
                        result_text = f"[Unknown tool: {tool_name}]"

                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": tool_id,
                        "content": result_text,
                    })

                # If we're at or near the tool call limit, tell Claude to wrap up now
                if tool_call_count >= MAX_TOOL_CALLS - 1:
                    print(f"[bio_engine] Approaching tool call limit ({tool_call_count}/{MAX_TOOL_CALLS}) — injecting finalize instruction")
                    tool_results.append({
                        "type": "text",
                        "text": (
                            "You have used most of your research budget. "
                            "Do NOT call fetch_url again. "
                            "Synthesize everything you have gathered so far and produce the final JSON output now. "
                            "If information was limited, set researchQuality to 'low' or 'medium' and write the "
                            "narrative from what you know."
                        ),
                    })

                messages.append({"role": "user", "content": tool_results})

            elif stop_reason == "max_tokens":
                # Claude ran out of tokens mid-response — try to salvage any text block
                print(f"[bio_engine] Turn {turn}: max_tokens hit — attempting to extract partial JSON")
                for block in content:
                    if block.get("type") == "text":
                        text = block["text"].strip()
                        result = _extract_json(text)
                        if result is not None:
                            return result
                print(f"[bio_engine] Turn {turn}: max_tokens and no salvageable JSON")
                return None

            else:
                print(f"[bio_engine] Turn {turn}: unexpected stop_reason={stop_reason}")
                break

    print(f"[bio_engine] Loop ended after {turn} turns / {tool_call_count} tool calls — attempting forced finalization")

    # One final call with tool use disabled — force Claude to synthesize from what it has
    messages.append({
        "role": "user",
        "content": (
            "Research budget exhausted. Do not use any tools. "
            "Using only the information gathered above (and the application data if nothing was found), "
            "produce the final JSON output now. No preamble — output ONLY the JSON block."
        ),
    })
    try:
        async with httpx.AsyncClient(timeout=60.0) as final_client:
            final_resp = await final_client.post(
                ANTHROPIC_URL,
                headers={
                    "x-api-key": ANTHROPIC_API_KEY,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json",
                },
                json={
                    "model": BIO_MODEL,
                    "max_tokens": 4000,
                    "messages": messages,  # No tools — forces text response
                },
            )
            if final_resp.status_code == 200:
                final_data = final_resp.json()
                for block in final_data.get("content", []):
                    if block.get("type") == "text":
                        result = _extract_json(block["text"])
                        if result is not None:
                            print(f"[bio_engine] Forced finalization succeeded  quality={result.get('researchQuality')}")
                            return result
                print(f"[bio_engine] Forced finalization: no parseable JSON in response")
            else:
                print(f"[bio_engine] Forced finalization: API error {final_resp.status_code}")
    except Exception as e:
        print(f"[bio_engine] Forced finalization exception: {e}")

    return None


async def generate_bio(
    deal_id: str,
    borrower_name: str,
    deal_type: str,
    profile_data: dict,
    analysis: dict,
) -> None:
    """
    Entry point called as a FastAPI BackgroundTask immediately after deal submission.
    Researches the borrower, synthesizes a profile, and saves it to the deal record.
    Never raises — all errors are logged.
    """
    from deals_store import update_deal_bio

    print(f"[bio_engine] Starting research — deal={deal_id}  borrower={borrower_name}")

    try:
        prompt = _build_research_prompt(borrower_name, deal_type, profile_data, analysis)
        result = await _run_research_loop(prompt)

        if result:
            borrower_profile = {
                **result,
                "generatedAt": datetime.now(timezone.utc).isoformat(),
                "dealId": deal_id,
                "borrowerName": borrower_name,
            }
            quality = result.get("researchQuality", "unknown")
        else:
            borrower_profile = {
                "narrative": "",
                "sources": [],
                "newsItems": [],
                "verificationFlags": [],
                "researchQuality": "failed",
                "researchNote": "Profile generation failed — no data retrieved.",
                "generatedAt": datetime.now(timezone.utc).isoformat(),
                "dealId": deal_id,
                "borrowerName": borrower_name,
            }
            quality = "failed"

        update_deal_bio(deal_id, borrower_profile)
        print(f"[bio_engine] Profile saved — deal={deal_id}  quality={quality}")

    except Exception:
        import traceback
        print(f"[bio_engine] ERROR for {deal_id}:\n{traceback.format_exc()}")
