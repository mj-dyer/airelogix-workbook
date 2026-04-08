"""
AireLogix API v0.2
FastAPI -- workbook generation + deal management
"""

import io, re, tempfile, os, traceback
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Any, Optional

from generate_workbook import generate_workbook
from deals_engine import run_analysis
from deals_store import (
    save_deal, load_deal, list_deals,
    update_deal_status, save_ioi, load_iois, generate_deal_id
)

app = FastAPI(title="AireLogix API", version="0.2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

class WorkbookRequest(BaseModel):
    analysis: Any

class DealSubmission(BaseModel):
    personal: dict
    aircraft: dict
    financial: dict
    loanPrefs: dict
    borrowerType: Optional[str] = "individual"
    transactionType: Optional[str] = "purchase"

class StatusUpdate(BaseModel):
    status: str

class IOISubmission(BaseModel):
    institution: str
    officerName: str
    officerEmail: str
    officerTitle: Optional[str] = ""
    officerPhone: Optional[str] = ""
    allInRate: float
    spread: Optional[int] = None
    index: Optional[str] = "SOFR OIS"
    term: int
    structure: Optional[str] = "Balloon"
    maxLtv: Optional[float] = None
    recourse: Optional[str] = "Full Recourse"
    prepay: Optional[str] = ""
    conditions: Optional[list] = []
    notes: Optional[str] = ""
    expiryDate: Optional[str] = ""
    ackNonBinding: Optional[bool] = True
    ackNonCirc: Optional[bool] = True
    ackIdentity: Optional[bool] = True


@app.get("/")
def health():
    return {"status": "ok", "service": "AireLogix API", "version": "0.2.0"}


@app.post("/workbook")
def generate(req: WorkbookRequest):
    try:
        analysis = req.analysis
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        generate_workbook(analysis, tmp_path)
        with open(tmp_path, "rb") as f:
            content = f.read()
        os.unlink(tmp_path)
        borrower = re.sub(r"[^\x00-\x7F]", "", analysis.get("borrowerName", "Deal"))
        borrower = re.sub(r"[^a-zA-Z0-9]", "_", borrower).strip("_")
        borrower = re.sub(r"_+", "_", borrower) or "Deal"
        date = re.sub(r"[^a-zA-Z0-9-]", "", analysis.get("analysisDate", "draft"))
        filename = f"AireLogix_UW_Workbook_{borrower}_{date}.xlsx"
        return StreamingResponse(
            io.BytesIO(content),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/deals")
def submit_deal(submission: DealSubmission):
    try:
        data = submission.dict()
        print(f"[/deals] keys={list(data.keys())}")

        deal_id = generate_deal_id()
        data["applicationId"] = deal_id
        print(f"[/deals] id={deal_id}")

        analysis = run_analysis(data)
        print(f"[/deals] rating={analysis['riskRating']['rating']}")

        aircraft = data.get("aircraft", {})
        personal = data.get("personal", {})

        deal = {
            "dealId": deal_id,
            "applicationId": deal_id,
            "status": "select_lender_pool",
            "stage": "select_lender_pool",
            "ioiCount": 0,
            "receivedDate": analysis["analysisDate"],
            "borrowerName": analysis["borrowerName"],
            "borrowerEmail": personal.get("email", ""),
            "aircraft": (str(aircraft.get("year","")) + " " + str(aircraft.get("make","")) + " " + str(aircraft.get("model",""))).strip(),
            "aircraftSub": (
                str(aircraft.get("engineProgram","No program")) + " · " +
                (str(aircraft.get("registration","")).strip() or "N-reg") + " · " +
                "{:,}".format(int(str(aircraft.get("aftt","0")).replace(",","") or 0)) + " AFTT"
            ),
            "loanAmount": analysis["transaction"]["loanAmount"],
            "ltv": analysis["transaction"]["ltv"],
            "rating": analysis.get("riskRating", {}).get("rating", 0),
            "band": analysis.get("riskRating", {}).get("band", ""),
            "disposition": analysis.get("riskRating", {}).get("disposition", ""),
            "gdscr": analysis.get("gdscr", {}).get("gdscr", 0),
            "nwCoverage": analysis["balanceSheet"]["netWorthCoverage"],
            "flagCount": len(analysis["flags"]),
            "criticalFlags": len([f for f in analysis["flags"] if f.get("severity") == "CRITICAL"]),
            "analysis": analysis,
            "transactionType": data.get("transactionType", "purchase"),
        }

        save_deal(deal)
        print(f"[/deals] saved {deal_id}")

        return {
            "success": True,
            "dealId": deal_id,
            "applicationId": deal_id,
            "status": "select_lender_pool",
            "rating": analysis.get("riskRating", {}).get("rating", 0),
            "band": analysis.get("riskRating", {}).get("band", ""),
            "gdscr": analysis.get("gdscr", {}).get("gdscr", 0),
            "loanAmount": analysis["transaction"]["loanAmount"],
            "aircraft": deal["aircraft"],
            "receivedDate": analysis["analysisDate"],
        }

    except Exception as e:
        err = traceback.format_exc()
        print(f"[/deals] ERROR:\n{err}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/deals")
def get_deals():
    try:
        deals = list_deals()
        queue = []
        for d in deals:
            queue.append({
                "id": d.get("dealId"),
                "anonId": _anon_id(d.get("dealId", "")),
                "stage": d.get("status", "under_review"),
                "receivedDate": d.get("receivedDate", ""),
                "aircraft": d.get("aircraft", ""),
                "aircraftSub": d.get("aircraftSub", ""),
                "loanAmount": d.get("loanAmount", 0),
                "ltv": d.get("ltv", 0),
                "rating": d.get("rating", "-"),
                "band": d.get("band", ""),
                "disposition": d.get("disposition", ""),
                "gdscr": d.get("gdscr", 0),
                "nwCoverage": d.get("nwCoverage", 0),
                "flagCount": d.get("flagCount", 0),
                "criticalFlags": d.get("criticalFlags", 0),
                "ioiCount": d.get("ioiCount", 0),
                "analysis": d.get("analysis"),
            })
        return {"deals": queue, "count": len(queue)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/deals/{deal_id}")
def get_deal(deal_id: str):
    deal = load_deal(deal_id)
    if not deal:
        raise HTTPException(status_code=404, detail=f"Deal {deal_id} not found")
    result = dict(deal)
    result["anonId"] = _anon_id(deal_id)
    result.pop("borrowerEmail", None)
    return result


@app.patch("/deals/{deal_id}/status")
def patch_status(deal_id: str, update: StatusUpdate):
    valid = [
        "application_submitted", "under_review", "select_lender_pool",
        "package_distributed", "ioi_received", "lender_selected", "closed", "passed"
    ]
    if update.status not in valid:
        raise HTTPException(status_code=400, detail=f"Invalid status: {update.status}")
    deal = update_deal_status(deal_id, update.status)
    if not deal:
        raise HTTPException(status_code=404, detail=f"Deal {deal_id} not found")
    return {"success": True, "dealId": deal_id, "status": update.status}


@app.post("/deals/{deal_id}/ioi")
def submit_ioi(deal_id: str, ioi: IOISubmission):
    deal = load_deal(deal_id)
    if not deal:
        raise HTTPException(status_code=404, detail=f"Deal {deal_id} not found")
    try:
        ioi_data = ioi.dict()
        ioi_id = save_ioi(deal_id, ioi_data)
        return {"success": True, "ioiId": ioi_id, "dealId": deal_id, "refNum": f"IOI-{ioi_id}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/deals/{deal_id}/ioi")
def get_iois(deal_id: str):
    deal = load_deal(deal_id)
    if not deal:
        raise HTTPException(status_code=404, detail=f"Deal {deal_id} not found")
    iois = load_iois(deal_id)
    return {"dealId": deal_id, "iois": iois, "count": len(iois)}


def _anon_id(deal_id: str) -> str:
    parts = deal_id.split("-")
    if len(parts) >= 3:
        return f"DEAL-{parts[-1][:4]}"
    return f"DEAL-{deal_id[-4:]}"


# ── Document Extraction ───────────────────────────────────────────────────────

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"

class ExtractionRequest(BaseModel):
    documents: list
    borrowerType: Optional[str] = "individual"

# System prompt for extraction — instructs Claude to act as a financial analyst
EXTRACTION_SYSTEM = """You are a senior aviation finance credit analyst extracting structured data from borrower financial documents. 
Extract numbers exactly as they appear. Follow these rules precisely:
- Use the LOWER of two tax years as qualifying income (conservative year governs)
- Exclude LP passive K-1 income (Limited Partner with no management role)
- Do NOT add back depreciation, Section 179, or any non-cash deductions
- Real estate K-1s: use net income after depreciation, not gross
- Margin loans and pledged amounts reduce net liquid assets
- Irrevocable trust assets are excluded unless borrower has direct access
- Return ONLY valid JSON, no markdown, no explanation"""

# Single unified extraction prompt — send ALL documents at once for best results
UNIFIED_PROMPT = """Analyze all uploaded financial documents and extract the following data.
Return ONLY a JSON object with exactly this structure:

{
  "borrowerName": "",
  "taxYears": [2023, 2024],
  "governingYear": 2023,
  "qualifyingIncome": 0,
  "k1Detail": [
    {
      "entityName": "",
      "participationType": "Member-Manager",
      "year1Box1": 0,
      "year2Box1": 0,
      "qualifyingIncome": 0,
      "excluded": false,
      "exclusionReason": ""
    }
  ],
  "wagesYear1": 0,
  "wagesYear2": 0,
  "interestIncome": 0,
  "dividendIncome": 0,
  "otherIncome": 0,
  "totalLiquidAssets": 0,
  "cashAndChecking": 0,
  "brokerageAccounts": 0,
  "retirementAccounts": 0,
  "marginLoans": 0,
  "pledgedAmounts": 0,
  "locBalance": 0,
  "netLiquidAssets": 0,
  "totalAssets": 0,
  "totalLiabilities": 0,
  "mortgageBalances": 0,
  "otherLiabilities": 0,
  "netWorth": 0,
  "existingAnnualDebtService": 0,
  "monthlyLivingExpenses": 0,
  "registration": "",
  "dataConfidence": "high"
}

Rules:
- qualifyingIncome = sum of all qualifying K-1 Box 1 income from the LOWER year + wages lower year + interest + dividends
- Exclude LP passive K-1s (set excluded: true, qualifyingIncome: 0)
- Do NOT add back depreciation or Section 179
- netLiquidAssets = totalLiquidAssets - marginLoans - pledgedAmounts - locBalance
- existingAnnualDebtService = sum of all recurring debt payments visible in statements (mortgages, loans, LOCs) annualized
- netWorth = totalAssets - totalLiabilities
- Use 0 for any field not found in the documents"""


@app.post("/extract")
def extract_documents(req: ExtractionRequest):
    if not ANTHROPIC_API_KEY:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not configured")

    try:
        import httpx, json as _json

        docs = req.documents
        if not docs:
            return {"success": True, "summary": _empty_summary(), "raw": {}}

        print(f"[/extract] Processing {len(docs)} documents")

        # Build content blocks — send ALL documents to Claude in one call
        # This gives Claude full context to cross-reference and apply methodology
        content_blocks = []

        for doc in docs:
            base64_data = doc.get("base64", "")
            media_type = doc.get("mediaType", "application/pdf")
            filename = doc.get("filename", "")
            doc_type = doc.get("type", "other")

            if not base64_data:
                continue

            print(f"[/extract] Adding {filename} ({doc_type})")

            if "pdf" in media_type.lower() or filename.lower().endswith(".pdf"):
                content_blocks.append({
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": base64_data
                    }
                })
            # Non-PDF files: add as text note
            else:
                content_blocks.append({
                    "type": "text",
                    "text": f"[File uploaded: {filename} — type: {doc_type}. Extract any financial data visible.]"
                })

        # Add the extraction prompt
        content_blocks.append({
            "type": "text",
            "text": UNIFIED_PROMPT
        })

        payload = {
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 4000,
            "system": EXTRACTION_SYSTEM,
            "messages": [{
                "role": "user",
                "content": content_blocks
            }]
        }

        print(f"[/extract] Calling Claude API with {len(content_blocks)-1} document(s)")

        response = httpx.post(
            ANTHROPIC_URL,
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json"
            },
            json=payload,
            timeout=120.0
        )

        if response.status_code != 200:
            print(f"[/extract] Claude API error {response.status_code}: {response.text[:300]}")
            return {"success": False, "summary": _empty_summary(), "error": f"Claude API error {response.status_code}"}

        raw_response = response.json()
        text = ""
        for block in raw_response.get("content", []):
            if block.get("type") == "text":
                text += block.get("text", "")

        # Clean JSON
        text = text.strip()
        if text.startswith("```"):
            import re
            text = re.sub(r"```[a-z]*", "", text).replace("```", "").strip()

        print(f"[/extract] Raw response: {text[:500]}")

        extracted = _json.loads(text)
        print(f"[/extract] Extraction successful — qualifyingIncome={extracted.get('qualifyingIncome')}, netWorth={extracted.get('netWorth')}")

        summary = _build_summary(extracted)
        print(f"[/extract] Summary: {summary}")

        return {
            "success": True,
            "raw": extracted,
            "summary": summary
        }

    except Exception as e:
        err = traceback.format_exc()
        print(f"[/extract] ERROR:\n{err}")
        return {"success": False, "summary": _empty_summary(), "error": str(e)}


@app.post("/extract/debug")
def extract_debug(req: ExtractionRequest):
    """Returns raw Claude response for debugging — do not expose in production."""
    result = extract_documents(req)
    return result


def _empty_summary() -> dict:
    return {
        "recurringCash": "0",
        "totalAssets": "0",
        "liquidAssets": "0",
        "existingDebt": "0",
        "contingentLiabilities": "0",
        "netWorth": "0",
        "documentSourced": False
    }


def _build_summary(extracted: dict) -> dict:
    """
    Map extracted fields to the engine-ready financial summary format.
    All values as strings to match wizard field format.
    """
    qualifying = extracted.get("qualifyingIncome", 0) or 0
    net_liquid = extracted.get("netLiquidAssets", 0) or 0
    total_assets = extracted.get("totalAssets", 0) or 0
    total_liabilities = extracted.get("totalLiabilities", 0) or 0
    net_worth = extracted.get("netWorth", 0) or (total_assets - total_liabilities)
    existing_ds = extracted.get("existingAnnualDebtService", 0) or 0
    contingent = (extracted.get("marginLoans", 0) or 0) + (extracted.get("locBalance", 0) or 0)

    # If net liquid not directly available, compute it
    if not net_liquid:
        gross_liquid = extracted.get("totalLiquidAssets", 0) or 0
        margin = extracted.get("marginLoans", 0) or 0
        pledged = extracted.get("pledgedAmounts", 0) or 0
        loc = extracted.get("locBalance", 0) or 0
        net_liquid = max(0, gross_liquid - margin - pledged - loc)

    return {
        "recurringCash": str(int(qualifying)),
        "totalAssets": str(int(total_assets or net_liquid)),
        "liquidAssets": str(int(net_liquid)),
        "existingDebt": str(int(existing_ds)),
        "contingentLiabilities": str(int(contingent)),
        "netWorth": str(int(net_worth)),
        "governingYear": extracted.get("governingYear"),
        "taxYears": extracted.get("taxYears", []),
        "k1Detail": extracted.get("k1Detail", []),
        "registration": extracted.get("registration", ""),
        "documentSourced": True,
        "dataConfidence": extracted.get("dataConfidence", "high")
    }
