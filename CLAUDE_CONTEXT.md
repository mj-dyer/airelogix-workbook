# AireLogix — Claude Code Context Document
**Version:** 2.0 | **April 2026** | Replaces AireLogix_Handoff_v1_8_Updated.md

---

## HOW TO USE THIS FILE

Place this file as `CLAUDE_CONTEXT.md` in the root of each repo. At the start of every Claude Code session:

```
Read CLAUDE_CONTEXT.md for full project context, then tell me what you understand about the current state before we begin.
```

Then state your goal for the session. Claude Code will confirm its understanding and proceed.

---

## ARCHITECTURE — READ THIS FIRST

### Three Repos / Three Deployments

| Repo | Live URL | Deploy Method | Primary File |
|---|---|---|---|
| `airelogix-deploy` | dev.airelogix.com | Vercel — **manual redeploy required** | `public/index.html` |
| `airelogix-lender` | lender.airelogix.com | Vercel — **manual redeploy required** | `public/index.html` |
| `airelogix-workbook` | airelogix-workbook-production.up.railway.app | Railway — **auto-deploys on git push** | `main.py`, `deals_engine.py`, `deals_store.py`, `generate_workbook.py` |

### Critical Frontend Constraint — Never Violate

Both `public/index.html` files are **single compiled HTML files** using **vanilla React `createElement` — NO JSX, NO npm, NO build step**. Every React component is written as:

```javascript
// CORRECT
el("div", {className: "card"}, el("span", null, "Hello"))

// NEVER do this
<div className="card"><span>Hello</span></div>
```

### Mandatory Workflow for Every Code Change

1. **Read** the relevant section of the file first — never edit blind
2. **Make a targeted string replacement** — never rewrite entire files
3. **Run syntax check** after every JS edit: `node --check public/index.html`
4. **Run syntax check** after every Python edit: `python3 -c "import ast; ast.parse(open('file.py').read()); print('OK')"`
5. **Commit and push** — ask for confirmation before pushing

### Critical ASI (Automatic Semicolon Insertion) Rule

The lender portal's `return el()` statement is one massive expression spanning thousands of lines. JavaScript's ASI causes failures when:
- A line ends with a value (`null,`, `"string"`, `)`)
- The next line starts with an expression that could be parsed as continuation

**The fix, used throughout:** Extract complex multiline ternary expressions as **pre-declared variables** before the `return` statement. Examples already in the file: `entityGuarCard`, `standaloneCard`, `corpBSTab`, `hnwBSTab`, `loanFieldEl`.

**Never** put a new complex ternary directly inside the `return el()` body if it spans multiple lines.

---

## ENVIRONMENT VARIABLES

### Railway (airelogix-workbook)
```
DATABASE_URL = ${{Postgres.DATABASE_URL}}   ← Railway reference syntax
ANTHROPIC_API_KEY = sk-ant-...
```

### Vercel (both frontend repos)
```
ANTHROPIC_API_KEY = sk-ant-...   ← Used by Sample Terms Calculator
```

---

## BACKEND FILE MANIFEST (airelogix-workbook)

| File | Purpose |
|---|---|
| `main.py` | FastAPI app — all API endpoints |
| `deals_engine.py` | Credit scoring engine — runs on every deal submission |
| `deals_store.py` | PostgreSQL via pg8000 — deal CRUD |
| `generate_workbook.py` | 9-tab Excel workbook generator |
| `requirements.txt` | fastapi, uvicorn, openpyxl, python-multipart, pg8000, httpx, python-docx |

### Requirements (locked)
```
fastapi==0.110.0
uvicorn==0.29.0
openpyxl==3.1.5
python-multipart==0.0.9
pg8000==1.31.2
httpx==0.27.0
python-docx==1.1.2
```

**Note:** pg8000 is the PostgreSQL driver — psycopg2-binary was dropped due to missing libpq on Railway. Never switch back.

---

## API ENDPOINTS (main.py)

| Method | Path | Purpose |
|---|---|---|
| POST | `/deals` | Submit new deal from wizard — runs engine, saves to Postgres |
| GET | `/deals` | List all deals for lender queue |
| GET | `/deals/{id}` | Get single deal |
| PATCH | `/deals/{id}/status` | Update deal stage — valid: `application_submitted`, `under_review`, `select_lender_pool`, `package_distributed`, `ioi_received`, `lender_selected`, `closed`, `passed`, `archived` |
| DELETE | `/deals/{id}` | Archive deal — sets stage to `archived`, preserves record |
| POST | `/extract` | Extract financials from uploaded documents via Claude API |
| POST | `/parse-spec` | Parse aircraft spec sheet PDF — returns year/make/model/serial/registration/AFTT/engineProgram |
| POST | `/ingest-section11` | Receive raw spreading prompt output → parse Section 11 JSON → save deal |
| POST | `/ioi` | Submit lender Indication of Interest |
| GET | `/iois/{deal_id}` | Get IOIs for a deal |
| GET | `/workbook` | Generate and stream 9-tab Excel workbook |
| GET | `/memo` | Generate and stream credit memo |

---

## CREDIT ENGINE (deals_engine.py)

### Deal Type Routing
The engine branches on `borrowerType` from the submission:

```python
is_corporate = borrowerType in ("private_company", "corporate", "company")
```

- **Individual (HNWI):** GDSCR = gross qualifying income / total pro forma DS
- **Corporate:** DSCR = Adjusted EBITDA / total pro forma DS (standalone test)

### Collateral Curves — All Three Active

| Aircraft | Function | Vintage Range | Data Points |
|---|---|---|---|
| Bombardier Challenger 350 | `get_cl350_fmv()` | 2014–2022 | 27 closed |
| Gulfstream G550 | `get_g550_fmv()` | 2003–2020 | 51 closed |
| Cessna Citation XLS/XLS+ | `get_xls_fmv()` | 2008–2022 | 25 closed |

Engine also routes: G650, G650ER → `get_g550_fmv()` (closest proxy).

**Hourly adjustment model:** Quadratic progressive model:
`hAdj = delta × A + |delta| × delta × B`
where `A = A_SCALE × bb_retail` and `B = B0 × exp(K_EXP × (10 - age))`

### Balloon Payment (LOCKED)
```python
fmv_at_maturity = fmv * ((1 - depr_rate) ** term_years)
balloon = fmv_at_maturity * 0.70
# NO loan amount cap — balloon is purely collateral-based
```
Depreciation rates: CL350 = 3.0%/yr, G550 = 2.8%/yr, XLS = 3.5%/yr

### Rate Methodology
SOFR OIS + 200bps flat spread. SOFR fallback = 4.30% (total = 6.30%).

### Risk Rating Scale
+ is BETTER, − is WORSE.
```
1.00–1.24 → 1  | 1.25–1.49 → 1+ | 1.50–1.74 → 2+ | 1.75–1.99 → 2
2.00–2.24 → 2− | 2.25–2.49 → 3+ | 2.50–2.74 → 3  | 2.75–2.99 → 3−
3.00–3.24 → 4+ | 3.25–3.49 → 4  | 3.50–3.74 → 4− | 3.75–3.99 → 5+
4.00–4.24 → 5  | 4.25–4.49 → 5− | 4.50–4.74 → 6+ | 4.75–4.99 → 6
5.00–5.99 → 6− | 6.00–7.99 → 7  | 8.00 → 8
```

### Corporate Scoring Thresholds (v1.8, LOCKED)
- DSCR: Strong >2.4x | Acceptable 1.2–2.4x | Marginal 1.0–1.2x | Weak <1.0x
- F4 = Financial Statement Quality: Audited=2.0 | Reviewed=3.5 | Compiled=5.5 | Management=8.0
- Current Ratio: Strong >2.0x | Acceptable 1.2–2.0x | Marginal 1.0–1.2x

### Individual Scoring Thresholds (v1.8, LOCKED)
- GDSCR: Strong >1.75x | Acceptable 1.25–1.75x | Marginal <1.25x
- No depreciation add-backs, no year averaging, lower year governs
- Capital gains excluded from qualifying income
- Irrevocable trusts: zero liquidity credit, hard exclusion

---

## EXTRACTION ENDPOINT (/extract)

Routes on `borrowerType`:
- **Individual:** `UNIFIED_PROMPT` — extracts qualifying income, K-1 detail, liquid assets, existing debt, net worth, registration
- **Corporate:** `CORPORATE_PROMPT` — extracts EBITDA by year, revenue by year, entity balance sheet (current assets, total assets, liabilities, NW, current ratio), debt stack, financial statement quality, guarantors

**DOCX support:** `.docx` files are decoded from base64, extracted via `python-docx` (paragraphs + tables → plain text), then sent to Claude as a text block. This fixed the root cause of all corporate extraction failures where docx files were being replaced with `[File uploaded: filename...]` placeholder.

**EIS vs CoA for year derivation:** EIS (Entry Into Service) date takes priority over Certificate of Airworthiness date when deriving aircraft year from spec sheet.

---

## DEMO DEALS

### Calloway (HNW — Primary Demo)
- **Borrower:** James R. Calloway, Nashville TN
- **Deal type:** individual
- **Aircraft:** 2016 Bombardier CL350, S/N 20632, N599JF, 2,424 AFTT, JSSI Essential Select Plus
- **Transaction:** $17M purchase, $15M loan (88.2% LTV), 120mo, 6.30%
- **Payment:** $112,239/mo (balloon-aware)
- **FMV:** $18,252,000 (CL350 v4 curve, JSSI ESP +3.5%)
- **Qualifying income:** $3,935,600 (FY2023 lower year)
- **GDSCR:** 1.80x
- **Existing DS:** $836,858/yr
- **Net liquid:** $25,698,842
- **ORR:** 2 Very Strong
- **Seeded deal ID:** AL-2026-CALLOWAY (hardcoded in deals_store.py)

### Meridian Industrial Solutions (Corporate — Demo)
- **Borrower:** Meridian Industrial Solutions Inc., Columbus OH (S-Corp)
- **Principals:** Robert T. Callahan (CEO, 60%), James W. Hartley (President/COO, 40%)
- **Deal type:** private_company
- **Aircraft:** 2020 Gulfstream G550, S/N 5481, XA-CHG, 1,820 AFTT, Rolls-Royce CorporateCare Enhanced
- **Transaction:** $20.4M purchase, $14M loan (68.6% LTV), 84mo, 6.30%
- **FY2023 Revenue:** $171.2M | FY2022: $158.4M
- **FY2023 Adj EBITDA:** $37.998M | FY2022 (governing — lower year): $30.150M
- **FY2023 Entity NW:** ~$60M | Total Liabilities: ~$41.8M
- **Current Assets:** $61.42M | Current Liabilities: $26.58M | Current Ratio: 2.31x
- **Debt/EBITDA (pro forma):** 0.79x
- **Corporate DSCR:** ~2.81x
- **Financial statements:** Audited by Gibson Krauss & Associates LLP
- **Expected ORR:** 3 Strong

---

## BORROWER APP (airelogix-deploy/public/index.html)

### Wizard Flow
1. Landing page with Sample Terms Calculator
2. Who is the borrower (individual / company)
3. Aircraft details + spec sheet upload
4. Loan preferences (amount/LTV toggle, structure, term)
5. Documents upload
6. Confirmation + extraction + submission

### Key Components
- `BorrowerApplication` — main wizard shell
- `renderAircraftStepExpanded` — aircraft entry + spec upload via `/parse-spec`
- `renderLoanPrefsStep` — loan amount/LTV dual input, structure selection
- `renderDocumentsStep` — upload rows branching on borrowerType
- `ConfirmationStep` — document extraction via `/extract`, submission to `/deals`
- `SampleTermsModal` — landing page calculator (calls Anthropic API directly via browser)
- `CollateralSnapshot` — frontend FMV display using JS curve functions

### Dual Loan Amount / LTV Input
In `renderLoanPrefsStep`, a `loanFieldEl` variable (pre-declared before return) provides a toggle between `$ Amount` and `% LTV` modes. Each auto-calculates the other from purchase price. State field: `loanPrefs.loanInputMode` (`"amount"` | `"ltv"`), `loanPrefs.ltvPct`.

### Corporate Application Flow
When `applicantType === "company"`:
- Individual option hidden in borrower type step
- Company name/entity type collected in personal step
- Corporate documents step shows: Combined Financial Package (1 file, replaces all) + individual upload rows as alternative
- `CAT_TO_TYPE` mapping includes: `corp_combined`, `corp_fin_1`, `corp_fin_2`, `corp_tax_1`, `corp_tax_2`, `corp_debt`

### Spec Upload
`handleSpecUpload` in `BorrowerApplication` calls `API_BASE + "/parse-spec"`. On success, populates: year (EIS priority over CoA), make, model, serial, registration, AFTT, engineProgram, and `setAcQuery` for the aircraft search field.

### API Base
```javascript
const API_BASE = "https://airelogix-workbook-production.up.railway.app";
```

---

## LENDER PORTAL (airelogix-lender/public/index.html)

### Component Architecture
- `App` — auth shell
- `LenderQueue` — deal list + sidebar nav
- `Sidebar` — left nav with deal flow stages
- `DealDetail` — full deal view with tabs
- `CollateralTab`, `GuarantorsTab` — standalone tab components

### Deal Type Auto-Detection
```javascript
var isCorp = a.dealType === "private_company" || a.dealType === "corporate" || 
             a.dealType === "company" || 
             (a.incomeNormalization && a.incomeNormalization.note && 
              a.incomeNormalization.note.indexOf("Corporate") > -1);
```

### Tab Set (switches automatically based on isCorp)

| HNW Deal | Corporate Deal |
|---|---|
| Overview | Overview |
| Scorecard | Scorecard |
| Income | EBITDA |
| Balance Sheet | Corporate B/S |
| Collateral | Collateral |
| Guarantors | Guarantors |
| Flags | Flags |

### Pre-Declared Variables in DealDetail (critical for ASI)
These are computed before the `return el()` statement to avoid JS parser failures:
- `entityGuarCard` — entity guarantee liquidity adjustment card
- `standaloneCard` — EBITDA standalone test card (corporate)
- `corpBSTab` — corporate balance sheet tab content
- `hnwBSTab` — individual balance sheet tab content

### Sidebar Navigation
```javascript
navItems = [
  ["queue","⊞","Deal Queue"],
  ["review","◎","Under Review"],
  ["ioi","✓","IOI Submitted"],
  ["awarded","★","Awarded"],
  ["funded","✓","Funded"],
  ["passed","✕","Passed On"],
  ["archived","□","Archived"]
]
```

### Deal Actions (DealDetail action bar)
- **← Deal Queue** (left, always)
- **Pass on Deal** (right, grouped) — `PATCH /deals/{id}/status` with `status: "passed"`
- **Archive** (right, grouped) — `DELETE /deals/{id}` then `onBack()`

Both Pass and Archive are wrapped in a `display:flex, gap:8` div, right-aligned.

### displayDeals Filter
```javascript
displayDeals = view==="review" ? reviewDeals_ 
  : view==="ioi" ? ioiDeals_
  : view==="passed" ? passedDeals
  : view==="archived" ? archivedDeals
  : activeDeals;
```

Active deals filter: `stage !== "passed" && stage !== "archived"`

### Corporate B/S Tab Fields (corpBSTab)
Reads from `a.balanceSheet`:
- `grossTotalAssets` — Total Assets
- `totalCurrentAssets` — Current Assets (falls back to `tier1NetLiquid`)
- `totalLiabilities` — Total Liabilities
- `totalCurrentLiabilities` — Current Liabilities
- `statedNetWorth` — Entity Net Worth
- `netWorthCoverage` — Net Worth Coverage
- `currentRatio` — Current Ratio (falls back to `liquidityRatio`)
- `leverageRatio` — Leverage (Liabilities/Assets)
- `debtToEbitda` — Debt/EBITDA
- `a.transaction.ltvVsFMV` — LTV vs FMV

All values show `—` when 0 or missing. Tab only shows data if `grossTotalAssets > 0 || statedNetWorth > 0`.

### Tab Label Dict
```javascript
{
  "overview":"Overview","scorecard":"Scorecard","income":"Income",
  "ebitda":"EBITDA","balance":"Balance Sheet","entity":"Corp B/S",
  "corpbs":"Corporate B/S","collateral":"Collateral",
  "guarantors":"Guarantors","flags":"Flags"
}
```

---

## SECTION 11 JSON INTEGRATION

### Endpoint
`POST /ingest-section11`

**Request:**
```json
{
  "raw_response": "<full spreading prompt output>",
  "deal_id": "AL-2026-XXXXXX",
  "borrower_name": "James Calloway"
}
```

**How it works:**
1. Finds last ` ```json ``` ` block in raw text
2. Routes individual vs corporate on `deal_type` field
3. Maps all fields to engine-compatible analysis dict
4. Saves to Postgres via `save_deal()`
5. Deal appears in lender queue immediately

**Individual mapping:**
- `individual.qualifyingIncome` → `incomeNormalization.qualifyingIncome`
- `individual.netLiquidAssets` → `balanceSheet.tier1NetLiquid`
- `individual.netWorth` → `balanceSheet.statedNetWorth`
- `individual.existingAnnualDebtService` → `transaction.existingAnnualDS`
- `individual.ratios.gdscr_pro_forma` → `gdscr.gdscr`

**Corporate mapping:**
- `corporate.ebitda.qualifying` → `incomeNormalization.qualifyingIncome`
- `corporate.balanceSheet.totalCurrentAssets` → `balanceSheet.tier1NetLiquid`
- `corporate.balanceSheet.entityNetWorth` → `balanceSheet.statedNetWorth`
- `corporate.debtStack.existingAnnualDS` → `transaction.existingAnnualDS`
- `corporate.dscr.proForma` → `gdscr.gdscr`

---

## 9-TAB WORKBOOK (generate_workbook.py)

### Tab Structure

**Individual deals:**
1. Dashboard | 2. Income | 3. K-1 Detail | 4. Debt Service | 5. Balance Sheet | 6. Net Worth & Leverage | 7. Collateral | 8. Risk Rating | 9. Trend Analysis

**Corporate deals (tabs 2 & 3 swap):**
1. Dashboard | 2. Revenue & EBITDA | 3. Entity Debt Stack | 4. Debt Service | 5. Balance Sheet | 6. Net Worth & Leverage | 7. Collateral | 8. Risk Rating | 9. Trend Analysis

Python calculates nothing — Excel calculates everything. Run recalc to confirm zero errors before delivery.

---

## BUSINESS MODEL

- **Fee:** Lender-paid 25–50bps success fee; free to borrowers
- **Distribution:** Controlled disclosure (not anonymization)
- **Identity reveal:** After IOI submission + borrower selects lenders to engage
- **Non-circumvention:** Engagement letters, lender platform agreements, timestamped introduction records

---

## OPEN ITEMS — PRIORITIZED

### P1 — Before Demo
1. **Structure Explorer payment not updating** — stale closure bug on `rateStr` in the payment calculator. The payment display doesn't recalculate when rate or term changes because the closure captures the initial `rateStr` value. Fix: use `useMemo` or `useRef` to ensure fresh reads.
2. **Lender routing thresholds** — all lenders show "below appetite" on live deals. Routing logic needs calibration against actual lender appetite profiles.
3. **Corporate branch validation** — Meridian deal needs to be deleted from Postgres and resubmitted after deploying the `.docx` extraction fix + corporate engine branch. Expected: ORR 3, DSCR 2.81x.

### P2 — Demo Polish
4. **Status timeline** — consolidate from current stages to 5 clean stages in the borrower dashboard
5. **Disclosure Step 1 copy** — shorten the text on the first disclosure screen

### P3 — Platform
6. **Existing lender relationship disclosure flow** — after credit engine identifies lender pool, show borrower matches and let them opt lenders in/out, flag pre-existing relationships, add a note. Sits between analysis output and distribution. Not yet built.
7. **Auth / deal-by-email lookup** — sign-in doesn't authenticate against API. Borrowers can't retrieve their deal status by email.
8. **Section 11 async doc generation** — the flow from spreading prompt → `/ingest-section11` → async workbook/memo generation is wired but the async doc generation itself is pending.
9. **Corporate workbook tabs** — Revenue/EBITDA and Entity Debt Stack tabs not yet built for the 9-tab workbook (corporate deals).

---

## KNOWN METHODOLOGY DECISIONS (LOCKED — NEVER CHANGE)

### Individual (HNWI)
- No depreciation add-backs — §179, §168k, amortization: zero
- No year-averaging — lower year = qualifying run-rate
- Capital gains excluded from qualifying income
- Irrevocable trusts: zero liquidity credit
- Margin loans: net available = gross minus margin balance
- LTV flags: ≤80% no flag | 80.1–90% Material | >90% Critical
- LTV is never a hard routing filter or disqualifying condition

### Corporate
- Income numerator: Adjusted EBITDA (NI + Interest + Taxes + D&A, less recurring CapEx >10% EBITDA)
- Depreciation: added back; aircraft D&A flagged if material
- Charter/lease revenue: excluded always
- Lower year governs qualifying EBITDA
- Financial statement quality: Audited=no flag | Reviewed=Informational | Compiled=Material | Management=Critical

### Collateral
- Balloon: `fmv_at_maturity × 0.70` — no loan amount cap
- Program premiums: JSSI ESP = +3.5%, Smart Parts Plus = +2.5%, off-program = −10%
- G550 RRCC: premium applied; off-program G550: discount applied

---

## BRAND

**Name:** AireLogix (formerly FlyFi)
**Tagline:** Aviation Finance Intelligence (DM Mono)
**Wordmark:** Inter 600 "Aire" + Inter 300 "logix" at −0.03em

**Color palette:**
```
Navy backgrounds: #020c1a, #04101f, #071428, #0B1E3A
Navy border: #1a3460
Gold: #E8CC9A (light), #C8A96E, #B8941E, #9A7A14, #A06810
Steel: #8ba4be
```

**Logo:** Two-tone swept parallelogram (wing + raked winglet)
- Wing polygon: "52,42 52,54 16,20 8,14" gradient #E8CC9A→#A06810
- Winglet polygon: "8,14 16,20 12,12 6,4" flat #E8CC9A

---

## CLAUDE CODE SESSION RULES — ALWAYS FOLLOW

```
1. Read the file section before editing — never edit blind
2. Make targeted string replacements — never rewrite entire files
3. After every JS edit: node --check public/index.html
4. After every Python edit: python3 -c "import ast; ast.parse(open('file.py').read())"
5. Pre-declare complex ternaries as variables before return el() in frontend files
6. Never use JSX, npm, or build tools on the frontend
7. Never cap balloon at loan*0.90 — balloon is purely FMV-based
8. Always confirm before git push
9. Commit messages: descriptive and specific ("Fix balloon cap in deals_engine.py")
10. Railway auto-deploys on push; Vercel requires manual redeploy trigger
```

---

## SPREADING PROMPT

The credit spreading prompt is a separate document: `AireLogix_Spreading_Prompt_v1.8.md`. It is a manual tool run in Claude chat sessions. The wizard uses the backend engine (`deals_engine.py`) directly — NOT the spreading prompt. The spreading prompt produces a Section 11 JSON block which can be ingested via `/ingest-section11`.

---

*CLAUDE_CONTEXT.md v2.0 | April 2026 | AireLogix Aviation Finance Intelligence*
*Supersedes AireLogix_Handoff_v1_8_Updated.md*
*Covers all work through April 15, 2026 — three validated demo deals (Calloway, Meridian), corporate engine branch, docx extraction, dual LTV/dollar input, archive function, Archived nav folder, button alignment fixes, Corp B/S tab rebuild.*
