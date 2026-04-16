"""
AireLogix API v0.2
FastAPI -- workbook generation + deal management
"""

import io, re, tempfile, os, traceback
import httpx
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


class Section11Request(BaseModel):
    raw_response: str          # Full spreading prompt response text
    deal_id: Optional[str] = None  # If provided, updates existing deal
    borrower_name: Optional[str] = None


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
        "package_distributed", "ioi_received", "lender_selected", "closed", "passed", "archived"
    ]
    if update.status not in valid:
        raise HTTPException(status_code=400, detail=f"Invalid status: {update.status}")
    deal = update_deal_status(deal_id, update.status)
    if not deal:
        raise HTTPException(status_code=404, detail=f"Deal {deal_id} not found")
    return {"success": True, "dealId": deal_id, "status": update.status}


@app.delete("/deals/{deal_id}")
def archive_deal(deal_id: str):
    """Archive a deal — removes from active queue, preserves record."""
    deal = update_deal_status(deal_id, "archived")
    if not deal:
        raise HTTPException(status_code=404, detail="Deal not found")
    return {"success": True, "deal_id": deal_id, "stage": "archived"}


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
CORPORATE_PROMPT = """Analyze all uploaded corporate financial documents and extract the following data.
Return ONLY a JSON object with exactly this structure:

{
  "entityName": "",
  "entityType": "",
  "fiscalYears": [2023, 2024],
  "governingYear": 2023,
  "financialStatementQuality": "audited",
  "revenue": {
    "year1": 0,
    "year2": 0,
    "qualifying": 0
  },
  "ebitda": {
    "year1": 0,
    "year2": 0,
    "qualifying": 0,
    "margin": 0.0,
    "aircraftDAIncluded": false,
    "aircraftDAAmount": 0
  },
  "netIncome": {
    "year1": 0,
    "year2": 0
  },
  "debtStack": {
    "totalDebt": 0,
    "seniorDebt": 0,
    "existingAnnualDS": 0,
    "debtToEbitda": 0.0
  },
  "balanceSheet": {
    "totalAssets": 0,
    "totalCurrentAssets": 0,
    "totalLiabilities": 0,
    "totalCurrentLiabilities": 0,
    "entityNetWorth": 0,
    "currentRatio": 0.0,
    "leverageRatio": 0.0
  },
  "guarantors": [
    {
      "name": "",
      "ownershipPct": 0.0,
      "netWorth": 0,
      "liquidAssets": 0,
      "guaranteeType": "Full"
    }
  ],
  "revenueConcentrationFlag": false,
  "charterRevenueExcluded": false,
  "charterRevenueAmount": 0,
  "registration": "",
  "dataConfidence": "high"
}

Rules:
- PRIORITY: Read entity-level financial statements (income statement, balance sheet). Ignore personal financial statements of guarantors for entity fields.
- Use the LOWER of two fiscal years as qualifying EBITDA (conservative year governs)
- EBITDA = net income + interest expense + depreciation + amortization. For S-Corps, there is no income tax line — do not add back taxes.
- One-time/non-recurring charges should be excluded from EBITDA (use normalized/adjusted EBITDA if shown)
- Do NOT add back aircraft D&A — it is a real operating cost
- revenue: read Total Net Revenue or Total Revenue from the income statement — not personal income
- ebitda.year1/year2: read from EBITDA reconciliation table if present, otherwise compute from financials
- balanceSheet: read from the entity balance sheet — Total Assets will be tens or hundreds of millions for a corporate borrower, not thousands
- totalCurrentAssets: read the Current Assets subtotal from the balance sheet
- totalCurrentLiabilities: read the Current Liabilities subtotal
- entityNetWorth: read Total Shareholders Equity or Total Equity from the balance sheet
- debtStack.totalDebt: sum of all interest-bearing debt (revolving credit, term loans, notes payable) — NOT accounts payable
- debtStack.existingAnnualDS: annual principal + interest payments on existing debt only (exclude proposed aircraft loan)
- guarantors: read from personal financial statements or guarantor schedules if present
- financialStatementQuality: "audited" | "reviewed" | "compiled" | "management" based on CPA opinion or lack thereof
- currentRatio = totalCurrentAssets / totalCurrentLiabilities
- leverageRatio = totalLiabilities / totalAssets
- debtToEbitda = totalDebt / qualifying EBITDA
- Use 0 for any field not found in the documents"""

UNIFIED_PROMPT = """Analyze all uploaded financial documents and extract the following data.
Return ONLY a JSON object with exactly this structure:

{
  "borrowerName": "",
  "taxYears": [2023, 2024],
  "governingYear": 2023,
  "qualifyingIncome": 0,
  "agi": 0,
  "federalTaxesPaid": 0,
  "k1Detail": [
    {
      "entityName": "",
      "entityType": "LLC",
      "ownershipPct": 0,
      "participationType": "Member-Manager",
      "year1Box1": 0,
      "year2Box1": 0,
      "year1Distributions": 0,
      "year2Distributions": 0,
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
  "liquidDetail": [
    {
      "institution": "",
      "accountType": "",
      "balance": 0,
      "pledged": false,
      "marginLoanBalance": 0,
      "note": ""
    }
  ],
  "marginLoans": 0,
  "pledgedAmounts": 0,
  "locBalance": 0,
  "netLiquidAssets": 0,
  "totalAssets": 0,
  "totalLiabilities": 0,
  "mortgageBalances": 0,
  "otherLiabilities": 0,
  "netWorth": 0,
  "debtDetail": [
    {
      "creditor": "",
      "accountType": "",
      "balance": 0,
      "monthlyPayment": 0,
      "annualPayment": 0,
      "note": ""
    }
  ],
  "trustStructures": [
    {
      "trustName": "",
      "trustType": "revocable",
      "liquidAssets": 0,
      "borrowerHasAccess": false,
      "note": ""
    }
  ],
  "existingAnnualDebtService": 0,
  "monthlyLivingExpenses": 0,
  "registration": "",
  "dataConfidence": "high"
}

Rules:
- qualifyingIncome = sum of all qualifying K-1 Box 1 income from the LOWER year + wages lower year + interest + dividends
- agi = Adjusted Gross Income from Form 1040 Line 11 for the governing year
- federalTaxesPaid = total federal income taxes paid in cash for the governing year (estimated payments + withholding minus any refund received)
- k1Detail.entityType = legal entity type: "LLC", "S-Corp", "Partnership", "C-Corp", or other as stated
- k1Detail.ownershipPct = borrower's ownership percentage in the entity (0-100)
- k1Detail.year1Distributions / year2Distributions = actual cash distributions received from K-1 Schedule M-2 or Box 19/16D
- liquidDetail = one entry per account from personal financial statement or brokerage statements; marginLoanBalance is the outstanding loan against that account
- debtDetail = one entry per creditor/loan found anywhere in the documents (mortgages, auto, student, personal loans, LOCs, margin loans)
- trustStructures = any trust entity mentioned; borrowerHasAccess=true if borrower is trustee or has discretionary withdrawal rights
- Exclude LP passive K-1s (set excluded: true, qualifyingIncome: 0)
- Do NOT add back depreciation or Section 179
- netLiquidAssets = totalLiquidAssets - marginLoans - pledgedAmounts - locBalance
- existingAnnualDebtService = sum of all recurring debt payments visible in statements (mortgages, loans, LOCs) annualized
- netWorth = totalAssets - totalLiabilities
- Use 0 for any numeric field not found in the documents; use [] for arrays with no data found"""



# ── Section 11 JSON Integration ───────────────────────────────────────────────

def _extract_section11_json(raw_text: str) -> Optional[dict]:
    """
    Extract and parse the Section 11 JSON block from spreading prompt output.
    The block is the last ```json ... ``` fenced code block in the response.
    """
    import json as _json, re as _re

    # Find all ```json blocks and take the last one
    pattern = r"```json\s*([\s\S]*?)```"
    matches = _re.findall(pattern, raw_text)
    if not matches:
        # Try without json tag
        pattern2 = r"```\s*(\{[\s\S]*?\})\s*```"
        matches = _re.findall(pattern2, raw_text)
    if not matches:
        return None

    json_str = matches[-1].strip()
    try:
        return _json.loads(json_str)
    except _json.JSONDecodeError as e:
        print(f"[Section11] JSON parse error: {e}")
        # Try cleaning common issues
        json_str = _re.sub(r",\s*}", "}", json_str)
        json_str = _re.sub(r",\s*]", "]", json_str)
        try:
            return _json.loads(json_str)
        except:
            return None


def _map_section11_to_analysis(s11: dict) -> dict:
    """
    Map Section 11 JSON (v1.8 schema) to the analysis format
    expected by the lender portal and workbook generator.
    Handles both individual and corporate deal types.
    """
    deal_type = s11.get("meta", {}).get("deal_type", "individual")
    dp = s11.get("dealParameters", {})
    col = s11.get("collateral", {})
    rr = s11.get("riskRating", {})
    flags = s11.get("flags", [])
    lr = s11.get("lenderRouting", [])
    rs = s11.get("repayment_sources", {})
    ind = s11.get("individual") or {}
    corp = s11.get("corporate") or {}

    # Build analysis dict in engine-compatible format
    analysis = {
        "analysisDate": s11.get("meta", {}).get("analysis_date", ""),
        "applicationId": s11.get("meta", {}).get("deal_id", ""),
        "borrowerName": dp.get("borrower_name", s11.get("meta", {}).get("borrower_last_name", "")),
        "dealType": deal_type,

        "aircraft": {
            "description": f"{dp.get('aircraft_year','')} {dp.get('aircraft_make','')} {dp.get('aircraft_model','')}".strip(),
            "serialNumber": dp.get("aircraft_sn", col.get("sn", "")),
            "registration": dp.get("aircraft_registration", col.get("registration", "")),
            "aftt": col.get("aftt", 0),
            "engineProgram": col.get("engine_program", ""),
            "year": dp.get("aircraft_year", col.get("year", 0)),
            "make": dp.get("aircraft_make", col.get("make", "")),
            "model": dp.get("aircraft_model", col.get("model", "")),
        },

        "collateral": {
            "year": col.get("year", dp.get("aircraft_year", 0)),
            "make": col.get("make", dp.get("aircraft_make", "")),
            "model": col.get("model", dp.get("aircraft_model", "")),
            "sn": col.get("sn", dp.get("aircraft_sn", "")),
            "registration": col.get("registration", dp.get("aircraft_registration", "")),
            "aftt": col.get("aftt", 0),
            "engine_program": col.get("engine_program", ""),
            "fmv_curve": col.get("fmv_curve", col.get("fmv_used", 0)),
            "fmv_purchase_price": col.get("fmv_purchase_price", dp.get("purchase_price", 0)),
            "fmv_used": col.get("fmv_used", col.get("fmv_curve", 0)),
            "estimatedFMV": col.get("fmv_curve", col.get("fmv_used", 0)),
            "ltv_on_curve_fmv": round(col.get("ltv_on_curve_fmv", 0) * 100, 1) if col.get("ltv_on_curve_fmv", 0) < 1 else col.get("ltv_on_curve_fmv", 0),
            "ltv_on_purchase_price": round(dp.get("ltv", 0) * 100, 1) if dp.get("ltv", 0) < 1 else dp.get("ltv", 0),
            "curveName": col.get("curve_name", "AireLogix Curve"),
        },

        "transaction": {
            "purchasePrice": dp.get("purchase_price", 0),
            "loanAmount": dp.get("loan_amount", 0),
            "ltv": dp.get("ltv", 0),
            "ltvVsFMV": col.get("ltv_on_curve_fmv", 0),
            "fmv": col.get("fmv_curve", col.get("fmv_used", 0)),
            "illustrativeRate": dp.get("rate", 0),
            "rateNote": f"SOFR OIS + spread. 30/360 basis. Not a lender commitment.",
            "termMonths": dp.get("term_months", 120),
            "monthlyPayment": dp.get("monthly_payment", 0),
            "annualAircraftDS": dp.get("annual_debt_service_aircraft", (dp.get("monthly_payment", 0) * 12)),
            "existingAnnualDS": (
                ind.get("existingAnnualDebtService", 0) or
                corp.get("debtStack", {}).get("existingAnnualDS", 0)
            ),
            "totalProFormaDS": (
                dp.get("annual_debt_service_aircraft", dp.get("monthly_payment", 0) * 12) +
                (ind.get("existingAnnualDebtService", 0) or
                 corp.get("debtStack", {}).get("existingAnnualDS", 0))
            ),
        },

        "riskRating": {
            "rating": rr.get("final_rating", ""),
            "band": _orr_to_band(rr.get("final_rating", "")),
            "disposition": rr.get("disposition", ""),
            "composite": {"finalComposite": rr.get("weighted_composite", 0)},
            "factorScores": {
                f.get("code", f"F{i+1}"): {
                    "score": f.get("score", 0),
                    "weight": f.get("weight", 0),
                    "weighted": f.get("weighted", 0),
                    "basis": f.get("label", "")
                }
                for i, f in enumerate(rr.get("factors", []))
            }
        },

        "flags": [
            {
                "severity": f.get("severity", ""),
                "title": f.get("title", ""),
                "detail": f.get("risk", ""),
                "mitigant": f.get("mitigant", ""),
                "action": f.get("net_assessment", ""),
            }
            for f in flags
        ],

        "lenderRouting": lr,
        "repaymentSources": {
            "primary": rs.get("primary", ""),
            "primaryBasis": rs.get("primary_justification", ""),
            "secondary": rs.get("secondary", ""),
            "secondaryAssessment": rs.get("secondary_assessment", ""),
            "dualSourceConfidence": rs.get("dual_source_confidence", ""),
        },
        "executiveSummary": s11.get("executive_summary", ""),
        "dataNeeded": s11.get("data_needed", []),
        "section11Raw": s11,  # Preserve full schema for future use
    }

    # Individual branch
    if deal_type == "individual" and ind:
        liq = ind.get("liquidity", {})
        ratios = ind.get("ratios", {})
        analysis["incomeNormalization"] = {
            "qualifyingIncome": ind.get("qualifyingIncome", 0),
            "qualifyingYear": ind.get("governingYear"),
            "taxYears": ind.get("taxYears", []),
            "afterTaxQualifyingIncome": ind.get("qualifyingIncome", 0) * 0.64,
            "k1Detail": [
                {
                    "entityName": k.get("entity", ""),
                    "participationType": k.get("participation", ""),
                    "year1Box1": k.get("year_1_box1", 0),
                    "year2Box1": k.get("year_2_box1", 0),
                    "qualifyingIncome": k.get("normalized", 0),
                    "excluded": k.get("excluded", False),
                    "exclusionReason": k.get("treatment", ""),
                }
                for k in ind.get("k1Detail", [])
            ],
            "year1Total": ind.get("qualifyingIncome", 0),
            "year2Total": ind.get("qualifyingIncome", 0),
        }
        analysis["gdscr"] = {
            "gdscr": ratios.get("gdscr_pro_forma", 0),
            "assessment": _gdscr_label(ratios.get("gdscr_pro_forma", 0)),
            "afterTaxIncome": ind.get("qualifyingIncome", 0) * 0.64,
            "totalAnnualDS": analysis["transaction"]["totalProFormaDS"],
        }
        analysis["balanceSheet"] = {
            "tier1NetLiquid": ind.get("netLiquidAssets", liq.get("eligible", 0)),
            "tier1GrossLiquid": ind.get("totalLiquidAssets", liq.get("gross", 0)),
            "marginLoanAdjustment": ind.get("marginLoans", 0),
            "grossTotalAssets": ind.get("totalAssets", 0),
            "totalLiabilities": ind.get("totalLiabilities", 0),
            "statedNetWorth": ind.get("netWorth", 0),
            "liquidityRatio": ratios.get("liquidity_ratio", 0),
            "liquidityAssessment": _liquidity_label(ratios.get("liquidity_ratio", 0)),
            "netWorthCoverage": ratios.get("net_worth_ratio", 0),
            "leverageRatio": ratios.get("leverage_ratio", 0),
            "tier1Assets": [
                {
                    "description": a.get("account_type", ""),
                    "value": a.get("gross_value", 0),
                    "encumbrance": a.get("margin_balance", 0) + a.get("pledged_amount", 0),
                    "net": a.get("net_eligible", 0),
                }
                for a in liq.get("accounts", [])
            ],
            "entityGuaranteeLiquidity": {"adjustmentApplied": False},
        }
        analysis["guarantors"] = [
            {
                "name": g.get("party", ""),
                "type": g.get("role", ""),
                "netWorth": 0,
                "liquidAssets": 0,
                "guaranteeType": g.get("status", ""),
            }
            for g in ind.get("guarantor_package", [])
        ]

    # Corporate branch
    elif deal_type in ("corporate", "company") and corp:
        ebitda = corp.get("ebitda", {})
        ratios = corp.get("ratios", {})
        bs = corp.get("balanceSheet", {})
        debt = corp.get("debtStack", {})
        analysis["incomeNormalization"] = {
            "qualifyingIncome": ebitda.get("qualifying", 0),
            "qualifyingYear": corp.get("governingYear"),
            "taxYears": corp.get("fiscalYears", []),
            "year1Total": ebitda.get("year1", 0),
            "year2Total": ebitda.get("year2", 0),
            "k1Detail": [],
            "note": f"Corporate deal — EBITDA basis. Entity: {corp.get('entityName','')}",
        }
        analysis["gdscr"] = {
            "gdscr": ratios.get("dscr_pro_forma", 0),
            "assessment": _gdscr_label(ratios.get("dscr_pro_forma", 0)),
            "afterTaxIncome": ebitda.get("qualifying", 0),
            "totalAnnualDS": analysis["transaction"]["totalProFormaDS"],
        }
        analysis["balanceSheet"] = {
            "tier1NetLiquid": bs.get("totalCurrentAssets", 0),
            "tier1GrossLiquid": bs.get("totalCurrentAssets", 0),
            "grossTotalAssets": bs.get("totalAssets", 0),
            "totalLiabilities": bs.get("totalLiabilities", 0),
            "statedNetWorth": bs.get("entityNetWorth", 0),
            "liquidityRatio": bs.get("currentRatio", 0),
            "liquidityAssessment": _liquidity_label(bs.get("currentRatio", 0)),
            "netWorthCoverage": bs.get("entityNetWorth", 0) / analysis["transaction"]["loanAmount"]
                                if analysis["transaction"]["loanAmount"] else 0,
            "leverageRatio": bs.get("leverageRatio", 0),
            "tier1Assets": [],
            "entityGuaranteeLiquidity": {"adjustmentApplied": False},
        }
        analysis["guarantors"] = [
            {
                "name": g.get("name", ""),
                "type": g.get("type", ""),
                "netWorth": g.get("netWorth", 0),
                "liquidAssets": g.get("liquidAssets", 0),
                "guaranteeType": g.get("guaranteeType", "Full"),
            }
            for g in corp.get("guarantors", [])
        ]

    return analysis


def _orr_to_band(rating: str) -> str:
    bands = {
        "1": "Exceptional", "1+": "Exceptional", "1-": "Exceptional",
        "2": "Very Strong", "2+": "Very Strong", "2-": "Very Strong",
        "3": "Strong", "3+": "Strong", "3-": "Strong",
        "4": "Good", "4+": "Good", "4-": "Good",
        "5": "Acceptable", "5+": "Acceptable", "5-": "Acceptable",
        "6": "Marginal", "6+": "Marginal", "6-": "Marginal",
        "7": "Difficult", "8": "Hard Decline",
    }
    return bands.get(str(rating).strip(), "")


def _gdscr_label(g: float) -> str:
    if g >= 2.0: return "Strong"
    if g >= 1.5: return "Adequate"
    if g >= 1.25: return "Marginal"
    return "Below threshold"


def _liquidity_label(r: float) -> str:
    if r >= 2.0: return "Strong"
    if r >= 1.0: return "Adequate"
    if r >= 0.5: return "Marginal"
    return "Critical"


@app.post("/ingest-section11")
async def ingest_section11(req: Section11Request):
    """
    Receive raw spreading prompt output, extract Section 11 JSON,
    map to analysis format, and store/update the deal.
    """
    try:
        # Extract JSON block from spreading prompt response
        s11 = _extract_section11_json(req.raw_response)
        if not s11:
            raise HTTPException(status_code=422, detail="No valid JSON block found in response")

        print(f"[/ingest-section11] deal_type={s11.get('meta',{}).get('deal_type')} "
              f"borrower={s11.get('meta',{}).get('borrower_last_name')}")

        # Map to engine-compatible analysis format
        analysis = _map_section11_to_analysis(s11)

        # If deal_id provided, update existing deal; otherwise create new
        deal_id = req.deal_id or s11.get("meta", {}).get("deal_id") or generate_deal_id()
        analysis["applicationId"] = deal_id

        dp = s11.get("dealParameters", {})
        col = s11.get("collateral", {})
        rr = s11.get("riskRating", {})
        flags = analysis.get("flags", [])

        deal = {
            "dealId": deal_id,
            "applicationId": deal_id,
            "status": "select_lender_pool",
            "stage": "select_lender_pool",
            "ioiCount": 0,
            "receivedDate": s11.get("meta", {}).get("analysis_date", ""),
            "borrowerName": analysis.get("borrowerName", req.borrower_name or ""),
            "borrowerEmail": "",
            "aircraft": f"{dp.get('aircraft_year','')} {dp.get('aircraft_make','')} {dp.get('aircraft_model','')}".strip(),
            "aircraftSub": (
                col.get("engine_program", "") + " · " +
                (col.get("registration", "") or "N-reg") + " · " +
                f"{col.get('aftt', 0):,} AFTT"
            ),
            "loanAmount": dp.get("loan_amount", 0),
            "ltv": dp.get("ltv", 0),
            "rating": rr.get("final_rating", ""),
            "band": _orr_to_band(rr.get("final_rating", "")),
            "disposition": rr.get("disposition", ""),
            "gdscr": analysis.get("gdscr", {}).get("gdscr", 0),
            "nwCoverage": analysis.get("balanceSheet", {}).get("netWorthCoverage", 0),
            "flagCount": len(flags),
            "criticalFlags": len([f for f in flags if f.get("severity") == "CRITICAL"]),
            "analysis": analysis,
            "transactionType": "purchase",
        }

        save_deal(deal)
        print(f"[/ingest-section11] Saved deal {deal_id}")

        return {
            "success": True,
            "dealId": deal_id,
            "rating": rr.get("final_rating", ""),
            "band": _orr_to_band(rr.get("final_rating", "")),
            "gdscr": analysis.get("gdscr", {}).get("gdscr", 0),
            "borrowerName": analysis.get("borrowerName", ""),
            "aircraft": deal["aircraft"],
        }

    except HTTPException:
        raise
    except Exception as e:
        err = traceback.format_exc()
        print(f"[/ingest-section11] ERROR:\n{err}")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/extract")
async def extract_documents(req: ExtractionRequest):
    if not ANTHROPIC_API_KEY:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not configured")

    try:
        import json as _json

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

            fname_lower = filename.lower()
            if "pdf" in media_type.lower() or fname_lower.endswith(".pdf"):
                content_blocks.append({
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": base64_data
                    }
                })
            elif fname_lower.endswith(".docx") or "wordprocessingml" in media_type.lower() or "msword" in media_type.lower():
                # Extract text from docx and send as plain text
                try:
                    import base64 as _b64, io
                    from docx import Document as _DocxDoc
                    raw_bytes = _b64.b64decode(base64_data)
                    doc_obj = _DocxDoc(io.BytesIO(raw_bytes))
                    docx_text = "\n".join(p.text for p in doc_obj.paragraphs if p.text.strip())
                    # Also extract tables
                    for table in doc_obj.tables:
                        for row in table.rows:
                            row_text = "  |  ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                            if row_text:
                                docx_text += "\n" + row_text
                    print(f"[/extract] Extracted {len(docx_text)} chars from {filename}")
                    content_blocks.append({
                        "type": "text",
                        "text": f"=== DOCUMENT: {filename} ({doc_type}) ===\n{docx_text}\n=== END {filename} ==="
                    })
                except Exception as docx_err:
                    print(f"[/extract] DOCX extraction failed for {filename}: {docx_err}")
                    content_blocks.append({
                        "type": "text",
                        "text": f"[Could not extract text from {filename} — {docx_err}]"
                    })
            else:
                # Unknown file type — add filename note
                content_blocks.append({
                    "type": "text",
                    "text": f"[File uploaded: {filename} — type: {doc_type}. Format not supported for extraction.]"
                })

        # Add the extraction prompt
        is_corporate = req.borrowerType in ("private_company", "corporate", "company")
        selected_prompt = CORPORATE_PROMPT if is_corporate else UNIFIED_PROMPT
        print(f"[/extract] Using {'CORPORATE' if is_corporate else 'INDIVIDUAL'} extraction prompt")
        content_blocks.append({
            "type": "text",
            "text": selected_prompt
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

        async with httpx.AsyncClient() as client:
            response = await client.post(
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
async def extract_debug(req: ExtractionRequest):
    """Returns raw Claude response for debugging — do not expose in production."""
    result = await extract_documents(req)
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
    Map extracted fields to engine-ready financial summary format.
    Handles both individual (HNWI) and corporate extraction results.
    """
    # Detect which schema was returned
    is_corporate = "ebitda" in extracted or "entityName" in extracted

    if is_corporate:
        ebitda = extracted.get("ebitda", {}) or {}
        bs = extracted.get("balanceSheet", {}) or {}
        debt = extracted.get("debtStack", {}) or {}
        qualifying = ebitda.get("qualifying", 0) or 0
        total_assets = bs.get("totalAssets", 0) or 0
        total_liabilities = bs.get("totalLiabilities", 0) or 0
        net_worth = bs.get("entityNetWorth", 0) or (total_assets - total_liabilities)
        existing_ds = debt.get("existingAnnualDS", 0) or 0
        # For corporate, use current assets as proxy for liquid assets
        liquid_assets = bs.get("totalCurrentAssets", 0) or 0
        contingent = 0

        return {
            "recurringCash": str(int(qualifying)),
            "totalAssets": str(int(total_assets)),
            "liquidAssets": str(int(liquid_assets)),
            "existingDebt": str(int(existing_ds)),
            "contingentLiabilities": str(int(contingent)),
            "netWorth": str(int(net_worth)),
            "governingYear": extracted.get("governingYear"),
            "taxYears": extracted.get("fiscalYears", []),
            "k1Detail": [],
            "registration": extracted.get("registration", ""),
            "documentSourced": True,
            "dataConfidence": extracted.get("dataConfidence", "high"),
            # Corporate-specific fields passed through for engine
            "ebitda": ebitda,
            "revenue": extracted.get("revenue", {}),
            "debtStack": debt,
            "corporateBalanceSheet": bs,
            "guarantors": extracted.get("guarantors", []),
            "entityName": extracted.get("entityName", ""),
            "entityType": extracted.get("entityType", ""),
            "financialStatementQuality": extracted.get("financialStatementQuality", ""),
            "isCorporate": True,
        }
    else:
        # Individual / HNWI path
        qualifying = extracted.get("qualifyingIncome", 0) or 0
        net_liquid = extracted.get("netLiquidAssets", 0) or 0
        total_assets = extracted.get("totalAssets", 0) or 0
        total_liabilities = extracted.get("totalLiabilities", 0) or 0
        net_worth = extracted.get("netWorth", 0) or (total_assets - total_liabilities)
        existing_ds = extracted.get("existingAnnualDebtService", 0) or 0
        contingent = (extracted.get("marginLoans", 0) or 0) + (extracted.get("locBalance", 0) or 0)

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
            "dataConfidence": extracted.get("dataConfidence", "high"),
            "isCorporate": False,
        }


@app.post("/parse-spec")
async def parse_spec(req: dict):
    """Parse aircraft spec sheet PDF and return structured fields."""
    if not ANTHROPIC_API_KEY:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not configured")

    base64_data = req.get("base64", "")
    if not base64_data:
        raise HTTPException(status_code=400, detail="No base64 data provided")

    try:
        async with httpx.AsyncClient() as client:
            response = await client.post(
                ANTHROPIC_URL,
                headers={
                    "x-api-key": ANTHROPIC_API_KEY,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json"
                },
                json={
                    "model": "claude-sonnet-4-20250514",
                    "max_tokens": 500,
                    "messages": [{
                        "role": "user",
                        "content": [
                            {
                                "type": "document",
                                "source": {
                                    "type": "base64",
                                    "media_type": "application/pdf",
                                    "data": base64_data
                                }
                            },
                            {
                                "type": "text",
                                "text": (
                                "Extract aircraft specification data from this spec sheet. "
                                "Return ONLY valid JSON with exactly these fields:\n"
                                "{\n"
                                "  \"year\": 2019,\n"
                                "  \"make\": \"Bombardier\",\n"
                                "  \"model\": \"Challenger 350\",\n"
                                "  \"serialNumber\": \"20632\",\n"
                                "  \"registration\": \"N599JF\",\n"
                                "  \"totalAirframeHours\": 2424,\n"
                                "  \"engineProgram\": \"JSSI Essential Select Plus\"\n"
                                "}\n"
                                "For year: look for an explicit model year first. If not stated, "
                                "derive it from dates in this priority order: "
                                "(1) Entry into Service date — use this year if present. "
                                "(2) Certificate of Airworthiness date — use only if no EIS date found. "
                                "Entry into Service is preferred because it reflects when the aircraft "
                                "entered the market and is the correct basis for pricing. "
                                "For example: EIS March 2020 → year = 2020. CoA October 2019 with EIS March 2020 → year = 2020. "
                                "Do not leave year null if any delivery, EIS, or airworthiness date is present. "
                                "For engineProgram: look for engine maintenance program enrollment. "
                                "Common programs: JSSI, JSSI Essential Select Plus, Honeywell MSP Gold, "
                                "Honeywell MSP, Rolls-Royce CorporateCare, Rolls-Royce CorporateCare Enhanced, "
                                "Smart Parts Plus. Return exact program name as listed, or null if not found. "
                                "Use null for any field not found. No markdown, no explanation, just JSON."
                            )
                            }
                        ]
                    }]
                },
                timeout=60.0
            )

        if response.status_code != 200:
            raise HTTPException(status_code=502, detail=f"Claude API error {response.status_code}")

        raw = response.json()
        text = ""
        for block in raw.get("content", []):
            if block.get("type") == "text":
                text += block.get("text", "")

        import json as _json, re as _re
        text = _re.sub(r"```[a-z]*", "", text).replace("```", "").strip()
        parsed = _json.loads(text)
        print(f"[/parse-spec] Parsed: {parsed}")
        return {"success": True, "data": parsed}

    except Exception as e:
        err = traceback.format_exc()
        print(f"[/parse-spec] ERROR: {err}")
        raise HTTPException(status_code=500, detail=str(e))
