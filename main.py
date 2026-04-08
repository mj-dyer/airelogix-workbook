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
    documents: list  # [{type: str, base64: str, mediaType: str, filename: str}]
    borrowerType: Optional[str] = "individual"

@app.post("/extract")
def extract_documents(req: ExtractionRequest):
    """
    Receive base64-encoded documents from the wizard.
    Route each to the correct extraction prompt.
    Return structured financial data ready for the spreading engine.
    """
    if not ANTHROPIC_API_KEY:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not configured")

    try:
        import httpx
        results = {
            "tax_returns": [],
            "k1s": [],
            "pfs": None,
            "liquidity": [],
            "other": []
        }

        for doc in req.documents:
            doc_type = doc.get("type", "other")
            base64_data = doc.get("base64", "")
            media_type = doc.get("mediaType", "application/pdf")
            filename = doc.get("filename", "")

            if not base64_data:
                continue

            print(f"[/extract] Processing {filename} ({doc_type})")

            prompt = _get_extraction_prompt(doc_type)
            if not prompt:
                continue

            # Build content block
            if media_type == "application/pdf" or filename.lower().endswith(".pdf"):
                content_block = {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": base64_data
                    }
                }
            else:
                # Word doc or other — send as text extraction request
                content_block = {
                    "type": "text",
                    "text": f"[Document: {filename}] — extract financial data per the instructions below."
                }

            payload = {
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 2000,
                "system": "You are a financial data extraction specialist for aviation finance underwriting. Extract only the requested fields. Respond ONLY with valid JSON — no markdown, no backticks, no explanation.",
                "messages": [{
                    "role": "user",
                    "content": [
                        content_block,
                        {"type": "text", "text": prompt}
                    ]
                }]
            }

            response = httpx.post(
                ANTHROPIC_URL,
                headers={
                    "x-api-key": ANTHROPIC_API_KEY,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json"
                },
                json=payload,
                timeout=60.0
            )

            if response.status_code != 200:
                print(f"[/extract] Claude API error {response.status_code}: {response.text[:200]}")
                continue

            raw = response.json()
            text = ""
            for block in raw.get("content", []):
                if block.get("type") == "text":
                    text += block.get("text", "")

            # Clean and parse JSON
            text = text.strip()
            if text.startswith("```"):
                text = re.sub(r"```[a-z]*\n?", "", text).replace("```", "").strip()

            try:
                extracted = __import__("json").loads(text)
                print(f"[/extract] Extracted from {filename}: {list(extracted.keys())}")
            except Exception as e:
                print(f"[/extract] JSON parse error for {filename}: {e}\nRaw: {text[:200]}")
                continue

            # Route to correct bucket
            if doc_type in ("tax_returns", "tax_return", "1040"):
                results["tax_returns"].append({"filename": filename, "data": extracted})
            elif doc_type in ("k1s", "k1", "k1_schedule"):
                results["k1s"].append({"filename": filename, "data": extracted})
            elif doc_type in ("pfs", "personal_financial_statement"):
                results["pfs"] = {"filename": filename, "data": extracted}
            elif doc_type in ("liquidity", "bank_statement", "brokerage_statement"):
                results["liquidity"].append({"filename": filename, "data": extracted})
            else:
                results["other"].append({"filename": filename, "data": extracted})

        # Synthesize into engine-ready financial summary
        summary = _synthesize_financials(results)
        print(f"[/extract] Synthesis complete: {summary}")

        return {
            "success": True,
            "raw": results,
            "summary": summary
        }

    except Exception as e:
        err = traceback.format_exc()
        print(f"[/extract] ERROR:\n{err}")
        raise HTTPException(status_code=500, detail=str(e))


def _get_extraction_prompt(doc_type: str) -> str:
    """Return the extraction prompt for each document type."""

    if doc_type in ("tax_returns", "tax_return", "1040"):
        return """Extract from this federal tax return (Form 1040). Return JSON with exactly these fields:
{
  "taxYear": 2024,
  "filingStatus": "Married Filing Jointly",
  "wages": 0,
  "taxableInterest": 0,
  "qualifiedDividends": 0,
  "scheduleEIncome": 0,
  "otherIncome": 0,
  "totalIncome": 0,
  "agi": 0,
  "k1Entities": [
    {
      "entityName": "Entity Name LLC",
      "ein": "00-0000000",
      "ordinaryIncome": 0,
      "netRentalIncome": 0,
      "section179": 0,
      "participationType": "General Partner",
      "ownershipPct": 100,
      "entityType": "operating"
    }
  ]
}
Notes: scheduleEIncome is from Schedule E Part II (K-1 pass-through total). Include ALL K-1 entities from Schedule E. entityType is "operating", "real_estate", or "lp_passive". section179 is always negative (deduction). Use 0 for missing fields, never null."""

    elif doc_type in ("k1s", "k1", "k1_schedule"):
        return """Extract from this Schedule K-1. Return JSON with exactly these fields:
{
  "taxYear": 2024,
  "entityName": "Entity Name LLC",
  "ein": "00-0000000",
  "partnerName": "Taxpayer Name",
  "ownershipPct": 100,
  "participationType": "General Partner",
  "entityType": "operating",
  "box1OrdinaryIncome": 0,
  "box2NetRental": 0,
  "box13Section179": 0,
  "box20NetIncome": 0,
  "selfEmploymentIncome": 0
}
Notes: entityType is "operating" for business LLCs, "real_estate" for rental/RE entities, "lp_passive" for limited partner interests with no management role. box13Section179 is the Section 179 deduction (positive number). Use 0 for missing fields."""

    elif doc_type in ("pfs", "personal_financial_statement"):
        return """Extract from this Personal Financial Statement. Return JSON with exactly these fields:
{
  "totalAssets": 0,
  "cashAndBankAccounts": 0,
  "brokerageAndInvestments": 0,
  "retirementAccounts": 0,
  "realEstateValue": 0,
  "businessInterests": 0,
  "otherAssets": 0,
  "totalLiabilities": 0,
  "mortgageBalances": 0,
  "vehicleLoans": 0,
  "creditCardBalances": 0,
  "marginLoans": 0,
  "otherLiabilities": 0,
  "netWorth": 0,
  "annualIncome": 0,
  "monthlyLivingExpenses": 0
}
Use 0 for missing fields. All values in dollars (no $ signs, no commas)."""

    elif doc_type in ("liquidity", "bank_statement", "brokerage_statement"):
        return """Extract from this bank or brokerage statement. Return JSON with exactly these fields:
{
  "institution": "Bank Name",
  "accountType": "Checking",
  "statementDate": "2025-03-31",
  "endingBalance": 0,
  "marginLoanBalance": 0,
  "pledgedAmount": 0,
  "locBalance": 0,
  "locLimit": 0,
  "netAvailable": 0,
  "monthlyDeposits": [],
  "recurringDebits": [
    {"description": "Mortgage Payment", "amount": 0, "frequency": "monthly"}
  ]
}
Notes: marginLoanBalance is any margin loan outstanding against the account. pledgedAmount is any amount pledged as collateral for a line of credit. locBalance is any line of credit balance drawn. recurringDebits should capture mortgage payments, loan payments, and other regular obligations visible in the transaction detail. netAvailable = endingBalance - marginLoanBalance - pledgedAmount. Use 0 for missing fields."""

    return ""


def _synthesize_financials(results: dict) -> dict:
    """
    Synthesize extracted document data into engine-ready financial summary.
    Applies v1.6 methodology: lower year governs, no §179 add-backs,
    LP passive excluded, real estate uses net after depreciation.
    """
    import json as _json

    # ── Income from tax returns ───────────────────────────────────────────────
    tax_years = {}
    for tr in results.get("tax_returns", []):
        data = tr.get("data", {})
        year = data.get("taxYear", 0)
        if not year:
            continue

        # Qualifying income per v1.6:
        # - Operating K-1s: Box 1 ordinary income (no §179 add-back)
        # - Real estate K-1s: net after depreciation (box20NetIncome equivalent)
        # - LP passive: EXCLUDED
        # - Interest + dividends: included
        # - Other income: included if recurring

        qualifying = 0
        k1_detail = []

        for entity in data.get("k1Entities", []):
            etype = entity.get("entityType", "operating")
            if etype == "lp_passive":
                k1_detail.append({
                    "entityName": entity.get("entityName", ""),
                    "ordinaryIncome": entity.get("ordinaryIncome", 0),
                    "qualifyingIncome": 0,
                    "excluded": True,
                    "reason": "LP passive — excluded per methodology"
                })
                continue

            if etype == "real_estate":
                # Use net after depreciation (scheduleE net)
                net = entity.get("netRentalIncome", 0) or entity.get("ordinaryIncome", 0)
                qualifying += net
                k1_detail.append({
                    "entityName": entity.get("entityName", ""),
                    "ordinaryIncome": entity.get("ordinaryIncome", 0),
                    "qualifyingIncome": net,
                    "excluded": False,
                    "note": "Net after depreciation (real estate)"
                })
            else:
                # Operating: Box 1 ordinary income, no §179 add-back
                income = entity.get("ordinaryIncome", 0)
                qualifying += income
                k1_detail.append({
                    "entityName": entity.get("entityName", ""),
                    "ordinaryIncome": income,
                    "qualifyingIncome": income,
                    "excluded": False
                })

        # Add interest, dividends, other income
        interest = data.get("taxableInterest", 0)
        dividends = data.get("qualifiedDividends", 0)
        other = data.get("otherIncome", 0)
        qualifying += interest + dividends + other

        tax_years[year] = {
            "taxYear": year,
            "totalIncome": data.get("totalIncome", 0),
            "qualifyingIncome": qualifying,
            "k1Detail": k1_detail,
            "interest": interest,
            "dividends": dividends,
            "otherIncome": other
        }

    # Lower year governs
    governing_income = 0
    governing_year = 0
    tax_year_list = sorted(tax_years.values(), key=lambda x: x["taxYear"])

    if tax_year_list:
        governing = min(tax_year_list, key=lambda x: x["qualifyingIncome"])
        governing_income = governing["qualifyingIncome"]
        governing_year = governing["taxYear"]

    # ── Liquid assets from statements ─────────────────────────────────────────
    total_liquid = 0
    total_loc_balance = 0
    total_margin_loans = 0
    monthly_existing_debt = 0

    for stmt in results.get("liquidity", []):
        data = stmt.get("data", {})
        net = data.get("netAvailable", 0) or data.get("endingBalance", 0)
        total_liquid += net
        total_loc_balance += data.get("locBalance", 0)
        total_margin_loans += data.get("marginLoanBalance", 0)

        # Capture recurring debits (mortgages, loans)
        for debit in data.get("recurringDebits", []):
            if debit.get("frequency") == "monthly":
                monthly_existing_debt += debit.get("amount", 0)

    # ── PFS data ─────────────────────────────────────────────────────────────
    pfs = results.get("pfs") or {}
    pfs_data = pfs.get("data", {})
    total_assets = pfs_data.get("totalAssets", 0)
    total_liabilities = pfs_data.get("totalLiabilities", 0)
    net_worth = pfs_data.get("netWorth", 0) or (total_assets - total_liabilities)

    # If no PFS, estimate from statements
    if not total_assets and total_liquid:
        total_assets = total_liquid
        net_worth = total_liquid - total_loc_balance - total_margin_loans

    # Use liquid from statements if better than PFS
    liquid_assets = total_liquid if total_liquid > pfs_data.get("cashAndBankAccounts", 0) else (
        pfs_data.get("cashAndBankAccounts", 0) +
        pfs_data.get("brokerageAndInvestments", 0)
    )

    # Net liquid of margin loans and LOC
    liquid_assets_net = max(0, liquid_assets - total_margin_loans - total_loc_balance)

    return {
        "recurringCash": str(int(governing_income)),
        "totalAssets": str(int(total_assets or liquid_assets)),
        "liquidAssets": str(int(liquid_assets_net)),
        "existingDebt": str(int(monthly_existing_debt * 12)),
        "contingentLiabilities": str(int(total_loc_balance + total_margin_loans)),
        "netWorth": str(int(net_worth)),
        "governingYear": governing_year,
        "governingIncome": governing_income,
        "taxYears": tax_year_list,
        "documentSourced": True
    }
