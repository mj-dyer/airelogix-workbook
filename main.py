"""
AireLogix API
FastAPI service — underwriter workbook generation + deal management endpoints.

Endpoints:
  GET  /                      — health check
  POST /workbook              — generate .xlsx from analysis JSON (existing)
  POST /deals                 — submit new deal from wizard, run credit engine
  GET  /deals                 — list all deals (lender portal queue)
  GET  /deals/{deal_id}       — get full deal detail
  PATCH /deals/{deal_id}/status — advance deal stage
  POST /deals/{deal_id}/ioi   — lender submits IOI
  GET  /deals/{deal_id}/ioi   — get IOIs for a deal (borrower dashboard)
"""

import io
import re
import tempfile
import os
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
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
    allow_origins=[
        "https://lender.airelogix.com",
        "https://dev.airelogix.com",
        "http://localhost:3000",
        "http://localhost:5173",
    ],
    allow_credentials=True,
    allow_methods=["GET", "POST", "PATCH", "OPTIONS"],
    allow_headers=["Content-Type", "Authorization"],
)


# ── Pydantic models ───────────────────────────────────────────────────────────

class WorkbookRequest(BaseModel):
    analysis: Any


class DealSubmission(BaseModel):
    """Wizard submission payload from borrower app."""
    personal: dict
    aircraft: dict
    financial: dict
    loanPrefs: dict
    borrowerType: Optional[str] = "individual"
    transactionType: Optional[str] = "purchase"


class StatusUpdate(BaseModel):
    status: str


class IOISubmission(BaseModel):
    """IOI payload from lender portal."""
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


# ── Existing endpoints ────────────────────────────────────────────────────────

@app.get("/")
def health():
    return {"status": "ok", "service": "AireLogix API", "version": "0.2.0"}


@app.post("/workbook")
def generate(req: WorkbookRequest):
    """Generate underwriter workbook (.xlsx) from analysis JSON."""
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


# ── Deal endpoints ────────────────────────────────────────────────────────────

@app.post("/deals")
def submit_deal(submission: DealSubmission):
    """
    Receive wizard submission, run credit analysis engine, store deal.
    Returns deal ID and summary for borrower dashboard.
    """
    try:
        data = submission.dict()

        # Generate deal ID
        deal_id = generate_deal_id()
        data["applicationId"] = deal_id

        # Run credit analysis engine
        analysis = run_analysis(data)

        # Build deal record
        aircraft = data.get("aircraft", {})
        aircraft_year = aircraft.get("year", "")
        aircraft_make = aircraft.get("make", "")
        aircraft_model = aircraft.get("model", "")

        personal = data.get("personal", {})
        loan_prefs = data.get("loanPrefs", {})

        deal = {
            "dealId": deal_id,
            "applicationId": deal_id,
            "status": "under_review",
            "stage": "under_review",
            "ioiCount": 0,
            "receivedDate": analysis["analysisDate"],

            # Borrower summary (identity withheld from lenders until engagement)
            "borrowerName": analysis["borrowerName"],
            "borrowerEmail": personal.get("email", ""),

            # Aircraft summary (shown to lenders)
            "aircraft": f"{aircraft_year} {aircraft_make} {aircraft_model}".strip(),
            "aircraftSub": (
                f"{aircraft.get('engineProgram', 'No program')} · "
                f"{aircraft.get('registration', 'N-TBD')} · "
                f"{aircraft.get('aftt', 0):,} AFTT"
            ),

            # Key credit metrics (shown to lenders)
            "loanAmount": analysis["transaction"]["loanAmount"],
            "ltv": analysis["transaction"]["ltv"],
            "rating": analysis["riskRating"]["rating"],
            "band": analysis["riskRating"]["band"],
            "disposition": analysis["riskRating"]["disposition"],
            "gdscr": analysis["gdscr"]["gdscr"],
            "nwCoverage": analysis["balanceSheet"]["netWorthCoverage"],
            "flagCount": len(analysis["flags"]),
            "criticalFlags": len([f for f in analysis["flags"] if f.get("severity") == "CRITICAL"]),

            # Full analysis (for lender portal detail view)
            "analysis": analysis,

            # Metadata
            "transactionType": data.get("transactionType", "purchase"),
        }

        save_deal(deal)

        return {
            "success": True,
            "dealId": deal_id,
            "applicationId": deal_id,
            "status": "under_review",
            "rating": analysis["riskRating"]["rating"],
            "band": analysis["riskRating"]["band"],
            "gdscr": analysis["gdscr"]["gdscr"],
            "loanAmount": analysis["transaction"]["loanAmount"],
            "aircraft": deal["aircraft"],
            "receivedDate": analysis["analysisDate"],
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/deals")
def get_deals():
    """
    Return all deals for lender portal queue.
    Returns anonymized summary — borrower identity withheld.
    """
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
                "rating": d.get("rating", "—"),
                "band": d.get("band", ""),
                "disposition": d.get("disposition", ""),
                "gdscr": d.get("gdscr", 0),
                "nwCoverage": d.get("nwCoverage", 0),
                "flagCount": d.get("flagCount", 0),
                "criticalFlags": d.get("criticalFlags", 0),
                "ioiCount": d.get("ioiCount", 0),
                # Analysis included for detail view
                "analysis": d.get("analysis"),
            })
        return {"deals": queue, "count": len(queue)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/deals/{deal_id}")
def get_deal(deal_id: str):
    """Return full deal detail for lender portal."""
    deal = load_deal(deal_id)
    if not deal:
        raise HTTPException(status_code=404, detail=f"Deal {deal_id} not found")
    # Return with anonymized ID
    result = dict(deal)
    result["anonId"] = _anon_id(deal_id)
    # Remove borrower PII from top-level (still in analysis.borrowerName — revealed post-IOI)
    result.pop("borrowerEmail", None)
    return result


@app.patch("/deals/{deal_id}/status")
def patch_status(deal_id: str, update: StatusUpdate):
    """Advance deal stage — called by borrower app demo switcher."""
    valid_statuses = [
        "application_submitted", "under_review", "select_lender_pool",
        "package_distributed", "ioi_received", "lender_selected", "closed"
    ]
    if update.status not in valid_statuses:
        raise HTTPException(status_code=400, detail=f"Invalid status: {update.status}")
    deal = update_deal_status(deal_id, update.status)
    if not deal:
        raise HTTPException(status_code=404, detail=f"Deal {deal_id} not found")
    return {"success": True, "dealId": deal_id, "status": update.status}


@app.post("/deals/{deal_id}/ioi")
def submit_ioi(deal_id: str, ioi: IOISubmission):
    """Lender submits IOI for a deal."""
    deal = load_deal(deal_id)
    if not deal:
        raise HTTPException(status_code=404, detail=f"Deal {deal_id} not found")
    try:
        ioi_data = ioi.dict()
        ioi_id = save_ioi(deal_id, ioi_data)
        return {
            "success": True,
            "ioiId": ioi_id,
            "dealId": deal_id,
            "refNum": f"IOI-{ioi_id}",
            "submittedAt": ioi_data.get("submittedAt", ""),
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/deals/{deal_id}/ioi")
def get_iois(deal_id: str):
    """Return IOIs for a deal — used by borrower dashboard."""
    deal = load_deal(deal_id)
    if not deal:
        raise HTTPException(status_code=404, detail=f"Deal {deal_id} not found")
    iois = load_iois(deal_id)
    return {"dealId": deal_id, "iois": iois, "count": len(iois)}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _anon_id(deal_id: str) -> str:
    """Convert full deal ID to anonymized display ID for lender portal."""
    # AL-2026-ABCD1234 → DEAL-1234
    parts = deal_id.split("-")
    if len(parts) >= 3:
        return f"DEAL-{parts[-1][:4]}"
    return f"DEAL-{deal_id[-4:]}"
