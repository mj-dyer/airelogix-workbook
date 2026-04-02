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
            "status": "under_review",
            "stage": "under_review",
            "ioiCount": 0,
            "receivedDate": analysis["analysisDate"],
            "borrowerName": analysis["borrowerName"],
            "borrowerEmail": personal.get("email", ""),
            "aircraft": (str(aircraft.get("year","")) + " " + str(aircraft.get("make","")) + " " + str(aircraft.get("model",""))).strip(),
            "aircraftSub": str(aircraft.get("engineProgram","No program")) + " / " + str(aircraft.get("registration","N-TBD")) + " / " + str(aircraft.get("aftt","0")) + " AFTT",
            "loanAmount": analysis["transaction"]["loanAmount"],
            "ltv": analysis["transaction"]["ltv"],
            "rating": analysis["riskRating"]["rating"],
            "band": analysis["riskRating"]["band"],
            "disposition": analysis["riskRating"]["disposition"],
            "gdscr": analysis["gdscr"]["gdscr"],
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
            "status": "under_review",
            "rating": analysis["riskRating"]["rating"],
            "band": analysis["riskRating"]["band"],
            "gdscr": analysis["gdscr"]["gdscr"],
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
        "package_distributed", "ioi_received", "lender_selected", "closed"
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
