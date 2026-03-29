"""
AireLogix Workbook Service
FastAPI microservice — receives analysis JSON, returns openpyxl .xlsx

Deploy to Railway:
    1. Create new Railway project from GitHub repo
    2. Railway auto-detects Python and installs requirements.txt
    3. Set start command: uvicorn main:app --host 0.0.0.0 --port $PORT
    4. Copy the Railway public URL into api/workbook.js as WORKBOOK_SERVICE_URL
"""

import io
import re
import tempfile
import os
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Any

from generate_workbook import generate_workbook

app = FastAPI(title="AireLogix Workbook Service")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://lender.airelogix.com",
        "https://dev.airelogix.com",
        "http://localhost:3000",
    ],
    allow_methods=["POST", "OPTIONS"],
    allow_headers=["Content-Type", "Authorization"],
)


class WorkbookRequest(BaseModel):
    analysis: Any


@app.get("/")
def health():
    return {"status": "ok", "service": "AireLogix Workbook Generator"}


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
