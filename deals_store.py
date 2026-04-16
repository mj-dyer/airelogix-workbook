"""
AireLogix Deal Store — PostgreSQL backend using pg8000 (pure Python, no native deps).
Uses Railway DATABASE_URL env var. Falls back to JSON files if not set.
"""

import json, os, uuid
from datetime import datetime, timezone
from typing import Optional
from urllib.parse import urlparse

DATABASE_URL = os.environ.get("DATABASE_URL")

def _parse_url(url):
    """Parse postgresql://user:pass@host:port/dbname into pg8000 kwargs."""
    p = urlparse(url)
    return {
        "host": p.hostname,
        "port": p.port or 5432,
        "database": p.path.lstrip("/"),
        "user": p.username,
        "password": p.password,
        "ssl_context": True,  # Railway requires SSL
    }

def _get_conn():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL not set")
    import pg8000.native
    params = _parse_url(DATABASE_URL)
    return pg8000.native.Connection(**params)

def _ensure_schema():
    conn = _get_conn()
    try:
        conn.run("""
            CREATE TABLE IF NOT EXISTS deals (
                deal_id      TEXT PRIMARY KEY,
                status       TEXT NOT NULL DEFAULT 'under_review',
                received_date TEXT,
                ioi_count    INTEGER NOT NULL DEFAULT 0,
                data         TEXT NOT NULL,
                created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                updated_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
        """)
        conn.run("""
            CREATE TABLE IF NOT EXISTS iois (
                ioi_id       TEXT PRIMARY KEY,
                deal_id      TEXT NOT NULL,
                institution  TEXT,
                submitted_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                data         TEXT NOT NULL
            )
        """)
        conn.run("CREATE INDEX IF NOT EXISTS idx_deals_created ON deals(created_at DESC)")
        conn.run("CREATE INDEX IF NOT EXISTS idx_iois_deal ON iois(deal_id)")
    finally:
        conn.close()

try:
    _ensure_schema()
    _USE_DB = True
    print("[deals_store] PostgreSQL ready")
except Exception as e:
    print(f"[deals_store] Falling back to file storage: {e}")
    _USE_DB = False

# ── File storage fallback ─────────────────────────────────────────────────────
DEALS_DIR = os.environ.get("DEALS_DIR", "/tmp/airelogix_deals")
IOIS_DIR  = os.environ.get("IOIS_DIR",  "/tmp/airelogix_iois")

def _ensure_dirs():
    os.makedirs(DEALS_DIR, exist_ok=True)
    os.makedirs(IOIS_DIR, exist_ok=True)

# ── ID generation ─────────────────────────────────────────────────────────────
def generate_deal_id() -> str:
    return f"AL-{datetime.now().year}-{str(uuid.uuid4())[:8].upper()}"

# ── Deal operations ───────────────────────────────────────────────────────────
def save_deal(deal: dict) -> str:
    deal_id = deal.get("dealId") or generate_deal_id()
    deal["dealId"] = deal_id
    now = datetime.now(timezone.utc).isoformat()
    deal.setdefault("createdAt", now)
    deal["updatedAt"] = now

    if _USE_DB:
        conn = _get_conn()
        try:
            conn.run("""
                INSERT INTO deals (deal_id, status, received_date, ioi_count, data)
                VALUES (:deal_id, :status, :received_date, :ioi_count, :data)
                ON CONFLICT (deal_id) DO UPDATE SET
                    status=EXCLUDED.status,
                    received_date=EXCLUDED.received_date,
                    ioi_count=EXCLUDED.ioi_count,
                    data=EXCLUDED.data,
                    updated_at=NOW()
            """, deal_id=deal_id,
                status=deal.get("status", "under_review"),
                received_date=deal.get("receivedDate", ""),
                ioi_count=deal.get("ioiCount", 0),
                data=json.dumps(deal))
        finally:
            conn.close()
    else:
        _ensure_dirs()
        with open(os.path.join(DEALS_DIR, f"{deal_id}.json"), "w") as f:
            json.dump(deal, f, indent=2)
    return deal_id

def load_deal(deal_id: str) -> Optional[dict]:
    if _USE_DB:
        conn = _get_conn()
        try:
            rows = conn.run("SELECT data FROM deals WHERE deal_id=:deal_id", deal_id=deal_id)
            return json.loads(rows[0][0]) if rows else None
        finally:
            conn.close()
    else:
        _ensure_dirs()
        p = os.path.join(DEALS_DIR, f"{deal_id}.json")
        return json.load(open(p)) if os.path.exists(p) else None

def list_deals() -> list:
    if _USE_DB:
        conn = _get_conn()
        try:
            rows = conn.run("SELECT data FROM deals ORDER BY created_at DESC LIMIT 200")
            return [json.loads(r[0]) for r in rows]
        finally:
            conn.close()
    else:
        _ensure_dirs()
        deals = []
        for f in sorted(os.listdir(DEALS_DIR), reverse=True):
            if f.endswith(".json"):
                try: deals.append(json.load(open(os.path.join(DEALS_DIR, f))))
                except: pass
        return deals

def update_deal_status(deal_id: str, status: str) -> Optional[dict]:
    deal = load_deal(deal_id)
    if not deal: return None
    deal["status"] = status
    deal["updatedAt"] = datetime.now(timezone.utc).isoformat()
    if _USE_DB:
        conn = _get_conn()
        try:
            conn.run("""
                UPDATE deals SET status=:status, data=:data, updated_at=NOW()
                WHERE deal_id=:deal_id
            """, status=status, data=json.dumps(deal), deal_id=deal_id)
        finally:
            conn.close()
    else:
        save_deal(deal)
    return deal

def update_deal_bio(deal_id: str, borrower_profile: dict) -> None:
    """Save borrowerProfile to an existing deal record without touching other fields."""
    deal = load_deal(deal_id)
    if not deal:
        print(f"[deals_store] update_deal_bio: deal {deal_id} not found")
        return
    deal["borrowerProfile"] = borrower_profile
    deal["updatedAt"] = datetime.now(timezone.utc).isoformat()
    if _USE_DB:
        conn = _get_conn()
        try:
            conn.run(
                "UPDATE deals SET data=:data, updated_at=NOW() WHERE deal_id=:deal_id",
                data=json.dumps(deal), deal_id=deal_id
            )
        finally:
            conn.close()
    else:
        save_deal(deal)


def save_ioi(deal_id: str, ioi: dict) -> str:
    ioi_id = ioi.get("ioiId") or str(uuid.uuid4())[:8].upper()
    ioi["ioiId"] = ioi_id
    ioi["dealId"] = deal_id
    ioi.setdefault("submittedAt", datetime.now(timezone.utc).isoformat())
    institution = ioi.get("institution", "")

    if _USE_DB:
        conn = _get_conn()
        try:
            conn.run("""
                INSERT INTO iois (ioi_id, deal_id, institution, data)
                VALUES (:ioi_id, :deal_id, :institution, :data)
                ON CONFLICT (ioi_id) DO UPDATE SET data=EXCLUDED.data
            """, ioi_id=ioi_id, deal_id=deal_id, institution=institution, data=json.dumps(ioi))
            conn.run("""
                UPDATE deals SET
                    ioi_count=(SELECT COUNT(*) FROM iois WHERE deal_id=:deal_id),
                    status=CASE WHEN status IN ('select_lender_pool','package_distributed','under_review')
                                THEN 'ioi_received' ELSE status END,
                    updated_at=NOW()
                WHERE deal_id=:deal_id
            """, deal_id=deal_id)
        finally:
            conn.close()
    else:
        _ensure_dirs()
        p = os.path.join(IOIS_DIR, f"{deal_id}_iois.json")
        iois = json.load(open(p)) if os.path.exists(p) else []
        idx = next((i for i,x in enumerate(iois) if x.get("institution")==institution), None)
        if idx is not None: iois[idx] = ioi
        else: iois.append(ioi)
        with open(p,"w") as f: json.dump(iois, f, indent=2)
        deal = load_deal(deal_id)
        if deal:
            deal["ioiCount"] = len(iois)
            if deal.get("status") in ("select_lender_pool","package_distributed","under_review"):
                deal["status"] = "ioi_received"
            save_deal(deal)
    return ioi_id

def load_iois(deal_id: str) -> list:
    if _USE_DB:
        conn = _get_conn()
        try:
            rows = conn.run("SELECT data FROM iois WHERE deal_id=:deal_id ORDER BY submitted_at", deal_id=deal_id)
            return [json.loads(r[0]) for r in rows]
        finally:
            conn.close()
    else:
        _ensure_dirs()
        p = os.path.join(IOIS_DIR, f"{deal_id}_iois.json")
        return json.load(open(p)) if os.path.exists(p) else []


# ── Calloway Demo Seed ────────────────────────────────────────────────────────

CALLOWAY_DEAL_ID = "AL-2026-CALLOWAY"

CALLOWAY_DEMO = {
    "dealId": CALLOWAY_DEAL_ID,
    "applicationId": CALLOWAY_DEAL_ID,
    "status": "ioi_received",
    "stage": "ioi_received",
    "ioiCount": 3,
    "receivedDate": "2026-03-28",
    "borrowerName": "James R. Calloway",
    "borrowerEmail": "jcalloway@callowayindustries.com",
    "aircraft": "2016 Bombardier Challenger 350",
    "aircraftSub": "Smart Parts Plus \u00b7 N-registered \u00b7 1,100 AFTT",
    "loanAmount": 15000000,
    "ltv": 0.882,
    "rating": 2,
    "band": "Very Strong",
    "disposition": "Approve",
    "gdscr": 1.8022,
    "nwCoverage": 3.65,
    "flagCount": 3,
    "criticalFlags": 0,
    "transactionType": "purchase",
    "analysis": {
        "borrowerName": "James R. Calloway",
        "analysisDate": "2026-03-28",
        "riskRating": {"rating": 2, "band": "Very Strong", "disposition": "Approve", "composite": 1.667},
        "transaction": {
            "purchasePrice": 17000000,
            "loanAmount": 15000000,
            "ltv": 0.882,
            "downPayment": 2000000,
            "illustrativeRate": 0.063,
            "illustrativeMonthlyPayment": 112239,
            "monthlyPayment": 112239,
            "annualAircraftDS": 1346864,
            "existingAnnualDS": 836858,
            "totalProFormaDS": 2183722,
            "term": 120,
            "termMonths": 120,
            "rateNote": "SOFR OIS + 200bps illustrative spread",
            "primary": {
                "rate": 0.063,
                "monthlyPayment": 112239,
                "term": 120,
            },
        },
        "gdscr": {
            "gdscr": 1.8022,
            "qualifyingIncome": 3935600,
            "governingYear": 2023,
            "totalAnnualDebtService": 2183722,
            "newLoanDebtService": 1346864,
            "existingDebtService": 836858,
        },
        "agi": 4312400,
        "federalTaxesPaid": 1338100,
        "incomeNormalization": {
            "taxYears": [2023, 2024],
            "qualifyingYear": 2023,
            "governingYear": 2023,
            "year1Total": 3935600,
            "year2Total": 4269300,
            "year1Income": 3935600,
            "year2Income": 4269300,
            "governingIncome": 3935600,
            "qualifyingIncome": 3935600,
            "afterTaxQualifyingIncome": 2597496,
            "variancePct": 8.5,
            "varianceFlag": False,
            "taxesPaidLowerYear": 1338100,
            "note": "2023 governs (lower year). LP passive excluded. Real estate net of depreciation.",
            "k1Detail": [
                {"entityName": "Calloway Industrial Group, LLC", "entityType": "LLC", "ownershipPct": 100, "participationType": "Member-Manager", "year1Box1": 2850000, "year2Box1": 3120000, "year1Distributions": 2400000, "year2Distributions": 2600000, "qualifyingIncome": 2850000, "excluded": False},
                {"entityName": "CIG Holdings II, LLC", "entityType": "LLC", "ownershipPct": 85, "participationType": "Member-Manager", "year1Box1": 480000, "year2Box1": 510000, "year1Distributions": 420000, "year2Distributions": 450000, "qualifyingIncome": 480000, "excluded": False},
                {"entityName": "CRW Real Estate Partners I, LLC", "entityType": "LLC", "ownershipPct": 60, "participationType": "Managing Member", "year1Box1": 122200, "year2Box1": 132400, "year1Distributions": 95000, "year2Distributions": 105000, "qualifyingIncome": 122200, "excluded": False, "note": "Net after depreciation"},
                {"entityName": "CRW Real Estate Partners II, LLC", "entityType": "LLC", "ownershipPct": 60, "participationType": "Managing Member", "year1Box1": 76400, "year2Box1": 87900, "year1Distributions": 60000, "year2Distributions": 72000, "qualifyingIncome": 76400, "excluded": False, "note": "Net after depreciation"},
                {"entityName": "Vantage Point Capital Partners, LLC", "entityType": "Partnership", "ownershipPct": 12, "participationType": "Limited Partner", "year1Box1": 185000, "year2Box1": 210000, "year1Distributions": 0, "year2Distributions": 0, "qualifyingIncome": 0, "excluded": True, "note": "LP passive — excluded per methodology"},
            ],
            "interestIncome": 187400,
            "dividendIncome": 124600,
            "otherIncome": 95000,
        },
        "debtDetail": [
            {"creditor": "First Horizon Bank", "accountType": "Primary Residence Mortgage", "balance": 4250000, "monthlyPayment": 28400, "annualPayment": 340800, "note": ""},
            {"creditor": "First Horizon Bank", "accountType": "Vacation Home Mortgage", "balance": 2820000, "monthlyPayment": 18800, "annualPayment": 225600, "note": "Vail, CO property"},
            {"creditor": "Morgan Stanley PWM", "accountType": "Private LOC (Secured)", "balance": 2500000, "monthlyPayment": 10938, "annualPayment": 131250, "note": "Secured by brokerage account; $5M limit"},
            {"creditor": "Morgan Stanley PWM", "accountType": "Margin Loan", "balance": 1850000, "monthlyPayment": 8167, "annualPayment": 98008, "note": "Against Individual Brokerage account"},
            {"creditor": "First Horizon Bank", "accountType": "Commercial RE Loan", "balance": 650000, "monthlyPayment": 5100, "annualPayment": 61200, "note": "CRW Real Estate Partners I — guaranteed by borrower"},
        ],
        "trustStructures": [
            {"trustName": "James R. Calloway Revocable Trust", "trustType": "revocable", "liquidAssets": 3200000, "borrowerHasAccess": True, "note": "Borrower is sole trustee; assets include brokerage and cash. Counted in liquid asset total."},
            {"trustName": "Calloway Family Irrevocable Trust (2019)", "trustType": "irrevocable", "liquidAssets": 1450000, "borrowerHasAccess": False, "note": "Spouse is trustee; borrower has no withdrawal rights. Excluded from liquid asset analysis."},
        ],
        "balanceSheet": {
            "totalAssets": 47798842,
            "grossTotalAssets": 47798842,
            "liquidAssets": 25698842,
            "tier1Assets": 3948382,
            "tier1NetLiquid": 3948382,
            "totalLiabilities": 12070000,
            "netWorth": 35728842,
            "statedNetWorth": 35728842,
            "netWorthCoverage": 3.65,
            "liquidityRatio": 2.62,
            "leverageRatio": 0.338,
            "entityGuaranteeLiquidity": 0,
            "liquidDetail": [
                {"institution": "First Horizon Bank", "accountType": "Checking + MMA", "balance": 3948382, "netAvailable": 3948382},
                {"institution": "Fidelity Investments", "accountType": "Individual Brokerage", "balance": 6100000, "netAvailable": 6100000},
                {"institution": "Morgan Stanley PWM", "accountType": "Margin Account (Pledged)", "balance": 22500460, "marginLoan": 1850000, "pledged": 5000000, "netAvailable": 15650460},
            ],
        },
        "collateral": {
            "year": 2016,
            "make": "Bombardier",
            "model": "Challenger 350",
            "sn": "20632",
            "registration": "N599JF",
            "aftt": 2424,
            "landings": 1457,
            "engine_program": "JSSI Essential Select Plus",
            "apu_enrolled": True,
            "camp_enrolled": True,
            "purchasePrice": 14000000,
            "fmv_curve": 18252000,
            "fmv_purchase_price": 14000000,
            "fmv_used": 18252000,
            "estimatedFMV": 18252000,
            "ltvVsPurchasePrice": 88.2,
            "ltv_on_curve_fmv": 82.2,
            "ltv_on_purchase_price": 88.2,
            "curveName": "AireLogix CL350 Curve v4",
            "collateral_quality_rating": "Strong",
        },
        "aircraft": {
            "description": "2016 Bombardier Challenger 350",
            "serialNumber": "20632",
            "registration": "N599JF",
            "aftt": 2424,
            "landings": 1457,
            "engineProgram": "JSSI Essential Select Plus",
            "year": 2016,
            "make": "Bombardier",
            "model": "Challenger 350",
        },
        "repaymentSources": {
            "primary": "K-1 distributions from operating entities",
            "primaryBasis": "Calloway Industrial Group LLC and CIG Holdings II LLC — 100% ownership, member-manager control",
            "secondary": "Investment income and liquid asset base",
            "tertiary": "Aircraft asset liquidation"
        },
        "flags": [
            {"flag": "Income Concentration", "detail": "Calloway Industrial Group represents 81% of qualifying K-1 income", "severity": "WATCH"},
            {"flag": "LOC Partially Drawn", "detail": "$2.5M drawn on $5.0M private LOC secured by Morgan Stanley account", "severity": "WATCH"},
            {"flag": "Margin Loan Outstanding", "detail": "$1.85M margin loan against Morgan Stanley brokerage account", "severity": "INFO"},
        ],
        "lenderRouting": {
            "recommended": ["US Bank Aviation Finance", "PNC Equipment Finance", "Citizens Private Bank"],
            "rationale": "Very Strong credit — ORR 2 qualifies for all eight lender relationships. Preferred distribution to top three given strong GDSCR (2.53x) and net worth coverage (3.65x).",
        },
    },
}


def seed_calloway():
    """Seed or refresh the Calloway demo deal on every startup."""
    try:
        save_deal(dict(CALLOWAY_DEMO))
        print(f"[deals_store] Calloway demo seeded/refreshed: {CALLOWAY_DEAL_ID}")
    except Exception as e:
        print(f"[deals_store] Calloway seed error: {e}")


# Run seed on import (after schema is ready)
if _USE_DB:
    seed_calloway()
