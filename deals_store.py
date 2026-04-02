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
