"""
AireLogix Deal Store
Simple JSON file-based storage for demo purposes.
Stores deals and IOIs as JSON files on Railway's filesystem.
"""

import json
import os
import uuid
from datetime import datetime
from typing import Optional

DEALS_DIR = os.environ.get("DEALS_DIR", "/tmp/airelogix_deals")
IOIS_DIR  = os.environ.get("IOIS_DIR",  "/tmp/airelogix_iois")


def _ensure_dirs():
    os.makedirs(DEALS_DIR, exist_ok=True)
    os.makedirs(IOIS_DIR, exist_ok=True)


def _deal_path(deal_id: str) -> str:
    return os.path.join(DEALS_DIR, f"{deal_id}.json")


def _ioi_path(deal_id: str) -> str:
    return os.path.join(IOIS_DIR, f"{deal_id}_iois.json")


def generate_deal_id() -> str:
    year = datetime.now().year
    short = str(uuid.uuid4())[:8].upper()
    return f"AL-{year}-{short}"


def save_deal(deal: dict) -> str:
    _ensure_dirs()
    deal_id = deal.get("dealId") or generate_deal_id()
    deal["dealId"] = deal_id
    deal["createdAt"] = deal.get("createdAt") or datetime.now().isoformat()
    deal["updatedAt"] = datetime.now().isoformat()
    with open(_deal_path(deal_id), "w") as f:
        json.dump(deal, f, indent=2)
    return deal_id


def load_deal(deal_id: str) -> Optional[dict]:
    _ensure_dirs()
    path = _deal_path(deal_id)
    if not os.path.exists(path):
        return None
    with open(path) as f:
        return json.load(f)


def list_deals() -> list:
    _ensure_dirs()
    deals = []
    for fname in sorted(os.listdir(DEALS_DIR), reverse=True):
        if fname.endswith(".json"):
            try:
                with open(os.path.join(DEALS_DIR, fname)) as f:
                    deal = json.load(f)
                deals.append(deal)
            except Exception:
                continue
    return deals


def update_deal_status(deal_id: str, status: str) -> Optional[dict]:
    deal = load_deal(deal_id)
    if not deal:
        return None
    deal["status"] = status
    deal["updatedAt"] = datetime.now().isoformat()
    save_deal(deal)
    return deal


def save_ioi(deal_id: str, ioi: dict) -> str:
    _ensure_dirs()
    ioi_id = ioi.get("ioiId") or str(uuid.uuid4())[:8].upper()
    ioi["ioiId"] = ioi_id
    ioi["dealId"] = deal_id
    ioi["submittedAt"] = ioi.get("submittedAt") or datetime.now().isoformat()

    # Load existing IOIs for this deal
    path = _ioi_path(deal_id)
    if os.path.exists(path):
        with open(path) as f:
            iois = json.load(f)
    else:
        iois = []

    # Replace if same institution submitted before, else append
    institution = ioi.get("institution", "")
    existing_idx = next((i for i, x in enumerate(iois) if x.get("institution") == institution), None)
    if existing_idx is not None:
        iois[existing_idx] = ioi
    else:
        iois.append(ioi)

    with open(path, "w") as f:
        json.dump(iois, f, indent=2)

    # Update deal IOI count
    deal = load_deal(deal_id)
    if deal:
        deal["ioiCount"] = len(iois)
        if deal.get("status") in ("select_lender_pool", "package_distributed", "under_review"):
            deal["status"] = "ioi_received"
        save_deal(deal)

    return ioi_id


def load_iois(deal_id: str) -> list:
    _ensure_dirs()
    path = _ioi_path(deal_id)
    if not os.path.exists(path):
        return []
    with open(path) as f:
        return json.load(f)
