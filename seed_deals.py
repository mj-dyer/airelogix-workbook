#!/usr/bin/env python3
"""
AireLogix Deal Seeder
Creates diverse test deals in the lender portal without going through the wizard.

Usage:
    python3 seed_deals.py                  # seed 8 fixture deals
    python3 seed_deals.py --clear          # archive all non-pinned deals first, then seed
    python3 seed_deals.py --clear --dry    # show what would be archived, don't run
    python3 seed_deals.py --url URL        # override API base (default: Railway prod)
"""

import argparse, json, sys, time, urllib.request, urllib.error

API_BASE = "https://airelogix-workbook-production.up.railway.app"
PINNED_IDS = {"AL-2026-CALLOWAY"}   # never archived by --clear


# ── HTTP helpers ──────────────────────────────────────────────────────────────

def _req(method, path, body=None):
    url = API_BASE.rstrip("/") + path
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    if data:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as e:
        print(f"  HTTP {e.code}: {e.read().decode()[:200]}")
        return None
    except Exception as e:
        print(f"  Error: {e}")
        return None

def post(path, body):    return _req("POST",   path, body)
def patch(path, body):   return _req("PATCH",  path, body)
def delete(path):        return _req("DELETE", path)
def get(path):           return _req("GET",    path)


# ── Fixture deals ─────────────────────────────────────────────────────────────
# Each fixture is (submission_payload, target_stage, label).
# Submission format matches DealSubmission model in main.py.
# The engine computes FMV, GDSCR/DSCR, flags, lender routing — we just supply inputs.

FIXTURES = [

    # ── 1. Thornton — HNW Strong (GDSCR ~2.1x, clean deal, low LTV) ──────────
    (
        {
            "borrowerType": "individual",
            "personal": {
                "firstName": "William",
                "lastName": "Thornton",
                "email": "wthornton@thorntonventures.com",
                "occupation": "Private Equity Principal",
                "state": "TX",
            },
            "aircraft": {
                "year": 2018,
                "make": "Bombardier",
                "model": "Challenger 350",
                "serial": "20754",
                "registration": "N218WT",
                "aftt": 1650,
                "engineProgram": "Smart Parts Plus",
                "purchasePrice": 20000000,
                "hangarIdent": "Dallas, TX (DAL)",
            },
            "financial": {
                "recurringCash": 2750000,   # targets GDSCR ~2.0x (actual aircraft DS ~$1.06M + $320K existing = ~$1.38M total)
                "liquidAssets": 38000000,
                "totalAssets": 82000000,
                "netWorth": 61000000,
                "existingDebt": 320000,     # existing annual DS
                "agi": 5800000,
                "federalTaxesPaid": 1740000,
                "k1Detail": [
                    {"entityName": "Thornton Capital Partners LP", "entityType": "LP",
                     "ownershipPct": 70, "participationType": "Active",
                     "year1Box1": 2800000, "year2Box1": 3100000,
                     "year1Distributions": 900000, "year2Distributions": 1100000,
                     "qualifyingIncome": 3100000, "excluded": False, "exclusionReason": ""},
                    {"entityName": "Thornton Real Estate Holdings LLC", "entityType": "LLC",
                     "ownershipPct": 100, "participationType": "Active",
                     "year1Box1": 1400000, "year2Box1": 1600000,
                     "year1Distributions": 500000, "year2Distributions": 600000,
                     "qualifyingIncome": 1600000, "excluded": False, "exclusionReason": ""},
                ],
                "liquidDetail": [
                    {"institution": "JP Morgan Private Bank", "accountType": "Investment",
                     "balance": 22000000, "pledged": False, "marginLoanBalance": 0, "note": ""},
                    {"institution": "Goldman Sachs", "accountType": "Brokerage",
                     "balance": 11000000, "pledged": False, "marginLoanBalance": 0, "note": ""},
                    {"institution": "First Republic", "accountType": "Checking/Savings",
                     "balance": 5000000, "pledged": False, "marginLoanBalance": 0, "note": ""},
                ],
                "debtDetail": [
                    {"creditor": "JP Morgan", "accountType": "Residential Mortgage",
                     "balance": 3800000, "monthlyPayment": 22000, "annualPayment": 264000, "note": "Primary residence"},
                    {"creditor": "JP Morgan", "accountType": "Commercial RE",
                     "balance": 2100000, "monthlyPayment": 4667, "annualPayment": 56000, "note": "Office building"},
                ],
                "trustStructures": [
                    {"trustName": "Thornton Family Revocable Trust",
                     "trustType": "revocable", "liquidAssets": 8500000,
                     "borrowerHasAccess": True, "note": "Full discretionary access"},
                ],
            },
            "loanPrefs": {
                "loanAmount": 13000000,
                "preferredTerm": "120",
                "structure": "balloon",
                "priorityRanking": ["rate", "advance", "amortization", "prepayment"],
            },
        },
        "under_review",
        "Thornton (HNW Strong ~2.1x)",
    ),

    # ── 2. Ashford — HNW Adequate (GDSCR ~1.55x, G550, mid-tier) ─────────────
    (
        {
            "borrowerType": "individual",
            "personal": {
                "firstName": "Katherine",
                "lastName": "Ashford",
                "email": "kashford@ashfordgroup.com",
                "occupation": "CEO / Founder",
                "state": "FL",
            },
            "aircraft": {
                "year": 2015,
                "make": "Gulfstream",
                "model": "G550",
                "serial": "5318",
                "registration": "N315KA",
                "aftt": 2850,
                "engineProgram": "Rolls-Royce Corporate Care Enhanced",
                "purchasePrice": 23500000,
                "hangarIdent": "Miami, FL (OPF)",
            },
            "financial": {
                "recurringCash": 2900000,   # targets GDSCR ~1.6x (aircraft DS ~$1.19M + $610K existing = ~$1.80M total)
                "liquidAssets": 21000000,
                "totalAssets": 58000000,
                "netWorth": 44000000,
                "existingDebt": 610000,
                "agi": 4650000,
                "federalTaxesPaid": 1395000,
                "k1Detail": [
                    {"entityName": "Ashford Group Holdings LLC", "entityType": "LLC",
                     "ownershipPct": 100, "participationType": "Active",
                     "year1Box1": 3200000, "year2Box1": 3600000,
                     "year1Distributions": 1200000, "year2Distributions": 1400000,
                     "qualifyingIncome": 3200000, "excluded": False, "exclusionReason": ""},
                ],
                "liquidDetail": [
                    {"institution": "UBS Wealth Management", "accountType": "Investment",
                     "balance": 14000000, "pledged": False, "marginLoanBalance": 0, "note": ""},
                    {"institution": "Wells Fargo Private", "accountType": "Brokerage",
                     "balance": 5500000, "pledged": False, "marginLoanBalance": 800000, "note": "Margin loan outstanding"},
                    {"institution": "Bank of America", "accountType": "Checking",
                     "balance": 1500000, "pledged": False, "marginLoanBalance": 0, "note": ""},
                ],
                "debtDetail": [
                    {"creditor": "Wells Fargo", "accountType": "Residential Mortgage",
                     "balance": 5200000, "monthlyPayment": 28000, "annualPayment": 336000, "note": "Primary — Naples FL"},
                    {"creditor": "Wells Fargo", "accountType": "Residential Mortgage",
                     "balance": 2800000, "monthlyPayment": 15000, "annualPayment": 180000, "note": "Vacation home"},
                    {"creditor": "UBS", "accountType": "Margin Loan",
                     "balance": 800000, "monthlyPayment": 4000, "annualPayment": 48000, "note": ""},
                ],
                "trustStructures": [],
            },
            "loanPrefs": {
                "loanAmount": 16000000,
                "preferredTerm": "120",
                "structure": "balloon",
                "priorityRanking": ["advance", "rate", "prepayment", "amortization"],
            },
        },
        "select_lender_pool",
        "Ashford (HNW Adequate ~1.55x)",
    ),

    # ── 3. Carver — HNW Marginal (GDSCR ~1.15x, NW-led, high LTV, flags) ─────
    (
        {
            "borrowerType": "individual",
            "personal": {
                "firstName": "Marcus",
                "lastName": "Carver",
                "email": "mcarver@carverfamilyoffice.com",
                "occupation": "Real Estate Developer",
                "state": "NV",
            },
            "aircraft": {
                "year": 2017,
                "make": "Bombardier",
                "model": "Challenger 350",
                "serial": "20681",
                "registration": "N717MC",
                "aftt": 3450,
                "engineProgram": "",           # off-program — triggers collateral flag
                "purchasePrice": 15800000,
                "hangarIdent": "Atlanta, GA (PDK)",
            },
            "financial": {
                "recurringCash": 2600000,
                "liquidAssets": 8200000,       # below loan amount — flag
                "totalAssets": 38000000,
                "netWorth": 22000000,
                "existingDebt": 820000,        # elevated existing DS
                "agi": 3100000,
                "federalTaxesPaid": 930000,
                "k1Detail": [
                    {"entityName": "Carver Development Partners LLC", "entityType": "LLC",
                     "ownershipPct": 80, "participationType": "Active",
                     "year1Box1": 1800000, "year2Box1": 1400000,
                     "year1Distributions": 600000, "year2Distributions": 400000,
                     "qualifyingIncome": 1400000, "excluded": False, "exclusionReason": ""},
                    {"entityName": "Carver Capital Group LLC", "entityType": "LLC",
                     "ownershipPct": 60, "participationType": "Passive",
                     "year1Box1": 900000, "year2Box1": 700000,
                     "year1Distributions": 0, "year2Distributions": 0,
                     "qualifyingIncome": 0, "excluded": True,
                     "exclusionReason": "Passive — no qualifying income"},
                ],
                "liquidDetail": [
                    {"institution": "Nevada State Bank", "accountType": "Checking/Savings",
                     "balance": 3200000, "pledged": False, "marginLoanBalance": 0, "note": ""},
                    {"institution": "Schwab", "accountType": "Brokerage",
                     "balance": 5000000, "pledged": False, "marginLoanBalance": 1800000, "note": "Margin loan"},
                ],
                "debtDetail": [
                    {"creditor": "First National Bank", "accountType": "Residential Mortgage",
                     "balance": 4800000, "monthlyPayment": 28000, "annualPayment": 336000, "note": "Primary"},
                    {"creditor": "First National Bank", "accountType": "Commercial RE",
                     "balance": 8200000, "monthlyPayment": 40000, "annualPayment": 480000, "note": "Mixed-use portfolio"},
                ],
                "trustStructures": [],
            },
            "loanPrefs": {
                "loanAmount": 13500000,        # 85% LTV — flag territory
                "preferredTerm": "120",
                "structure": "balloon",
                "priorityRanking": ["advance", "prepayment", "rate", "amortization"],
            },
        },
        "ioi_received",
        "Carver (HNW Marginal ~1.15x)",
    ),

    # ── 4. Hargrove — HNW Weak (GDSCR ~0.88x, cash flow negative, flags) ──────
    (
        {
            "borrowerType": "individual",
            "personal": {
                "firstName": "Derek",
                "lastName": "Hargrove",
                "email": "dhargrove@hargroveinvestments.net",
                "occupation": "Investor / Entrepreneur",
                "state": "CA",
            },
            "aircraft": {
                "year": 2010,
                "make": "Gulfstream",
                "model": "G550",
                "serial": "5182",
                "registration": "N410DH",
                "aftt": 5200,
                "engineProgram": "",           # off-program — critical flag
                "purchasePrice": 17000000,
                "hangarIdent": "Scottsdale, AZ (SDL)",
            },
            "financial": {
                "recurringCash": 2200000,      # income below DS — weak coverage
                "liquidAssets": 6500000,
                "totalAssets": 31000000,
                "netWorth": 14000000,
                "existingDebt": 1100000,       # heavy existing obligations
                "agi": 2800000,
                "federalTaxesPaid": 840000,
                "k1Detail": [
                    {"entityName": "Hargrove Ventures LLC", "entityType": "LLC",
                     "ownershipPct": 100, "participationType": "Active",
                     "year1Box1": 1800000, "year2Box1": 1100000,
                     "year1Distributions": 400000, "year2Distributions": 200000,
                     "qualifyingIncome": 1100000, "excluded": False, "exclusionReason": ""},
                ],
                "liquidDetail": [
                    {"institution": "Chase Private Client", "accountType": "Brokerage",
                     "balance": 4000000, "pledged": True, "marginLoanBalance": 1500000, "note": "Pledged as LOC collateral"},
                    {"institution": "Bank of the West", "accountType": "Checking",
                     "balance": 2500000, "pledged": False, "marginLoanBalance": 0, "note": ""},
                ],
                "debtDetail": [
                    {"creditor": "Chase", "accountType": "Residential Mortgage",
                     "balance": 6800000, "monthlyPayment": 38000, "annualPayment": 456000, "note": "Primary — Malibu"},
                    {"creditor": "Chase", "accountType": "LOC",
                     "balance": 3500000, "monthlyPayment": 17500, "annualPayment": 210000, "note": "Business LOC"},
                    {"creditor": "Chase", "accountType": "Commercial RE",
                     "balance": 5200000, "monthlyPayment": 36200, "annualPayment": 434400, "note": "Investment property"},
                ],
                "trustStructures": [],
            },
            "loanPrefs": {
                "loanAmount": 13500000,
                "preferredTerm": "84",         # shorter term
                "structure": "balloon",
                "priorityRanking": ["advance", "amortization", "rate", "prepayment"],
            },
        },
        "under_review",
        "Hargrove (HNW Weak ~0.88x)",
    ),

    # ── 5. Blackwell Manufacturing — Corporate Strong (DSCR ~2.6x, audited) ───
    (
        {
            "borrowerType": "private_company",
            "personal": {
                "firstName": "Robert",
                "lastName": "Blackwell",
                "email": "rblackwell@blackwellmfg.com",
            },
            "aircraft": {
                "year": 2019,
                "make": "Gulfstream",
                "model": "G550",
                "serial": "5447",
                "registration": "N519BM",
                "aftt": 1420,
                "engineProgram": "Rolls-Royce Corporate Care Enhanced",
                "purchasePrice": 34000000,
                "hangarIdent": "Chicago, IL (MDW)",
            },
            "financial": {
                "entityName": "Blackwell Manufacturing Inc.",
                "financialStatementQuality": "audited",
                "ebitda": {
                    "year1": 7800000,          # FY2024 — targets DSCR ~2.5x (aircraft DS ~$1.56M + $1.1M existing = ~$2.66M)
                    "year2": 6700000,          # FY2023 (governing — lower year)
                    "qualifying": 6700000,
                },
                "revenue": {
                    "year1": 148000000,
                    "year2": 134000000,
                    "qualifying": 134000000,
                },
                "corporateBalanceSheet": {
                    "totalAssets": 185000000,
                    "totalCurrentAssets": 62000000,
                    "totalCurrentLiabilities": 21000000,
                    "totalLiabilities": 78000000,
                    "entityNetWorth": 107000000,
                    "currentRatio": 2.95,
                },
                "debtStack": {
                    "totalDebt": 58000000,
                    "existingAnnualDS": 1100000,
                    "debtToEbitda": 0.71,
                },
                "existingDebt": 1100000,
                "guarantors": [
                    {"name": "Robert Blackwell", "role": "CEO", "ownership": "65%",
                     "netWorth": 48000000, "liquidAssets": 12000000, "guaranteeType": "Full"},
                    {"name": "Susan Blackwell", "role": "CFO", "ownership": "35%",
                     "netWorth": 31000000, "liquidAssets": 7500000, "guaranteeType": "Full"},
                ],
                "taxYears": [2024, 2023],
                "governingYear": 2023,
            },
            "loanPrefs": {
                "loanAmount": 22000000,
                "preferredTerm": "84",
                "structure": "balloon",
                "priorityRanking": ["rate", "advance", "amortization", "prepayment"],
            },
        },
        "select_lender_pool",
        "Blackwell Manufacturing (Corp Strong ~2.6x)",
    ),

    # ── 6. Crestline Partners — Corporate Adequate (DSCR ~1.45x, reviewed) ────
    (
        {
            "borrowerType": "private_company",
            "personal": {
                "firstName": "James",
                "lastName": "Crestline",
                "email": "jmcalister@crestlinepartners.com",
            },
            "aircraft": {
                "year": 2018,
                "make": "Bombardier",
                "model": "Challenger 350",
                "serial": "20738",
                "registration": "N218CP",
                "aftt": 1980,
                "engineProgram": "JSSI Essential Select Plus",
                "purchasePrice": 20500000,
                "hangarIdent": "Nashville, TN (BNA)",
            },
            "financial": {
                "entityName": "Crestline Partners LLC",
                "financialStatementQuality": "reviewed",
                "ebitda": {
                    "year1": 2800000,          # FY2024 — targets DSCR ~1.4x (aircraft DS ~$1.20M + $540K existing = ~$1.74M)
                    "year2": 2450000,          # FY2023 (governing)
                    "qualifying": 2450000,
                },
                "revenue": {
                    "year1": 52000000,
                    "year2": 46500000,
                    "qualifying": 46500000,
                },
                "corporateBalanceSheet": {
                    "totalAssets": 62000000,
                    "totalCurrentAssets": 18500000,
                    "totalCurrentLiabilities": 9200000,
                    "totalLiabilities": 31000000,
                    "entityNetWorth": 31000000,
                    "currentRatio": 2.01,
                },
                "debtStack": {
                    "totalDebt": 24000000,
                    "existingAnnualDS": 540000,
                    "debtToEbitda": 1.46,
                },
                "existingDebt": 540000,
                "guarantors": [
                    {"name": "James McAlister", "role": "Managing Partner", "ownership": "55%",
                     "netWorth": 14500000, "liquidAssets": 3800000, "guaranteeType": "Full"},
                    {"name": "Patricia Voss", "role": "Partner", "ownership": "45%",
                     "netWorth": 9200000, "liquidAssets": 2100000, "guaranteeType": "Full"},
                ],
                "taxYears": [2024, 2023],
                "governingYear": 2023,
            },
            "loanPrefs": {
                "loanAmount": 14000000,
                "preferredTerm": "120",
                "structure": "balloon",
                "priorityRanking": ["rate", "amortization", "advance", "prepayment"],
            },
        },
        "package_distributed",
        "Crestline Partners (Corp Adequate ~1.45x)",
    ),

    # ── 7. Delta Ridge Holdings — Corporate Marginal (DSCR ~1.02x, compiled) ──
    (
        {
            "borrowerType": "private_company",
            "personal": {
                "firstName": "Thomas",
                "lastName": "Ridgeway",
                "email": "tridgeway@deltaridgeholdings.com",
            },
            "aircraft": {
                "year": 2016,
                "make": "Bombardier",
                "model": "Challenger 350",
                "serial": "20623",
                "registration": "N716DR",
                "aftt": 3100,
                "engineProgram": "Smart Parts Plus",
                "purchasePrice": 15200000,
                "hangarIdent": "Houston, TX (HOU)",
            },
            "financial": {
                "entityName": "Delta Ridge Holdings LLC",
                "financialStatementQuality": "compiled",  # flag — lower quality
                "ebitda": {
                    "year1": 1950000,          # FY2024 — targets DSCR ~1.05x (aircraft DS ~$1.04M + $620K existing = ~$1.66M)
                    "year2": 1750000,          # FY2023 (governing)
                    "qualifying": 1750000,
                },
                "revenue": {
                    "year1": 41000000,
                    "year2": 38500000,
                    "qualifying": 38500000,
                },
                "corporateBalanceSheet": {
                    "totalAssets": 38000000,
                    "totalCurrentAssets": 9800000,
                    "totalCurrentLiabilities": 7600000,
                    "totalLiabilities": 24500000,
                    "entityNetWorth": 13500000,
                    "currentRatio": 1.29,
                },
                "debtStack": {
                    "totalDebt": 19000000,
                    "existingAnnualDS": 620000,
                    "debtToEbitda": 2.21,
                },
                "existingDebt": 620000,
                "guarantors": [
                    {"name": "Thomas Ridgeway", "role": "CEO / Owner", "ownership": "100%",
                     "netWorth": 8200000, "liquidAssets": 1900000, "guaranteeType": "Full"},
                ],
                "taxYears": [2024, 2023],
                "governingYear": 2023,
            },
            "loanPrefs": {
                "loanAmount": 12500000,
                "preferredTerm": "120",
                "structure": "balloon",
                "priorityRanking": ["advance", "rate", "amortization", "prepayment"],
            },
        },
        "under_review",
        "Delta Ridge Holdings (Corp Marginal ~1.02x)",
    ),

    # ── 8. Whitmore — HNW, post-IOI, funded (closed deal for archive testing) ─
    (
        {
            "borrowerType": "individual",
            "personal": {
                "firstName": "Elizabeth",
                "lastName": "Whitmore",
                "email": "ewhitmore@whitmorefamily.com",
                "occupation": "Family Office Principal",
                "state": "NY",
            },
            "aircraft": {
                "year": 2020,
                "make": "Bombardier",
                "model": "Challenger 350",
                "serial": "20821",
                "registration": "N820EW",
                "aftt": 880,
                "engineProgram": "Smart Parts Plus",
                "purchasePrice": 22500000,
                "hangarIdent": "Los Angeles, CA (VNY)",
            },
            "financial": {
                "recurringCash": 5900000,
                "liquidAssets": 44000000,
                "totalAssets": 96000000,
                "netWorth": 78000000,
                "existingDebt": 210000,
                "agi": 6400000,
                "federalTaxesPaid": 1920000,
                "k1Detail": [
                    {"entityName": "Whitmore Capital Management LP", "entityType": "LP",
                     "ownershipPct": 80, "participationType": "Active",
                     "year1Box1": 4200000, "year2Box1": 4600000,
                     "year1Distributions": 1500000, "year2Distributions": 1800000,
                     "qualifyingIncome": 4200000, "excluded": False, "exclusionReason": ""},
                    {"entityName": "Whitmore Real Estate Trust LLC", "entityType": "LLC",
                     "ownershipPct": 100, "participationType": "Active",
                     "year1Box1": 980000, "year2Box1": 1100000,
                     "year1Distributions": 300000, "year2Distributions": 350000,
                     "qualifyingIncome": 980000, "excluded": False, "exclusionReason": ""},
                ],
                "liquidDetail": [
                    {"institution": "Morgan Stanley Wealth", "accountType": "Investment",
                     "balance": 28000000, "pledged": False, "marginLoanBalance": 0, "note": ""},
                    {"institution": "Citibank Private Bank", "accountType": "Brokerage",
                     "balance": 12000000, "pledged": False, "marginLoanBalance": 0, "note": ""},
                    {"institution": "JP Morgan", "accountType": "Checking",
                     "balance": 4000000, "pledged": False, "marginLoanBalance": 0, "note": ""},
                ],
                "debtDetail": [
                    {"creditor": "Morgan Stanley", "accountType": "Residential Mortgage",
                     "balance": 2800000, "monthlyPayment": 14000, "annualPayment": 168000, "note": "Primary — Greenwich CT"},
                ],
                "trustStructures": [
                    {"trustName": "Whitmore Family Trust",
                     "trustType": "revocable", "liquidAssets": 12000000,
                     "borrowerHasAccess": True, "note": "Full trustee discretion"},
                ],
            },
            "loanPrefs": {
                "loanAmount": 15000000,
                "preferredTerm": "120",
                "structure": "balloon",
                "priorityRanking": ["rate", "advance", "prepayment", "amortization"],
            },
        },
        "lender_selected",
        "Whitmore (HNW Strong) — lender selected / closing",
    ),
]


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    global API_BASE
    parser = argparse.ArgumentParser(description="AireLogix deal seeder")
    parser.add_argument("--clear", action="store_true",
                        help="Archive all non-pinned deals before seeding")
    parser.add_argument("--dry", action="store_true",
                        help="Dry run — show what would happen, make no changes")
    parser.add_argument("--url", default=API_BASE,
                        help=f"API base URL (default: {API_BASE})")
    args = parser.parse_args()

    API_BASE = args.url.rstrip("/")

    print(f"\nAireLogix Deal Seeder")
    print(f"Target: {API_BASE}")
    print(f"Dry run: {args.dry}\n")

    # Health check
    h = get("/")
    if not h:
        print("ERROR: API unreachable. Check URL and Railway status.")
        sys.exit(1)
    print(f"API OK — {h.get('service','?')} v{h.get('version','?')}\n")

    # ── Clear non-pinned deals ────────────────────────────────────────────────
    if args.clear:
        print("── Clearing non-pinned deals ──────────────────────────────")
        resp = get("/deals") or {}
        deals = resp.get("deals", resp) if isinstance(resp, dict) else resp
        to_archive = [d for d in deals if isinstance(d, dict)
                      and d.get("id") not in PINNED_IDS
                      and d.get("stage") != "archived"]
        if not to_archive:
            print("  Nothing to archive.\n")
        else:
            for d in to_archive:
                did = d.get("id", "?")
                name = d.get("anonId", d.get("id", "?"))
                stage = d.get("stage", "?")
                if args.dry:
                    print(f"  [DRY] Would archive {did} — {name} ({stage})")
                else:
                    res = delete(f"/deals/{did}")
                    ok = res and res.get("success")
                    print(f"  {'✓' if ok else '✗'} Archived {did} — {name}")
            print()

    if args.dry:
        print("── Fixtures that would be created ─────────────────────────")
        for i, (_, stage, label) in enumerate(FIXTURES, 1):
            print(f"  {i}. {label}  →  stage: {stage}")
        print()
        return

    # ── Seed fixture deals ────────────────────────────────────────────────────
    print("── Seeding fixture deals ───────────────────────────────────")
    created = []
    for i, (payload, target_stage, label) in enumerate(FIXTURES, 1):
        print(f"\n  [{i}/{len(FIXTURES)}] {label}")

        res = post("/deals", payload)
        if not res:
            print(f"    ✗ Failed to create deal")
            continue

        deal_id = res.get("dealId") or res.get("deal_id") or res.get("applicationId")
        gdscr = res.get("gdscr") or (res.get("analysis") or {}).get("gdscr", {}).get("gdscr")
        gdscr_str = f"GDSCR/DSCR {gdscr:.2f}x" if gdscr else "GDSCR computed by engine"
        print(f"    ✓ Created {deal_id}  —  {gdscr_str}")

        if target_stage and target_stage != "select_lender_pool":
            time.sleep(0.3)
            patch_res = patch(f"/deals/{deal_id}/status", {"status": target_stage})
            if patch_res:
                print(f"    ✓ Stage → {target_stage}")
            else:
                print(f"    ✗ Stage patch failed (deal still created)")

        created.append((deal_id, label, target_stage))

    print(f"\n── Done ────────────────────────────────────────────────────")
    print(f"  {len(created)}/{len(FIXTURES)} deals created\n")
    for did, label, stage in created:
        print(f"  {did}  {label}  [{stage}]")
    print()


if __name__ == "__main__":
    main()
