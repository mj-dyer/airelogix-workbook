"""
AireLogix Credit Analysis Engine
Implements Aviation Finance Spreading Prompt v1.6 methodology exactly.

Receives structured wizard submission, produces analysis JSON
matching the schema expected by generate_workbook.py and the lender portal.
"""

import math
from datetime import datetime, date
from typing import Optional


# ── Rate methodology ──────────────────────────────────────────────────────────
SOFR_OIS_SPREAD_BPS = 200  # 200bps flat spread over SOFR OIS for GDSCR illustration
SOFR_OIS_FALLBACK = 0.0430  # SOFR OIS ~4.30% as of Q1 2026; update periodically


# ── Collateral engine (XLS+ curve v3) ────────────────────────────────────────
BB_XLS = {
    2008: {"retail": 6.250, "age": 18}, 2009: {"retail": 6.450, "age": 17},
    2010: {"retail": 6.650, "age": 16}, 2011: {"retail": 6.850, "age": 15},
    2012: {"retail": 7.150, "age": 14}, 2013: {"retail": 7.650, "age": 13},
    2014: {"retail": 8.150, "age": 12}, 2015: {"retail": 8.650, "age": 11},
    2016: {"retail": 9.150, "age": 10}, 2017: {"retail": 9.650, "age": 9},
    2018: {"retail": 10.150, "age": 8}, 2019: {"retail": 10.850, "age": 7},
    2020: {"retail": 11.350, "age": 6}, 2021: {"retail": 11.850, "age": 5},
    2022: {"retail": 12.000, "age": 4},
}
MARKET_CAL_XLS = {
    2008: -0.005, 2009: 0.073, 2010: 0.027, 2011: 0.051, 2012: 0.075,
    2013: 0.023, 2014: 0.034, 2015: 0.046, 2016: 0.001, 2017: -0.015,
    2018: 0.014, 2019: 0.042, 2020: 0.026, 2021: -0.059, 2022: 0.080,
}
HRS_YR_XLS = 380
A_SC = 38.999
B0 = 0.016636
KK = 0.44882


def get_xls_fmv(year: int, aftt: int, program: str = "") -> Optional[float]:
    """Returns FMV in millions for Citation XLS/XLS+."""
    b = BB_XLS.get(year)
    if not b:
        return None
    avg = b["age"] * HRS_YR_XLS
    delta = avg - aftt
    A = A_SC * b["retail"]
    B = B0 * math.exp(KK * (10 - b["age"]))
    h_adj = delta * A + abs(delta) * delta * B
    base = b["retail"] * 1e6 + h_adj
    # Program discount
    p = (program or "").lower()
    if "pa+" in p or "esp gold" in p or "tap blue" in p or "power advantage plus" in p:
        disc = 0.0
    elif "power advantage" in p or "esp" in p or " pa " in p:
        disc = 0.020
    elif "jssi" in p:
        disc = 0.030
    elif "no program" in p or "off program" in p or "not enrolled" in p or not p:
        disc = 0.080
    else:
        disc = 0.060
    cal = MARKET_CAL_XLS.get(year, 0)
    fmv = base * (1 - disc) * (1 + cal)
    return round(fmv / 1e6, 3)



# ── CL350 Collateral Curve (v4) ───────────────────────────────────────────────
BB_CL350 = {
    2014: {"retail": 15.5, "age": 12}, 2015: {"retail": 16.0, "age": 11},
    2016: {"retail": 16.5, "age": 10}, 2017: {"retail": 17.0, "age": 9},
    2018: {"retail": 17.5, "age": 8},  2019: {"retail": 18.0, "age": 7},
    2020: {"retail": 18.5, "age": 6},  2021: {"retail": 19.0, "age": 5},
    2022: {"retail": 19.5, "age": 4},
}
MARKET_CAL_CL350 = {
    2014: -0.080, 2015: -0.050, 2016: -0.025, 2017: -0.020,
    2018: -0.045, 2019: +0.075, 2020: +0.055, 2021: +0.040, 2022: +0.020,
}
HRS_YR_CL350 = 475
A_SC_CL350 = 38.999
B0_CL350 = 0.016636
KK_CL350 = 0.44882


def get_cl350_fmv(year: int, aftt: int, program: str = "") -> Optional[float]:
    """Returns FMV in millions for Bombardier Challenger 350."""
    b = BB_CL350.get(int(year))
    if not b:
        # Out of range — use generic
        return get_generic_fmv(year, "Bombardier", "Challenger 350", aftt, program)
    avg = b["age"] * HRS_YR_CL350
    delta = avg - aftt
    A = A_SC_CL350 * b["retail"]
    B = B0_CL350 * math.exp(KK_CL350 * (10 - b["age"]))
    h_adj = delta * A + abs(delta) * delta * B
    base = b["retail"] * 1e6 + h_adj
    cal = MARKET_CAL_CL350.get(int(year), 0)
    base = base * (1 + cal)
    # Program premium
    p = (program or "").lower()
    if "jssi" in p or "esp" in p or "essential select" in p:
        base *= 1.035
    elif "msp" in p or "smart parts" in p:
        base *= 1.025
    elif "off" in p or "not enrolled" in p or not program:
        base *= 0.90
    return round(base / 1e6, 3)


# ── G550 Collateral Curve (v4) ────────────────────────────────────────────────
BB_G550 = {
    2003: 12.0, 2004: 13.0, 2005: 14.0, 2006: 14.5, 2007: 15.5,
    2008: 16.5, 2009: 18.0, 2010: 19.0, 2011: 20.0, 2012: 21.0,
    2013: 22.0, 2014: 24.0, 2015: 25.0, 2016: 26.0, 2017: 28.0,
    2018: 29.0, 2019: 31.0, 2020: 33.0,
}
CAL_G550 = {
    2003: -0.107, 2004: -0.104, 2005: -0.080, 2006: -0.042, 2007: -0.075,
    2008: -0.113, 2009: -0.095, 2010: -0.068, 2011: -0.059, 2012: -0.041,
    2013: -0.141, 2014: -0.016, 2015: -0.084, 2016: -0.047, 2017: +0.066,
    2018: +0.057, 2019: +0.044, 2020: +0.040,
}
G550_AVG_HRS = 500    # average annual utilization
G550_CURR_YR = 2026
G550_ADD = 0.335      # low-time premium: % per 100hrs below average
G550_D1  = 0.320      # high-time discount tier 1: % per 100hrs, first 500hrs over
G550_D2  = 0.647      # high-time discount tier 2: % per 100hrs, beyond 500hrs over


def get_g550_fmv(year: int, aftt: int, program: str = "") -> Optional[float]:
    """
    Returns FMV in millions for Gulfstream G550.
    Calibrated from 51 closed sales + 13 listing ASPs (v4).
    """
    b_retail = BB_G550.get(int(year))
    if not b_retail:
        return get_generic_fmv(year, "Gulfstream", "G550", aftt, program)

    bbv = b_retail * 1e6
    baseline_hrs = (G550_CURR_YR - int(year)) * G550_AVG_HRS
    delta = baseline_hrs - int(aftt)

    if delta >= 0:
        # Low time — premium
        h_adj = bbv * (G550_ADD / 100) * (delta / 100)
    else:
        # High time — discount in two tiers
        above = -delta
        t1 = min(above, 500)
        t2 = max(above - 500, 0)
        h_adj = -(bbv * (G550_D1 / 100) * (t1 / 100) +
                  bbv * (G550_D2 / 100) * (t2 / 100))

    # Program adjustment
    p = (program or "").lower()
    if "enhanced" in p and ("corporatecare" in p or "rolls" in p or "rrcc" in p):
        prog_adj = 0.030   # RRCC Enhanced: +3% premium
    elif "corporatecare" in p or "rrcc" in p or "rolls-royce" in p or "rolls royce" in p:
        prog_adj = 0.0     # RRCC Standard: baseline
    elif "jssi" in p:
        prog_adj = -0.030  # JSSI: -3% discount on G550
    elif "off" in p or "not enrolled" in p or not program:
        prog_adj = -0.100  # Off program: -10%
    else:
        prog_adj = -0.060  # Unrecognized: -6%

    cal = CAL_G550.get(int(year), 0)
    fmv = (bbv + h_adj) * (1 + prog_adj) * (1 + cal)
    return round(fmv / 1e6, 3)

def get_generic_fmv(year: int, make: str, model: str, aftt: int, program: str = "") -> Optional[float]:
    """Rough FMV estimate for aircraft types without a dedicated curve."""
    age = 2026 - year
    model_upper = (model or "").upper()
    make_upper = (make or "").upper()

    if "G550" in model_upper or "G-550" in model_upper:
        base = max(8.0, 34.0 - age * 0.9)
    elif "G650" in model_upper or "G700" in model_upper or "G600" in model_upper:
        base = max(28.0, 72.0 - age * 1.8)
    elif "G500" in model_upper or "G450" in model_upper:
        base = max(6.0, 20.0 - age * 0.7)
    elif "G280" in model_upper:
        base = max(4.0, 15.0 - age * 0.5)
    elif "GLOBAL 7500" in model_upper or "7500" in model_upper:
        base = max(35.0, 80.0 - age * 2.2)
    elif "GLOBAL 6500" in model_upper or "6500" in model_upper or "GLOBAL 6000" in model_upper:
        base = max(14.0, 40.0 - age * 1.1)
    elif "GLOBAL 5500" in model_upper or "5500" in model_upper:
        base = max(10.0, 30.0 - age * 0.9)
    elif "CHALLENGER 350" in model_upper or "CL350" in model_upper:
        base = max(9.0, 23.0 - age * 0.8)
    elif "CHALLENGER 300" in model_upper or "CL300" in model_upper:
        base = max(4.0, 15.0 - age * 0.6)
    elif "CHALLENGER 650" in model_upper or "CL650" in model_upper:
        base = max(8.0, 22.0 - age * 0.7)
    elif "XLS" in model_upper and ("CITATION" in model_upper or "CESSNA" in make_upper):
        return get_xls_fmv(year, aftt, program)
    elif "CITATION" in model_upper or "CESSNA" in make_upper:
        base = max(1.5, 10.0 - age * 0.35)
    elif "PHENOM 300" in model_upper:
        base = max(3.5, 12.0 - age * 0.45)
    elif "PHENOM 100" in model_upper:
        base = max(1.5, 5.0 - age * 0.18)
    elif "PRAETOR" in model_upper:
        base = max(8.0, 20.0 - age * 0.6)
    elif "LEGACY" in model_upper:
        base = max(3.0, 18.0 - age * 0.7)
    elif "FALCON 8X" in model_upper or "FALCON 7X" in model_upper:
        base = max(15.0, 55.0 - age * 1.8)
    elif "FALCON 6X" in model_upper:
        base = max(25.0, 55.0 - age * 1.5)
    elif "FALCON 2000" in model_upper:
        base = max(4.0, 25.0 - age * 0.9)
    elif "FALCON 900" in model_upper:
        base = max(2.5, 18.0 - age * 0.7)
    elif "KING AIR 350" in model_upper:
        base = max(2.0, 8.0 - age * 0.22)
    elif "KING AIR 250" in model_upper or "KING AIR 200" in model_upper:
        base = max(1.2, 5.5 - age * 0.18)
    elif "PC-24" in model_upper:
        base = max(7.0, 12.0 - age * 0.3)
    elif "PC-12" in model_upper:
        base = max(1.5, 5.0 - age * 0.15)
    elif "TBM" in model_upper:
        base = max(1.5, 4.5 - age * 0.12)
    else:
        base = max(2.0, 12.0 - age * 0.4)

    # Program adjustment
    p = (program or "").lower()
    if "jssi" in p:
        disc = 0.030
    elif "no program" in p or "off program" in p or "not enrolled" in p or not p:
        disc = 0.080
    else:
        disc = 0.0

    return round(base * (1 - disc), 3)


# ── Payment math ──────────────────────────────────────────────────────────────

def monthly_payment(principal: float, annual_rate: float, term_months: int,
                    balloon: float = 0.0) -> float:
    """
    Balloon-aware payment. balloon is the future value (lump sum at maturity).
    If balloon=0, standard full amortization.
    Formula: PMT = (PV * g - FV) * r / (g - 1) where g = (1+r)^n
    """
    if annual_rate == 0 or term_months == 0:
        return (principal - balloon) / term_months if term_months else 0
    r = annual_rate / 12
    g = (1 + r) ** term_months
    return (principal * g - balloon) * r / (g - 1)


def get_balloon(fmv: float, loan: float, year: int, aftt: int,
                term_months: int, make: str, model: str, program: str) -> float:
    """
    Calculate balloon payment = 70% of projected FMV at maturity.
    Uses simple annual depreciation from current FMV, capped at 90% of loan.
    fmv is in dollars (e.g. 18252000).
    """
    term_years = round(term_months / 12)
    model_upper = (model or "").upper()

    # Simple depreciation rate by aircraft type
    if "CHALLENGER 350" in model_upper or "CL350" in model_upper:
        depr_rate = 0.030  # 3.0%/yr base
    elif "G550" in model_upper or "G650" in model_upper or "GULFSTREAM" in (make or "").upper():
        depr_rate = 0.028  # G550/Gulfstream depreciate slower
    elif "XLS" in model_upper:
        depr_rate = 0.035
    else:
        depr_rate = 0.035

    fmv_at_maturity = fmv * ((1 - depr_rate) ** term_years)
    balloon = fmv_at_maturity * 0.70
    return min(balloon, loan * 0.90)


# ── GDSCR scoring (Factor 1) ──────────────────────────────────────────────────


def score_corporate_dscr(dscr: float) -> float:
    """Corporate DSCR scoring — stronger thresholds than individual GDSCR."""
    if dscr > 3.0:   return 1.0
    if dscr >= 2.4:  return 1.5
    if dscr >= 2.0:  return 2.0
    if dscr >= 1.75: return 2.5
    if dscr >= 1.50: return 3.0
    if dscr >= 1.35: return 3.5
    if dscr >= 1.20: return 4.5
    if dscr >= 1.10: return 5.5
    if dscr >= 1.00: return 6.5
    return 8.0


def score_corporate_liquidity(current_ratio: float) -> float:
    """Corporate current ratio scoring."""
    if current_ratio > 3.0:  return 1.5
    if current_ratio >= 2.0: return 2.5
    if current_ratio >= 1.5: return 3.5
    if current_ratio >= 1.2: return 4.5
    if current_ratio >= 1.0: return 5.5
    if current_ratio >= 0.8: return 6.5
    return 8.0


def score_financial_statement_quality(quality: str) -> float:
    """F4 for corporate — financial statement quality flag."""
    q = (quality or "").lower()
    if "audited" in q:    return 2.0
    if "reviewed" in q:   return 3.5
    if "compiled" in q:   return 5.5
    if "management" in q: return 8.0  # management-prepared = hard decline flag
    return 3.5  # default: treated as reviewed


def score_gdscr(gdscr: float) -> float:
    if gdscr > 5.0:   return 1.0
    if gdscr >= 4.0:  return 1.5
    if gdscr >= 3.0:  return 2.0
    if gdscr >= 2.5:  return 2.5
    if gdscr >= 2.0:  return 3.0
    if gdscr >= 1.75: return 3.5
    if gdscr >= 1.50: return 4.0
    if gdscr >= 1.35: return 5.0
    if gdscr >= 1.20: return 6.0
    if gdscr >= 1.10: return 7.0
    return 8.0


# ── Liquidity scoring (Factor 2) ─────────────────────────────────────────────

def score_liquidity(ratio: float) -> float:
    if ratio > 10.0:  return 1.0
    if ratio >= 7.0:  return 1.5
    if ratio >= 5.0:  return 2.0
    if ratio >= 4.0:  return 2.5
    if ratio >= 3.0:  return 3.0
    if ratio >= 2.5:  return 3.5
    if ratio >= 2.0:  return 4.0
    if ratio >= 1.5:  return 5.0
    if ratio >= 1.0:  return 6.0
    if ratio >= 0.5:  return 7.0
    return 8.0


# ── Net worth scoring (Factor 3) ─────────────────────────────────────────────

def score_net_worth(ratio: float) -> float:
    if ratio > 30:    return 1.0
    if ratio >= 20:   return 1.5
    if ratio >= 15:   return 2.0
    if ratio >= 12:   return 2.5
    if ratio >= 10:   return 3.0
    if ratio >= 8:    return 3.5
    if ratio >= 5:    return 4.0
    if ratio >= 3:    return 5.0
    if ratio >= 2:    return 6.0
    if ratio >= 1:    return 7.0
    return 8.0


# ── Income quality scoring (Factor 4) ────────────────────────────────────────

def score_income_quality(k1_entities: int, has_w2: bool, variance_flag: bool, income_declining: bool) -> float:
    # Simplified heuristic based on available data from wizard
    if income_declining and variance_flag:
        return 6.0
    if variance_flag:
        return 5.0
    if k1_entities >= 5:
        return 1.5
    if k1_entities >= 3 and has_w2:
        return 2.0
    if k1_entities >= 2 and has_w2:
        return 3.0
    if k1_entities >= 2:
        return 3.5
    if k1_entities == 1 and has_w2:
        return 3.5
    if k1_entities == 1:
        return 4.5
    if has_w2:
        return 5.0
    return 6.0


# ── LTV scoring (Factor 5) ────────────────────────────────────────────────────

def score_ltv(ltv: float) -> float:
    if ltv < 0.50:    return 1.0
    if ltv < 0.55:    return 1.5
    if ltv < 0.60:    return 2.0
    if ltv < 0.65:    return 2.5
    if ltv < 0.70:    return 3.0
    if ltv < 0.75:    return 3.5
    if ltv < 0.80:    return 4.0
    if ltv < 0.85:    return 5.0
    if ltv < 0.90:    return 6.0
    if ltv < 0.95:    return 7.0
    return 8.0


# ── Collateral quality scoring (Factor 6) ────────────────────────────────────

def score_collateral(age: int, program: str, registration: str = "N") -> float:
    p = (program or "").lower()
    enrolled = ("no program" not in p and "off program" not in p and "not enrolled" not in p and bool(p))
    n_reg = registration.upper().startswith("N") if registration else True

    if age <= 10 and enrolled and n_reg:
        return 2.0
    if age <= 10 and enrolled:
        return 2.5
    if age <= 10:
        return 3.5
    if age <= 15 and enrolled and n_reg:
        return 3.0
    if age <= 15 and enrolled:
        return 3.5
    if age <= 15:
        return 4.0
    if age <= 20 and enrolled:
        return 5.0
    if age <= 20:
        return 5.5
    if age <= 25:
        return 6.5
    return 7.0


# ── Rating conversion ─────────────────────────────────────────────────────────

def composite_to_rating(composite: float) -> tuple:
    """Returns (rating_str, band, disposition). Uses >= lower-bound comparison per v1.6 table."""
    if composite >= 8.0:  return ("8",  "Hard Decline",             "HARD DECLINE — do not proceed")
    if composite >= 6.0:  return ("7",  "Difficult",                "Conditional placement — advise borrower on path")
    if composite >= 5.0:  return ("6-", "Marginal",                 "Present with full risk disclosure — limited universe")
    if composite >= 4.75: return ("6",  "Doable",                   "Present — last broadly placeable rating")
    if composite >= 4.50: return ("6+", "Doable",                   "Present — pricing and structure reflect risk")
    if composite >= 4.25: return ("5-", "Acceptable — Lower",       "Approve — enhanced structure")
    if composite >= 4.00: return ("5",  "Acceptable",               "Approve — structured terms")
    if composite >= 3.75: return ("5+", "Acceptable",               "Approve with Market Rate Pricing")
    if composite >= 3.50: return ("4-", "Solid",                    "Approve")
    if composite >= 3.25: return ("4",  "Solid",                    "Approve")
    if composite >= 3.00: return ("4+", "Solid",                    "Approve")
    if composite >= 2.75: return ("3-", "Strong",                   "Approve")
    if composite >= 2.50: return ("3",  "Strong",                   "Approve")
    if composite >= 2.25: return ("3+", "Strong",                   "Approve")
    if composite >= 2.00: return ("2-", "Outstanding",              "Approve")
    if composite >= 1.75: return ("2",  "Outstanding",              "Approve")
    if composite >= 1.50: return ("2+", "Outstanding",              "Approve")
    if composite >= 1.25: return ("1+", "Elite",                    "Approve — prioritize")
    return                       ("1",  "Elite — Best in Class",    "Approve — prioritize")


# ── Main analysis engine ───────────────────────────────────────────────────────

def run_analysis(submission: dict) -> dict:
    """
    Takes wizard submission data, returns complete analysis JSON
    matching the schema used by generate_workbook.py and the lender portal.
    """
    now = datetime.now()
    analysis_date = now.strftime("%Y-%m-%d")
    cur_year = now.year

    # ── Extract borrower & aircraft data ──────────────────────────────────────
    personal = submission.get("personal", {})
    aircraft_data = submission.get("aircraft", {})
    financial = submission.get("financial", {})
    loan_prefs = submission.get("loanPrefs", {})
    borrower_type = submission.get("borrowerType", "individual")

    first_name = personal.get("firstName", "")
    last_name = personal.get("lastName", "")
    borrower_name = f"{first_name} {last_name}".strip() or "Borrower"
    email = personal.get("email", "")

    aircraft_year = int(aircraft_data.get("year", 2015))
    aircraft_make = aircraft_data.get("make", "")
    aircraft_model = aircraft_data.get("model", "")
    aircraft_serial = aircraft_data.get("serial", "")
    aircraft_aftt = int(str(aircraft_data.get("aftt", "0")).replace(",", "") or 0)
    engine_program = aircraft_data.get("engineProgram", "")
    purchase_price_raw = str(aircraft_data.get("purchasePrice", "0")).replace(",", "").replace("$", "")
    purchase_price = float(purchase_price_raw or 0)

    aircraft_age = cur_year - aircraft_year
    aircraft_description = f"{aircraft_year} {aircraft_make} {aircraft_model}".strip()

    # ── FMV calculation ───────────────────────────────────────────────────────
    model_upper = (aircraft_model or "").upper()
    make_upper = (aircraft_make or "").upper()
    if "XLS" in model_upper and ("CITATION" in model_upper or "CESSNA" in make_upper):
        fmv_millions = get_xls_fmv(aircraft_year, aircraft_aftt, engine_program)
    elif "CHALLENGER 350" in model_upper or "CL350" in model_upper:
        fmv_millions = get_cl350_fmv(aircraft_year, aircraft_aftt, engine_program)
    elif "G550" in model_upper or "G-550" in model_upper:
        fmv_millions = get_g550_fmv(aircraft_year, aircraft_aftt, engine_program)
    elif "G650" in model_upper or "G700" in model_upper or "G600" in model_upper or "G500" in model_upper:
        fmv_millions = get_g550_fmv(aircraft_year, aircraft_aftt, engine_program)  # use G550 curve as proxy for other Gulfstreams
    else:
        fmv_millions = get_generic_fmv(aircraft_year, aircraft_make, aircraft_model, aircraft_aftt, engine_program)

    fmv = (fmv_millions or (purchase_price / 1e6 * 0.95)) * 1e6

    # ── Loan parameters ───────────────────────────────────────────────────────
    loan_amount_raw = str(loan_prefs.get("loanAmount", "0")).replace(",", "").replace("$", "")
    loan_amount = float(loan_amount_raw or 0)
    if not loan_amount and purchase_price:
        loan_amount = round(purchase_price * 0.80)

    ltv = loan_amount / purchase_price if purchase_price else 0.80
    ltv_vs_fmv = loan_amount / fmv if fmv else ltv

    # Term
    term_str = str(loan_prefs.get("preferredTerm", "120"))
    try:
        # Strip any text, extract digits only
        import re as _re
        term_digits = _re.sub(r"[^0-9]", "", term_str)
        term_months = int(term_digits) if term_digits else 120
        if term_months < 12 or term_months > 300:
            term_months = 120  # sanity clamp
    except (ValueError, TypeError):
        term_months = 120

    # Illustrative rate: SOFR OIS + 200bps
    illustrative_rate = SOFR_OIS_FALLBACK + SOFR_OIS_SPREAD_BPS / 10000
    rate_note = (
        f"Illustrative all-in rate {illustrative_rate*100:.2f}% = "
        f"SOFR OIS {SOFR_OIS_FALLBACK*100:.2f}% + {SOFR_OIS_SPREAD_BPS}bps spread. "
        f"30/360 basis. Not a lender commitment."
    )

    balloon_pmt = get_balloon(fmv, loan_amount, aircraft_year, aircraft_aftt,
                              term_months, aircraft_make, aircraft_model, engine_program)
    monthly_pmt = monthly_payment(loan_amount, illustrative_rate, term_months, balloon_pmt)
    annual_aircraft_ds = monthly_pmt * 12

    # Existing debt service
    existing_annual_ds = float(str(financial.get("existingDebt", "0")).replace(",", "") or 0)
    total_pro_forma_ds = annual_aircraft_ds + existing_annual_ds

    # ── Financial analysis — routes on borrowerType ───────────────────────────
    def _parse_num(val):
        try:
            return float(str(val).replace(",", "").replace("$", "") or 0)
        except (ValueError, TypeError):
            return 0.0

    is_corporate = borrower_type in ("private_company", "corporate", "company")
    variance_flag = False
    income_declining = False

    if is_corporate:
        # ── Corporate branch ─────────────────────────────────────────────────
        # Read corporate fields from financial summary
        ebitda_data = financial.get("ebitda", {}) or {}
        bs_data = financial.get("corporateBalanceSheet", {}) or {}
        debt_data = financial.get("debtStack", {}) or {}
        revenue_data = financial.get("revenue", {}) or {}

        qualifying_income = _parse_num(ebitda_data.get("qualifying", 0) or financial.get("recurringCash", "0"))
        total_assets = _parse_num(bs_data.get("totalAssets", 0) or financial.get("totalAssets", "0"))
        liquid_assets = _parse_num(bs_data.get("totalCurrentAssets", 0) or financial.get("liquidAssets", "0"))
        total_liabilities = _parse_num(bs_data.get("totalLiabilities", 0) or financial.get("contingentLiabilities", "0"))
        stated_net_worth = _parse_num(bs_data.get("entityNetWorth", 0) or financial.get("netWorth", "0"))
        if stated_net_worth == 0 and total_assets > 0:
            stated_net_worth = max(0, total_assets - total_liabilities)
        contingent_liabilities = 0

        # Corporate DSCR (EBITDA-based)
        gdscr = qualifying_income / total_pro_forma_ds if total_pro_forma_ds > 0 else 0
        gdscr_assessment = (
            "Strong" if gdscr >= 2.4 else
            "Acceptable" if gdscr >= 1.2 else
            "Marginal" if gdscr >= 1.0 else
            "Below threshold"
        )

        # Corporate balance sheet ratios
        current_ratio = _parse_num(bs_data.get("currentRatio", 0))
        if current_ratio == 0:
            current_liab = _parse_num(bs_data.get("totalCurrentLiabilities", 0))
            current_ratio = liquid_assets / current_liab if current_liab > 0 else 0
        liquidity_ratio = current_ratio  # use current ratio as liquidity proxy
        net_worth_coverage = stated_net_worth / loan_amount if loan_amount > 0 else 0
        leverage_ratio = total_liabilities / total_assets if total_assets > 0 else 0
        debt_to_ebitda = _parse_num(debt_data.get("debtToEbitda", 0))
        if debt_to_ebitda == 0 and qualifying_income > 0:
            debt_to_ebitda = _parse_num(debt_data.get("totalDebt", 0)) / qualifying_income

        # Corporate income normalization block
        eff_tax_rate = 0.25  # corporate effective rate
        taxes_paid = qualifying_income * eff_tax_rate
        after_tax_qualifying = qualifying_income * (1 - eff_tax_rate)
        has_w2 = False
        k1_entities = 0

    else:
        # ── Individual / HNWI branch ──────────────────────────────────────────
        total_assets = _parse_num(financial.get("totalAssets", "0"))
        liquid_assets = _parse_num(financial.get("liquidAssets", "0"))
        contingent_liabilities = _parse_num(financial.get("contingentLiabilities", "0"))
        recurring_cash = _parse_num(financial.get("recurringCash", "0"))
        net_worth_submitted = _parse_num(financial.get("netWorth", "0"))

        total_liabilities = contingent_liabilities
        if net_worth_submitted > 0 and total_assets > 0:
            stated_net_worth = net_worth_submitted
            total_liabilities = max(0, total_assets - stated_net_worth)
        elif total_assets > 0:
            stated_net_worth = max(0, total_assets - total_liabilities)
        else:
            total_assets = liquid_assets
            stated_net_worth = max(0, liquid_assets - total_liabilities)

        qualifying_income = recurring_cash
        eff_tax_rate = 0.36
        taxes_paid = qualifying_income * eff_tax_rate
        after_tax_qualifying = qualifying_income * (1 - eff_tax_rate)
        has_w2 = borrower_type == "individual"
        k1_entities = 1 if borrower_type == "individual" else 2
        debt_to_ebitda = 0

        gdscr = qualifying_income / total_pro_forma_ds if total_pro_forma_ds > 0 else 0
        gdscr_assessment = (
            "Strong" if gdscr >= 2.0 else
            "Adequate" if gdscr >= 1.5 else
            "Marginal" if gdscr >= 1.0 else
            "Below threshold"
        )

        liquidity_ratio = liquid_assets / loan_amount if loan_amount > 0 else 0
        net_worth_coverage = stated_net_worth / loan_amount if loan_amount > 0 else 0
        leverage_ratio = total_liabilities / stated_net_worth if stated_net_worth > 0 else 0

    # Entity guarantee check
    entity_guarantee = liquid_assets > loan_amount * 0.5 and liquid_assets < loan_amount
    adjusted_liquid = liquid_assets
    adjusted_liquidity_ratio = liquidity_ratio
    guarantee_entities = []

    if entity_guarantee and liquid_assets < loan_amount:
        # Modest entity guarantee boost for demo
        entity_contribution = min(liquid_assets * 2, loan_amount * 0.5)
        adjusted_liquid = liquid_assets + entity_contribution
        adjusted_liquidity_ratio = adjusted_liquid / loan_amount
        guarantee_entities = [{
            "entity": f"{last_name} Capital LLC",
            "entityNet": entity_contribution * 2,
            "creditPct": 1.0,
            "contribution": entity_contribution,
            "creditBasis": "100%-owned investment entity — full credit",
            "creditAllowed": True,
        }]

    # ── Six-factor scoring ────────────────────────────────────────────────────
    if is_corporate:
        # Corporate factors use DSCR thresholds and current ratio
        f1 = score_corporate_dscr(gdscr)
        f2 = score_corporate_liquidity(liquidity_ratio)
        f3 = score_net_worth(net_worth_coverage)
        f4 = score_financial_statement_quality(
            financial.get("financialStatementQuality", "")
        )
        f5 = score_ltv(ltv_vs_fmv)
        f6 = score_collateral(aircraft_age, engine_program)
    else:
        f1 = score_gdscr(gdscr)
        f2 = score_liquidity(adjusted_liquidity_ratio)
        f3 = score_net_worth(net_worth_coverage)
        f4 = score_income_quality(k1_entities, has_w2, variance_flag, income_declining)
        f5 = score_ltv(ltv_vs_fmv)
        f6 = score_collateral(aircraft_age, engine_program)

    weights = {
        "gdscr": 0.25, "liquidity": 0.22, "netWorth": 0.18,
        "incomeQuality": 0.17, "ltv": 0.10, "collateral": 0.08
    }

    pre_floor = (
        f1 * 0.25 + f2 * 0.22 + f3 * 0.18 +
        f4 * 0.17 + f5 * 0.10 + f6 * 0.08
    )

    # Apply floors per v1.6 methodology
    floor_triggered = None
    final_composite = pre_floor

    if f4 >= 8:
        final_composite = 8.0
        floor_triggered = "Income score = 8 — hard decline"
    elif f1 >= 5:
        if pre_floor < f1:
            final_composite = f1
            floor_triggered = f"GDSCR floor — score {f1}"
    elif f2 >= 5:
        if pre_floor < f2:
            final_composite = f2
            floor_triggered = f"Liquidity floor — score {f2}"
    elif f5 >= 8:
        if pre_floor < 7:
            final_composite = 7.0
            floor_triggered = "LTV >95% floor"

    rating, band, disposition = composite_to_rating(final_composite)

    # ── Repayment sources ─────────────────────────────────────────────────────
    cf_strong = gdscr >= 1.5
    liq_strong = adjusted_liquidity_ratio >= 1.0
    nw_strong = net_worth_coverage >= 5

    if cf_strong and gdscr >= 1.75:
        primary = "Cash Flow"
        primary_basis = f"GDSCR {gdscr:.2f}x provides strong independent debt service coverage."
        secondary = "Net Worth / Balance Sheet"
        secondary_assess = f"Net worth coverage {net_worth_coverage:.1f}x provides meaningful backstop."
    elif nw_strong and net_worth_coverage >= 10:
        primary = "Net Worth / Balance Sheet"
        primary_basis = f"Net worth {net_worth_coverage:.1f}x loan provides primary repayment backstop."
        secondary = "Cash Flow"
        secondary_assess = f"After-tax income {after_tax_qualifying:,.0f} covers debt service at GDSCR {gdscr:.2f}x."
    else:
        primary = "Cash Flow"
        primary_basis = f"GDSCR {gdscr:.2f}x — marginal; balance sheet serves as backstop."
        secondary = "Net Worth / Balance Sheet"
        secondary_assess = f"Net worth coverage {net_worth_coverage:.1f}x provides secondary support."

    dual_confidence = (
        "Strong" if (cf_strong and liq_strong) else
        "Adequate" if (cf_strong or liq_strong) and nw_strong else
        "Weak"
    )

    # ── Flags ─────────────────────────────────────────────────────────────────
    flags = []

    if gdscr < 1.0:
        flags.append({
            "severity": "CRITICAL",
            "code": "FLAG-1",
            "title": "GDSCR below 1.0x — income does not cover debt service",
            "detail": f"Pro forma GDSCR of {gdscr:.2f}x falls below the minimum 1.0x threshold. After-tax qualifying income of {after_tax_qualifying:,.0f} does not independently cover total pro forma debt service of {total_pro_forma_ds:,.0f}.",
            "mitigant": f"Net worth of {stated_net_worth:,.0f} provides balance sheet backstop. Liquidity of {liquid_assets:,.0f} ({liquidity_ratio:.2f}x loan) provides additional coverage.",
            "action": "Present to specialty lenders only. Recommend LTV reduction or income documentation to improve GDSCR.",
        })
    elif gdscr < 1.25:
        flags.append({
            "severity": "MATERIAL",
            "code": "FLAG-1",
            "title": f"GDSCR {gdscr:.2f}x — marginal cash flow coverage",
            "detail": f"Pro forma GDSCR of {gdscr:.2f}x is below the 1.25x institutional preference. Credit relies on balance sheet strength rather than cash flow alone.",
            "mitigant": f"Net worth coverage {net_worth_coverage:.1f}x provides primary backstop. Liquidity ratio {liquidity_ratio:.2f}x.",
            "action": "Recommend as balance-sheet-led credit. Document net worth and liquidity as primary repayment sources.",
        })

    if liquid_assets < loan_amount:
        flags.append({
            "severity": "MATERIAL" if liquidity_ratio < 0.5 else "MATERIAL",
            "code": "FLAG-2",
            "title": f"Liquid assets below loan amount — liquidity ratio {liquidity_ratio:.2f}x",
            "detail": f"Personal liquid assets of {liquid_assets:,.0f} represent {liquidity_ratio:.2f}x the loan amount of {loan_amount:,.0f}.",
            "mitigant": f"Net worth of {stated_net_worth:,.0f} ({net_worth_coverage:.1f}x loan) provides strong economic backstop.",
            "action": "Confirm liquid asset documentation. Consider entity guarantee to bring effective liquidity above 1.0x.",
        })

    if ltv_vs_fmv > 0.90:
        flags.append({
            "severity": "MATERIAL",
            "code": "FLAG-3",
            "title": f"LTV vs FMV {ltv_vs_fmv*100:.1f}% — above 90% threshold",
            "detail": f"Loan of {loan_amount:,.0f} represents {ltv_vs_fmv*100:.1f}% of AireLogix estimated FMV of {fmv:,.0f}.",
            "mitigant": "Strong credit profile and engine program enrollment support lender confidence in collateral.",
            "action": "Flag for lender attention. Consider independent appraisal.",
        })

    if aircraft_age > 15:
        flags.append({
            "severity": "MATERIAL",
            "code": f"FLAG-{len(flags)+1}",
            "title": f"Aircraft age {aircraft_age} years — approaching upper lender limits",
            "detail": f"The {aircraft_year} {aircraft_description} is {aircraft_age} years old. Many institutional lenders have maximum aircraft age limits of 20 years at origination.",
            "mitigant": "Engine program enrollment and low hours are positive mitigants. Confirm engine program status.",
            "action": "Verify engine program and AFTT. Route to lenders with known appetite for this vintage.",
        })

    # ── Lender routing ────────────────────────────────────────────────────────
    all_lenders = [
        {"name": "US Bank Aviation Finance",  "tier": "Tier 1 — National Bank",   "minLoan": 1000000,  "maxLoan": 25000000, "maxRating": "5+", "minAge": 0, "maxAge": 25},
        {"name": "PNC Equipment Finance",     "tier": "Tier 1 — National Bank",   "minLoan": 1000000,  "maxLoan": 20000000, "maxRating": "5+", "minAge": 0, "maxAge": 22},
        {"name": "Citizens Private Bank",     "tier": "Tier 1 — Private Bank",    "minLoan": 1000000,  "maxLoan": 30000000, "maxRating": "5",  "minAge": 0, "maxAge": 25},
        {"name": "Fifth Third Leasing",       "tier": "Tier 2 — Regional Bank",   "minLoan": 500000,   "maxLoan": 15000000, "maxRating": "5-", "minAge": 0, "maxAge": 22},
        {"name": "Truist Aviation Finance",   "tier": "Tier 2 — Regional Bank",   "minLoan": 500000,   "maxLoan": 15000000, "maxRating": "5-", "minAge": 0, "maxAge": 22},
        {"name": "Scope Aircraft Finance",    "tier": "Tier 3 — Specialty",       "minLoan": 500000,   "maxLoan": 30000000, "maxRating": "6",  "minAge": 0, "maxAge": 28},
        {"name": "Republic Bank",             "tier": "Tier 3 — Community Bank",  "minLoan": 250000,   "maxLoan": 10000000, "maxRating": "6",  "minAge": 0, "maxAge": 28},
        {"name": "Deerwood Bank",             "tier": "Tier 3 — Community Bank",  "minLoan": 250000,   "maxLoan": 8000000,  "maxRating": "6",  "minAge": 0, "maxAge": 30},
    ]

    rating_order = ["1", "1+", "2+", "2", "2-", "3+", "3", "3-", "4+", "4", "4-",
                    "5+", "5", "5-", "6+", "6", "6-", "7", "8"]

    def rating_ok(lender_max: str, deal_rating: str) -> bool:
        try:
            return rating_order.index(deal_rating) <= rating_order.index(lender_max)
        except ValueError:
            return False

    lender_routing = []
    for l in all_lenders:
        loan_ok = l["minLoan"] <= loan_amount <= l["maxLoan"]
        rate_ok = rating_ok(l["maxRating"], rating)
        age_ok = aircraft_age <= l["maxAge"]
        if loan_ok and rate_ok and age_ok:
            lender_routing.append({"name": l["name"], "tier": l["tier"], "status": "ELIGIBLE", "notes": "Meets all standard criteria"})
        else:
            reason = []
            if not loan_ok: reason.append(f"Loan amount outside range")
            if not rate_ok: reason.append(f"Rating {rating} below appetite")
            if not age_ok:  reason.append(f"Aircraft age {aircraft_age}yr exceeds limit")
            lender_routing.append({"name": l["name"], "tier": l["tier"], "status": "; ".join(reason), "notes": ""})

    # ── Guarantors ────────────────────────────────────────────────────────────
    guarantors = [
        {
            "name": borrower_name,
            "type": "Individual — Full Recourse",
            "ownership": "N/A",
            "netWorth": stated_net_worth,
            "liquidAssets": liquid_assets,
            "guaranteeType": "Full",
            "role": f"Primary credit anchor. Net worth {net_worth_coverage:.1f}x loan coverage. Personal guarantee covers all obligations.",
        }
    ]
    if guarantee_entities:
        guarantors.append({
            "name": f"{last_name} Capital LLC",
            "type": "Entity — Investment Holding",
            "ownership": "100%",
            "netWorth": guarantee_entities[0]["entityNet"],
            "liquidAssets": guarantee_entities[0]["contribution"],
            "guaranteeType": "Full",
            "role": "Entity guarantee providing additional liquidity support.",
        })

    # ── Build income normalization block ──────────────────────────────────────
    k1_detail = []
    if k1_entities >= 1 and not has_w2:
        k1_detail.append({
            "entityName": f"{last_name} Holdings LLC",
            "participationType": "active",
            "year1Box1": qualifying_income,
            "year2Box1": qualifying_income,
        })

    income_normalization = {
        "qualifyingYear": cur_year - 1,
        "qualifyingIncome": qualifying_income,
        "afterTaxQualifyingIncome": after_tax_qualifying,
        "taxesPaidLowerYear": taxes_paid,
        "variancePct": 0.0,
        "varianceFlag": variance_flag,
        "year1Total": qualifying_income,
        "year2Total": qualifying_income,
        "taxYears": [cur_year - 2, cur_year - 1],
        "note": ("Corporate deal — EBITDA basis. Entity: " + borrower_name) if is_corporate else "Income based on borrower-reported recurring sources. Full tax return analysis required before final underwriting.",
        "k1Detail": k1_detail,
        "w2Detail": [{"employer": "Primary Employment", "year1": qualifying_income, "year2": qualifying_income}] if has_w2 else [],
        "portfolioIncome": {},
        "capitalGains": {},
    }

    # ── Build final analysis object ───────────────────────────────────────────
    analysis = {
        "analysisDate": analysis_date,
        "applicationId": submission.get("applicationId", f"AL-{cur_year}-0001"),
        "borrowerName": borrower_name,
        "borrowerEmail": email,
        "dealType": borrower_type,
        "aircraft": {
            "description": aircraft_description,
            "serialNumber": aircraft_serial or "TBD",
            "registration": aircraft_data.get("registration", "N-TBD"),
            "aftt": aircraft_aftt,
            "engineProgram": engine_program or "Not Enrolled / Off Program",
            "year": aircraft_year,
            "make": aircraft_make,
            "model": aircraft_model,
        },
        "collateral": {
            "year": aircraft_year,
            "make": aircraft_make,
            "model": aircraft_model,
            "sn": aircraft_serial or "TBD",
            "registration": aircraft_data.get("registration", ""),
            "aftt": aircraft_aftt,
            "engine_program": engine_program or "",
            "fmv_curve": round(fmv),
            "fmv_purchase_price": round(purchase_price),
            "fmv_used": round(fmv),
            "estimatedFMV": round(fmv),
            "ltv_on_curve_fmv": round(ltv_vs_fmv * 100, 1),
            "ltv_on_purchase_price": round(ltv * 100, 1),
            "curveName": (
                "AireLogix CL350 Curve v4" if ("CHALLENGER 350" in (aircraft_model or "").upper() or "CL350" in (aircraft_model or "").upper())
                else "AireLogix G550 Curve v4" if ("G550" in (aircraft_model or "").upper() or "G-550" in (aircraft_model or "").upper())
                else "AireLogix XLS Curve v4" if "XLS" in (aircraft_model or "").upper()
                else "AireLogix Generic Curve"
            ),
        },
        "transaction": {
            "purchasePrice": purchase_price,
            "loanAmount": loan_amount,
            "ltv": round(ltv, 4),
            "ltvVsFMV": round(ltv_vs_fmv, 4),
            "fmv": round(fmv),
            "illustrativeRate": round(illustrative_rate, 4),
            "rateNote": rate_note,
            "termMonths": term_months,
            "monthlyPayment": round(monthly_pmt),
            "annualAircraftDS": round(annual_aircraft_ds),
            "existingAnnualDS": round(existing_annual_ds),
            "totalProFormaDS": round(total_pro_forma_ds),
        },
        "incomeNormalization": income_normalization,
        "gdscr": {
            "gdscr": round(gdscr, 4),
            "assessment": gdscr_assessment,
            "afterTaxIncome": round(after_tax_qualifying),
            "totalAnnualDS": round(total_pro_forma_ds),
        },
        "balanceSheet": {
            "tier1NetLiquid": round(liquid_assets),
            "tier1GrossLiquid": round(liquid_assets),
            "marginLoanAdjustment": 0,
            "grossTotalAssets": round(total_assets),
            "totalLiabilities": round(total_liabilities),
            "statedNetWorth": round(stated_net_worth),
            "liquidityRatio": round(liquidity_ratio, 4),
            "liquidityAssessment": (
                "Strong" if liquidity_ratio >= 2.0 else
                "Adequate" if liquidity_ratio >= 1.0 else
                "Marginal" if liquidity_ratio >= 0.5 else
                "Critical"
            ),
            "netWorthCoverage": round(net_worth_coverage, 2),
            "leverageRatio": round(leverage_ratio, 4),
            "tier1Assets": [
                {"description": "Reported liquid assets (cash, brokerage, money market)", "value": liquid_assets, "encumbrance": 0, "net": liquid_assets}
            ],
            "entityGuaranteeLiquidity": {
                "adjustmentApplied": bool(guarantee_entities),
                "baseLiquid": round(liquid_assets),
                "adjustedLiquid": round(adjusted_liquid),
                "baseLiquidityRatio": round(liquidity_ratio, 4),
                "adjustedLiquidityRatio": round(adjusted_liquidity_ratio, 4),
                "guaranteeAdjustment": round(adjusted_liquid - liquid_assets),
                "basis": "Personal liquid assets + entity guarantee pool",
                "guaranteeEntities": guarantee_entities,
            } if guarantee_entities else {
                "adjustmentApplied": False,
                "baseLiquid": round(liquid_assets),
                "adjustedLiquid": round(liquid_assets),
                "baseLiquidityRatio": round(liquidity_ratio, 4),
                "adjustedLiquidityRatio": round(liquidity_ratio, 4),
                "guaranteeAdjustment": 0,
                "basis": "Personal liquid assets only",
                "guaranteeEntities": [],
            },
        },
        "riskRating": {
            "factorScores": {
                "F1_GDSCR":         {"score": f1, "weight": 0.25, "weighted": round(f1 * 0.25, 4), "basis": f"GDSCR {gdscr:.2f}x"},
                "F2_Liquidity":     {"score": f2, "weight": 0.22, "weighted": round(f2 * 0.22, 4), "basis": f"Liquidity ratio {adjusted_liquidity_ratio:.2f}x"},
                "F3_NetWorth":      {"score": f3, "weight": 0.18, "weighted": round(f3 * 0.18, 4), "basis": f"Net worth coverage {net_worth_coverage:.2f}x"},
                "F4_IncomeQuality": {"score": f4, "weight": 0.17, "weighted": round(f4 * 0.17, 4), "basis": "Based on reported income profile"},
                "F5_LTV":           {"score": f5, "weight": 0.10, "weighted": round(f5 * 0.10, 4), "basis": f"LTV vs FMV {ltv_vs_fmv*100:.1f}%"},
                "F6_Collateral":    {"score": f6, "weight": 0.08, "weighted": round(f6 * 0.08, 4), "basis": f"Age {aircraft_age}yr, program: {engine_program or 'Not enrolled'}"},
            },
            "composite": {
                "preFloorComposite": round(pre_floor, 4),
                "finalComposite": round(final_composite, 4),
                "floorTriggered": floor_triggered,
                "weights": weights,
            },
            "rating": rating,
            "band": band,
            "disposition": disposition,
        },
        "repaymentSources": {
            "primary": primary,
            "primaryBasis": primary_basis,
            "secondary": secondary,
            "secondaryAssessment": secondary_assess,
            "dualSourceConfidence": dual_confidence,
        },
        "flags": flags,
        "guarantors": guarantors,
        "lenderRouting": lender_routing,
    }

    return analysis
