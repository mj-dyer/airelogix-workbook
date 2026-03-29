"""
AireLogix Underwriter Workbook Generator
Produces a 7-tab branded .xlsx from run_spreading_analysis() JSON output.

Usage:
    python generate_workbook.py <analysis_json_path> <output_path>
    python generate_workbook.py hargrove_analysis.json hargrove_workbook.xlsx

Or import directly:
    from generate_workbook import generate_workbook
    generate_workbook(analysis_dict, "./output/workbook.xlsx")

Tabs:
    1. Cover          — deal summary, rating, disposition
    2. Income         — normalization, variance, qualifying income build
    3. Balance Sheet  — tiered assets, entity guarantee, coverage metrics
    4. GDSCR          — pro forma debt service build, coverage calculation
    5. Scorecard      — six-factor scoring with weighted composite
    6. Lenders        — eligible/eliminated routing table
    7. Flags          — risk flags with mitigants and required actions
"""

import json
import sys
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ── Brand colors (openpyxl uses ARGB hex, no #) ──────────────────────────────
NAVY_DEEP   = "FF060F1E"
NAVY_BASE   = "FF071428"
NAVY_MID    = "FF0B1E3A"
NAVY_CARD   = "FF0E2244"
NAVY_BORDER = "FF1A3460"
GOLD        = "FFC8A96E"
GOLD_LIGHT  = "FFE8CC9A"
GOLD_DARK   = "FFA06810"
CREAM       = "FFF5F3EE"
STEEL       = "FF8BA4BE"
WHITE       = "FFFFFFFF"
GREEN       = "FF2D9E5F"
AMBER       = "FFD4800A"
RED         = "FFC0392B"
BLACK       = "FF000000"
BLUE_INPUT  = "FF0000FF"   # industry standard — hardcoded inputs
BLACK_CALC  = "FF000000"   # industry standard — formulas
GREEN_LINK  = "FF008000"   # industry standard — cross-sheet links

# ── Fonts ─────────────────────────────────────────────────────────────────────
FONT_BODY    = "Arial"
FONT_MONO    = "Courier New"

# ── Number formats ────────────────────────────────────────────────────────────
FMT_CURRENCY  = '$#,##0;($#,##0);"-"'
FMT_CURRENCY2 = '$#,##0.00;($#,##0.00);"-"'
FMT_PCT       = '0.00%;(0.00%);"-"'
FMT_PCT1      = '0.0%;(0.0%);"-"'
FMT_MULTIPLE  = '0.00x'
FMT_NUMBER    = '#,##0;(#,##0);"-"'
FMT_TEXT      = '@'

# ── Helpers ───────────────────────────────────────────────────────────────────

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color=BLACK, size=10, italic=False, name=FONT_BODY):
    return Font(name=name, bold=bold, color=color, size=size, italic=italic)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border(color=NAVY_BORDER):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _bottom_border(color=GOLD_DARK):
    return Border(bottom=Side(style="medium", color=color))

def safe(v, fallback="—"):
    if v is None or v == "":
        return fallback
    return v

def fmt_currency(v):
    if v is None: return "—"
    return v  # let Excel format it

def fmt_pct(v):
    if v is None: return "—"
    return v  # let Excel format it

def set_col_width(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width

def apply_header_row(ws, row, values, widths=None, bg=NAVY_BASE, fg=GOLD_LIGHT):
    """Write a styled header row."""
    for ci, val in enumerate(values, 1):
        c = ws.cell(row=row, column=ci, value=val)
        c.font = _font(bold=True, color=fg, size=10)
        c.fill = _fill(bg)
        c.alignment = _align("center")
        c.border = _border()

def apply_section_title(ws, row, col, text, colspan=1, bg=NAVY_MID, fg=GOLD):
    """Section heading spanning multiple columns."""
    c = ws.cell(row=row, column=col, value=text)
    c.font = _font(bold=True, color=fg, size=11)
    c.fill = _fill(bg)
    c.alignment = _align("left")
    c.border = _bottom_border()
    if colspan > 1:
        ws.merge_cells(
            start_row=row, start_column=col,
            end_row=row, end_column=col + colspan - 1
        )

def apply_kv_row(ws, row, label, value, fmt=None, shade=False, value_color=BLACK_CALC, label_col=1, value_col=2):
    bg = "FFE8EFF5" if shade else "FFFAFBFC"
    lc = ws.cell(row=row, column=label_col, value=label)
    lc.font = _font(bold=True, color="FF374151", size=10)
    lc.fill = _fill(bg)
    lc.alignment = _align("left")
    lc.border = _border("FFD1D5DB")

    vc = ws.cell(row=row, column=value_col, value=value)
    vc.font = _font(color=value_color, size=10, name=FONT_MONO)
    vc.fill = _fill(bg)
    vc.alignment = _align("right")
    vc.border = _border("FFD1D5DB")
    if fmt:
        vc.number_format = fmt
    return lc, vc

def data_row(ws, row, values, fmts=None, shade=False, colors=None):
    bg = "FFE8EFF5" if shade else "FFFFFFFF"
    for ci, val in enumerate(values, 1):
        c = ws.cell(row=row, column=ci, value=val)
        color = colors[ci-1] if colors and ci-1 < len(colors) else BLACK_CALC
        if not color or len(str(color)) < 6:
            color = BLACK_CALC
        c.font = _font(color=color, size=10)
        c.fill = _fill(bg)
        c.alignment = _align("right" if ci > 1 else "left")
        c.border = _border("FFD1D5DB")
        if fmts and ci-1 < len(fmts) and fmts[ci-1]:
            c.number_format = fmts[ci-1]

def spacer_row(ws, row, ncols=8):
    for ci in range(1, ncols+1):
        c = ws.cell(row=row, column=ci, value=None)
        c.fill = _fill("FFFFFFFF")

def total_row(ws, row, values, fmts=None):
    for ci, val in enumerate(values, 1):
        c = ws.cell(row=row, column=ci, value=val)
        c.font = _font(bold=True, color=WHITE, size=10)
        c.fill = _fill(NAVY_MID)
        c.alignment = _align("right" if ci > 1 else "left")
        c.border = _border(NAVY_BORDER)
        if fmts and ci-1 < len(fmts) and fmts[ci-1]:
            c.number_format = fmts[ci-1]


# ── Tab 1: Cover ──────────────────────────────────────────────────────────────

def build_cover(wb, a):
    ws = wb.active
    ws.title = "Cover"
    ws.sheet_view.showGridLines = False

    # Column widths
    for i, w in enumerate([2, 28, 22, 18, 18, 2], 1):
        set_col_width(ws, i, w)

    r = 1
    # Title band
    ws.merge_cells(f"B{r}:E{r}")
    c = ws.cell(row=r, column=2, value="AIRELOGIX — UNDERWRITER WORKBOOK")
    c.font = _font(bold=True, color=GOLD_LIGHT, size=14)
    c.fill = _fill(NAVY_BASE)
    c.alignment = _align("left", "center")
    ws.row_dimensions[r].height = 32
    r += 1

    ws.merge_cells(f"B{r}:E{r}")
    c = ws.cell(row=r, column=2, value="CONFIDENTIAL — INTERNAL USE ONLY")
    c.font = _font(bold=False, color=STEEL, size=9, italic=True)
    c.fill = _fill(NAVY_BASE)
    c.alignment = _align("left", "center")
    ws.row_dimensions[r].height = 18
    r += 2

    # Deal summary block
    apply_section_title(ws, r, 2, "DEAL SUMMARY", colspan=4)
    ws.row_dimensions[r].height = 22
    r += 1

    summary_rows = [
        ("Analysis Date",       safe(a.get("analysisDate")),         FMT_TEXT,     BLUE_INPUT),
        ("Application ID",      safe(a.get("applicationId")),        FMT_TEXT,     BLUE_INPUT),
        ("Borrower",            safe(a.get("borrowerName")),          FMT_TEXT,     BLUE_INPUT),
        ("Aircraft",            safe(a["aircraft"]["description"]),   FMT_TEXT,     BLUE_INPUT),
        ("Serial Number",       safe(a["aircraft"]["serialNumber"]),  FMT_TEXT,     BLUE_INPUT),
        ("Registration",        safe(a["aircraft"]["registration"]),  FMT_TEXT,     BLUE_INPUT),
        ("AFTT (hours)",        a["aircraft"].get("aftt"),            FMT_NUMBER,   BLUE_INPUT),
        ("Engine Program",      safe(a["aircraft"]["engineProgram"]), FMT_TEXT,     BLUE_INPUT),
    ]
    for i, (label, val, fmt, vc) in enumerate(summary_rows):
        apply_kv_row(ws, r, label, val, fmt=fmt, shade=i%2==1, value_color=vc, label_col=2, value_col=3)
        # Merge value across remaining cols
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # Rating block
    apply_section_title(ws, r, 2, "RISK RATING", colspan=4)
    ws.row_dimensions[r].height = 22
    r += 1

    rr = a["riskRating"]
    rating_rows = [
        ("Composite Score",  rr["composite"]["finalComposite"],  "0.0000",    BLACK_CALC),
        ("AireLogix Rating", safe(rr["rating"]),                  FMT_TEXT,    BLUE_INPUT),
        ("Rating Band",      safe(rr["band"]),                    FMT_TEXT,    BLACK_CALC),
        ("Disposition",      safe(rr["disposition"]),             FMT_TEXT,    BLACK_CALC),
        ("Floor Applied",    safe(rr["composite"].get("floorTriggered"), "None"), FMT_TEXT, BLACK_CALC),
    ]
    for i, (label, val, fmt, vc) in enumerate(rating_rows):
        apply_kv_row(ws, r, label, val, fmt=fmt, shade=i%2==1, value_color=vc, label_col=2, value_col=3)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # Transaction block
    apply_section_title(ws, r, 2, "TRANSACTION PARAMETERS", colspan=4)
    ws.row_dimensions[r].height = 22
    r += 1

    tx = a["transaction"]
    tx_rows = [
        ("Purchase Price",           tx["purchasePrice"],      FMT_CURRENCY,  BLUE_INPUT),
        ("Requested Loan Amount",    tx["loanAmount"],          FMT_CURRENCY,  BLUE_INPUT),
        ("Loan-to-Value",            tx["ltv"],                 FMT_PCT,       BLACK_CALC),
        ("Proposed Term (months)",   tx["termMonths"],          FMT_NUMBER,    BLUE_INPUT),
        ("Illustrative Rate",        tx["illustrativeRate"],    FMT_PCT,       BLUE_INPUT),
        ("Monthly P&I (Illus.)",     tx["monthlyPayment"],      FMT_CURRENCY,  BLACK_CALC),
        ("Annual Aircraft DS",       tx["annualAircraftDS"],    FMT_CURRENCY,  BLACK_CALC),
        ("Existing Annual DS",       tx["existingAnnualDS"],    FMT_CURRENCY,  BLUE_INPUT),
        ("Total Pro Forma DS",       tx["totalProFormaDS"],     FMT_CURRENCY,  BLACK_CALC),
    ]
    for i, (label, val, fmt, vc) in enumerate(tx_rows):
        apply_kv_row(ws, r, label, val, fmt=fmt, shade=i%2==1, value_color=vc, label_col=2, value_col=3)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # Rate note
    ws.merge_cells(f"B{r}:E{r}")
    c = ws.cell(row=r, column=2, value=safe(tx.get("rateNote")))
    c.font = _font(italic=True, color=STEEL, size=8)
    c.alignment = _align("left", "center", wrap=True)
    ws.row_dimensions[r].height = 28

    # Freeze top row
    ws.freeze_panes = "B3"


# ── Tab 2: Income ─────────────────────────────────────────────────────────────

def build_income(wb, a):
    ws = wb.create_sheet("Income")
    ws.sheet_view.showGridLines = False

    inc = a["incomeNormalization"]
    years = inc.get("taxYears", ["Year 1", "Year 2"])
    y1, y2 = str(years[0]), str(years[1])

    for i, w in enumerate([2, 32, 16, 16, 20, 2], 1):
        set_col_width(ws, i, w)

    r = 1
    ws.merge_cells(f"B{r}:E{r}")
    c = ws.cell(row=r, column=2, value="INCOME NORMALIZATION")
    c.font = _font(bold=True, color=GOLD_LIGHT, size=13)
    c.fill = _fill(NAVY_BASE)
    c.alignment = _align("left", "center")
    ws.row_dimensions[r].height = 28
    r += 2

    # Income source table
    apply_section_title(ws, r, 2, "INCOME BY SOURCE", colspan=4)
    ws.row_dimensions[r].height = 22
    r += 1

    apply_header_row(ws, r, ["", "Income Source", y1, y2, "Treatment"])
    ws.row_dimensions[r].height = 18
    r += 1

    sources = []
    for w2 in inc.get("w2Detail", []):
        sources.append(("W-2", safe(w2.get("employer","W-2 Income")), w2.get("year1",0), w2.get("year2",0), "Qualifying"))
    for k1 in inc.get("k1Detail", []):
        y1v = k1.get("year1Box1", 0) or 0
        y2v = k1.get("year2Box1", 0) or 0
        treatment = "Loss — excluded" if y1v < 0 or y2v < 0 else f"Qualifying ({k1.get('participationType','')})"
        sources.append(("K-1", safe(k1.get("entityName","")), y1v, y2v, treatment))
    pi = inc.get("portfolioIncome", {})
    if pi.get("year1") or pi.get("year2"):
        sources.append(("Portfolio", "Interest / Dividends", pi.get("year1",0), pi.get("year2",0), "Qualifying"))
    cg = inc.get("capitalGains", {})
    if cg.get("year1") or cg.get("year2"):
        sources.append(("Cap Gains", "Capital Gains", cg.get("year1",0), cg.get("year2",0), "Excluded — non-recurring"))

    for i, (stype, name, v1, v2, treatment) in enumerate(sources):
        shade = i % 2 == 1
        bg = "FFE8EFF5" if shade else "FFFFFFFF"
        for ci, (val, fmt, col) in enumerate([
            (stype, FMT_TEXT, "FF6B7280"),
            (name,  FMT_TEXT, BLACK_CALC),
            (v1,    FMT_CURRENCY, BLUE_INPUT),
            (v2,    FMT_CURRENCY, BLUE_INPUT),
            (treatment, FMT_TEXT, BLACK_CALC),
        ], 1):
            c = ws.cell(row=r, column=ci+1, value=val)
            c.font = _font(color=col, size=10)
            c.fill = _fill(bg)
            c.alignment = _align("left" if ci <= 2 else "right")
            c.border = _border("FFD1D5DB")
            c.number_format = fmt
        ws.row_dimensions[r].height = 17
        r += 1

    # Totals
    total_row(ws, r, ["", "TOTAL QUALIFYING INCOME", inc.get("year1Total",0), inc.get("year2Total",0), ""],
              fmts=[None, None, FMT_CURRENCY, FMT_CURRENCY, None])
    ws.row_dimensions[r].height = 20
    r += 2

    # Normalization summary
    apply_section_title(ws, r, 2, "NORMALIZATION SUMMARY — LOWER YEAR METHODOLOGY", colspan=4)
    ws.row_dimensions[r].height = 22
    r += 1

    norm_rows = [
        ("Qualifying Year (Lower Year Governs)",  inc.get("qualifyingYear"),               FMT_NUMBER,  BLUE_INPUT),
        (f"{y1} Total Qualifying Income",          inc.get("year1Total"),                  FMT_CURRENCY, BLUE_INPUT),
        (f"{y2} Total Qualifying Income",          inc.get("year2Total"),                  FMT_CURRENCY, BLUE_INPUT),
        ("Income Variance ($)",                    inc.get("variance"),                    FMT_CURRENCY, BLACK_CALC),
        ("Income Variance (%)",                    inc.get("variancePct"),                 FMT_PCT,      BLACK_CALC),
        ("Variance Flag (>20%)",                   "YES ⚠" if inc.get("varianceFlag") else "No", FMT_TEXT, RED if inc.get("varianceFlag") else GREEN),
        ("Qualifying Income (Gross)",              inc.get("qualifyingIncome"),            FMT_CURRENCY, BLACK_CALC),
        ("Taxes Paid — Qualifying Year",           inc.get("taxesPaidLowerYear"),          FMT_CURRENCY, BLUE_INPUT),
        ("After-Tax Qualifying Income",            inc.get("afterTaxQualifyingIncome"),    FMT_CURRENCY, BLACK_CALC),
    ]
    for i, (label, val, fmt, vc) in enumerate(norm_rows):
        lc, vc_cell = apply_kv_row(ws, r, label, val, fmt=fmt, shade=i%2==1, value_color=vc, label_col=2, value_col=4)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1
    ws.merge_cells(f"B{r}:E{r}")
    c = ws.cell(row=r, column=2, value=safe(inc.get("note")))
    c.font = _font(italic=True, color=STEEL, size=8)
    c.alignment = _align("left", wrap=True)
    ws.row_dimensions[r].height = 22

    ws.freeze_panes = "B2"


# ── Tab 3: Balance Sheet ──────────────────────────────────────────────────────

def build_balance_sheet(wb, a):
    ws = wb.create_sheet("Balance Sheet")
    ws.sheet_view.showGridLines = False

    bs = a["balanceSheet"]
    gl = bs.get("entityGuaranteeLiquidity", {})

    for i, w in enumerate([2, 32, 18, 16, 16, 2], 1):
        set_col_width(ws, i, w)

    r = 1
    ws.merge_cells(f"B{r}:E{r}")
    c = ws.cell(row=r, column=2, value="BALANCE SHEET ANALYSIS")
    c.font = _font(bold=True, color=GOLD_LIGHT, size=13)
    c.fill = _fill(NAVY_BASE)
    c.alignment = _align("left", "center")
    ws.row_dimensions[r].height = 28
    r += 2

    # Tier 1 liquid assets
    apply_section_title(ws, r, 2, "TIER 1 — PERSONAL LIQUID ASSETS", colspan=4)
    ws.row_dimensions[r].height = 22
    r += 1
    apply_header_row(ws, r, ["", "Asset Description", "Gross Value", "Encumbrance", "Net Value"])
    ws.row_dimensions[r].height = 18
    r += 1

    for i, asset in enumerate(bs.get("tier1Assets", [])):
        gross = asset.get("value") or asset.get("gross") or 0
        enc   = asset.get("encumbrance", 0) or 0
        net   = asset.get("net") if asset.get("net") is not None else gross - enc
        data_row(ws, r,
            ["", safe(asset.get("description") or asset.get("institution")), gross, enc, net],
            fmts=[None, None, FMT_CURRENCY, FMT_CURRENCY, FMT_CURRENCY],
            shade=i%2==1,
            colors=["", BLACK_CALC, BLUE_INPUT, BLUE_INPUT, BLACK_CALC]
        )
        ws.row_dimensions[r].height = 17
        r += 1

    # Margin loans
    for ml in bs.get("marginLoans", []):
        data_row(ws, r,
            ["", f"Margin Loan — {safe(ml.get('institution',''))}", 0, ml.get("balance",0), -ml.get("balance",0)],
            fmts=[None, None, FMT_CURRENCY, FMT_CURRENCY, FMT_CURRENCY],
            shade=True,
            colors=["", RED, None, RED, RED]
        )
        ws.row_dimensions[r].height = 17
        r += 1

    total_row(ws, r, ["", "NET PERSONAL LIQUID (POST-MARGIN)", "", "", bs.get("tier1NetLiquid",0)],
              fmts=[None, None, None, None, FMT_CURRENCY])
    ws.row_dimensions[r].height = 20
    r += 2

    # Entity guarantee adjustment
    if gl.get("adjustmentApplied"):
        apply_section_title(ws, r, 2, "ENTITY GUARANTEE — LIQUIDITY ADJUSTMENT", colspan=4)
        ws.row_dimensions[r].height = 22
        r += 1
        apply_header_row(ws, r, ["", "Entity", "Net Assets", "Credit %", "Contribution"])
        ws.row_dimensions[r].height = 18
        r += 1

        for i, ent in enumerate(gl.get("guaranteeEntities", [])):
            if ent.get("creditAllowed"):
                data_row(ws, r,
                    ["", safe(ent.get("entity")), ent.get("entityNet",0), ent.get("creditPct",0), ent.get("contribution",0)],
                    fmts=[None, None, FMT_CURRENCY, FMT_PCT1, FMT_CURRENCY],
                    shade=i%2==1,
                    colors=["", BLACK_CALC, BLUE_INPUT, BLUE_INPUT, BLACK_CALC]
                )
                ws.row_dimensions[r].height = 17
                r += 1

        total_row(ws, r, ["", "EFFECTIVE LIQUID (ADJUSTED)", gl.get("adjustedLiquid",0), "", ""],
                  fmts=[None, None, FMT_CURRENCY, None, None])
        ws.row_dimensions[r].height = 20
        r += 2

    # Balance sheet summary
    apply_section_title(ws, r, 2, "BALANCE SHEET SUMMARY", colspan=4)
    ws.row_dimensions[r].height = 22
    r += 1

    bs_rows = [
        ("Gross Total Assets",          bs.get("grossTotalAssets"),     FMT_CURRENCY, BLUE_INPUT),
        ("Total Liabilities",           bs.get("totalLiabilities"),     FMT_CURRENCY, BLUE_INPUT),
        ("Stated Net Worth",            bs.get("statedNetWorth"),       FMT_CURRENCY, BLACK_CALC),
        ("Leverage Ratio",              bs.get("leverageRatio"),        FMT_MULTIPLE, BLACK_CALC),
    ]
    for i, (label, val, fmt, vc) in enumerate(bs_rows):
        lc, vc_cell = apply_kv_row(ws, r, label, val, fmt=fmt, shade=i%2==1, value_color=vc, label_col=2, value_col=4)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # Coverage metrics
    apply_section_title(ws, r, 2, "COVERAGE METRICS", colspan=4)
    ws.row_dimensions[r].height = 22
    r += 1

    loan = a["transaction"]["loanAmount"]
    cov_rows = [
        ("Loan Amount (Denominator)",            loan,                           FMT_CURRENCY, GREEN_LINK),
        ("Personal Liquid (Numerator)",          bs.get("tier1NetLiquid"),       FMT_CURRENCY, GREEN_LINK),
        ("Personal Liquidity Ratio",             bs.get("liquidityRatio"),       FMT_MULTIPLE, BLACK_CALC),
        ("Effective Liquid (w/ Guarantee)",      gl.get("adjustedLiquid") if gl.get("adjustmentApplied") else bs.get("tier1NetLiquid"), FMT_CURRENCY, GREEN_LINK),
        ("Adjusted Liquidity Ratio (F2 Input)",  gl.get("adjustedLiquidityRatio") if gl.get("adjustmentApplied") else bs.get("liquidityRatio"), FMT_MULTIPLE, BLACK_CALC),
        ("Net Worth (Numerator)",                bs.get("statedNetWorth"),       FMT_CURRENCY, GREEN_LINK),
        ("Net Worth Coverage",                   bs.get("netWorthCoverage"),     FMT_MULTIPLE, BLACK_CALC),
    ]
    for i, (label, val, fmt, vc) in enumerate(cov_rows):
        lc, vc_cell = apply_kv_row(ws, r, label, val, fmt=fmt, shade=i%2==1, value_color=vc, label_col=2, value_col=4)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 18
        r += 1

    ws.freeze_panes = "B2"


# ── Tab 4: GDSCR ──────────────────────────────────────────────────────────────

def build_gdscr(wb, a):
    ws = wb.create_sheet("GDSCR")
    ws.sheet_view.showGridLines = False

    tx  = a["transaction"]
    inc = a["incomeNormalization"]
    g   = a["gdscr"]

    for i, w in enumerate([2, 34, 22, 16, 2], 1):
        set_col_width(ws, i, w)

    r = 1
    ws.merge_cells(f"B{r}:D{r}")
    c = ws.cell(row=r, column=2, value="GLOBAL DEBT SERVICE COVERAGE RATIO")
    c.font = _font(bold=True, color=GOLD_LIGHT, size=13)
    c.fill = _fill(NAVY_BASE)
    c.alignment = _align("left", "center")
    ws.row_dimensions[r].height = 28
    r += 2

    # Income
    apply_section_title(ws, r, 2, "QUALIFYING INCOME", colspan=3)
    ws.row_dimensions[r].height = 22
    r += 1

    inc_rows = [
        ("Qualifying Year",                    inc.get("qualifyingYear"),                  FMT_NUMBER,   BLUE_INPUT),
        ("Gross Qualifying Income",            inc.get("qualifyingIncome"),                FMT_CURRENCY, GREEN_LINK),
        ("Taxes Paid — Qualifying Year",       inc.get("taxesPaidLowerYear"),              FMT_CURRENCY, GREEN_LINK),
        ("After-Tax Qualifying Income",        inc.get("afterTaxQualifyingIncome"),        FMT_CURRENCY, BLACK_CALC),
    ]
    for i2, (label, val, fmt, vc) in enumerate(inc_rows):
        lc, vc_cell = apply_kv_row(ws, r, label, val, fmt=fmt, shade=i2%2==1, value_color=vc, label_col=2, value_col=3)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # Debt service build
    apply_section_title(ws, r, 2, "PRO FORMA DEBT SERVICE BUILD", colspan=3)
    ws.row_dimensions[r].height = 22
    r += 1

    ds_rows = [
        ("Loan Amount",                        tx.get("loanAmount"),           FMT_CURRENCY, BLUE_INPUT),
        ("Illustrative Rate (All-In)",         tx.get("illustrativeRate"),     FMT_PCT,      BLUE_INPUT),
        ("Term (months)",                      tx.get("termMonths"),           FMT_NUMBER,   BLUE_INPUT),
        ("Monthly P&I Payment (Illustrative)", tx.get("monthlyPayment"),       FMT_CURRENCY, BLACK_CALC),
        ("Annual Aircraft Debt Service",       tx.get("annualAircraftDS"),     FMT_CURRENCY, BLACK_CALC),
        ("Existing Annual Debt Service",       tx.get("existingAnnualDS"),     FMT_CURRENCY, BLUE_INPUT),
        ("Total Pro Forma Annual DS",          tx.get("totalProFormaDS"),      FMT_CURRENCY, BLACK_CALC),
    ]
    for i2, (label, val, fmt, vc) in enumerate(ds_rows):
        lc, vc_cell = apply_kv_row(ws, r, label, val, fmt=fmt, shade=i2%2==1, value_color=vc, label_col=2, value_col=3)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # GDSCR result
    apply_section_title(ws, r, 2, "GDSCR RESULT", colspan=3)
    ws.row_dimensions[r].height = 22
    r += 1

    gdscr_rows = [
        ("After-Tax Qualifying Income",  g.get("afterTaxIncome"),    FMT_CURRENCY, GREEN_LINK),
        ("Total Pro Forma Annual DS",    g.get("totalAnnualDS"),     FMT_CURRENCY, GREEN_LINK),
        ("Pro Forma GDSCR",              g.get("gdscr"),             FMT_MULTIPLE, BLACK_CALC),
        ("Assessment",                   safe(g.get("assessment")),  FMT_TEXT,     BLACK_CALC),
    ]
    for i2, (label, val, fmt, vc) in enumerate(gdscr_rows):
        color = vc
        if label == "Pro Forma GDSCR":
            gdscr_val = g.get("gdscr") or 0
            color = GREEN if gdscr_val >= 1.5 else AMBER if gdscr_val >= 1.0 else RED
        lc, vc_cell = apply_kv_row(ws, r, label, val, fmt=fmt, shade=i2%2==1, value_color=color, label_col=2, value_col=3)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 2
    ws.merge_cells(f"B{r}:D{r}")
    c = ws.cell(row=r, column=2, value=safe(tx.get("rateNote")))
    c.font = _font(italic=True, color=STEEL, size=8)
    c.alignment = _align("left", wrap=True)
    ws.row_dimensions[r].height = 28

    ws.freeze_panes = "B2"


# ── Tab 5: Scorecard ──────────────────────────────────────────────────────────

def build_scorecard(wb, a):
    ws = wb.create_sheet("Scorecard")
    ws.sheet_view.showGridLines = False

    rr = a["riskRating"]
    fs = rr["factorScores"]

    for i, w in enumerate([2, 26, 10, 10, 12, 28, 2], 1):
        set_col_width(ws, i, w)

    r = 1
    ws.merge_cells(f"B{r}:F{r}")
    c = ws.cell(row=r, column=2, value="RISK RATING SCORECARD")
    c.font = _font(bold=True, color=GOLD_LIGHT, size=13)
    c.fill = _fill(NAVY_BASE)
    c.alignment = _align("left", "center")
    ws.row_dimensions[r].height = 28
    r += 2

    apply_section_title(ws, r, 2, "SIX-FACTOR SCORING MODEL", colspan=5)
    ws.row_dimensions[r].height = 22
    r += 1

    apply_header_row(ws, r, ["", "Factor", "Weight", "Score /8", "Weighted", "Basis"])
    ws.row_dimensions[r].height = 18
    r += 1

    factors = [
        ("F1_GDSCR",         "F1 — GDSCR",           0.25),
        ("F2_Liquidity",     "F2 — Liquidity",        0.22),
        ("F3_NetWorth",      "F3 — Net Worth",        0.18),
        ("F4_IncomeQuality", "F4 — Income Quality",   0.17),
        ("F5_LTV",           "F5 — LTV",              0.10),
        ("F6_Collateral",    "F6 — Collateral",       0.08),
    ]

    for i, (key, label, weight) in enumerate(factors):
        factor = fs.get(key, {})
        score  = factor.get("score")
        weighted = factor.get("weighted")
        basis  = safe(factor.get("basis"))
        shade  = i % 2 == 1
        bg     = "FFE8EFF5" if shade else "FFFFFFFF"

        score_color = GREEN if score and score <= 2 else \
                      GOLD  if score and score <= 3.5 else \
                      AMBER if score and score <= 5.5 else RED

        for ci, (val, fmt, col, align) in enumerate([
            ("",       None,         BLACK_CALC,  "left"),
            (label,    FMT_TEXT,     BLACK_CALC,  "left"),
            (weight,   FMT_PCT1,     BLACK_CALC,  "center"),
            (score,    "0.0",        score_color, "center"),
            (weighted, "0.0000",     BLACK_CALC,  "center"),
            (basis,    FMT_TEXT,     "FF6B7280",  "left"),
        ], 1):
            c = ws.cell(row=r, column=ci, value=val)
            c.font = _font(color=col, size=10, bold=(ci==2))
            c.fill = _fill(bg)
            c.alignment = _align(align)
            c.border = _border("FFD1D5DB")
            if fmt: c.number_format = fmt
        ws.row_dimensions[r].height = 18
        r += 1

    # Composite total row
    comp = rr["composite"]
    for ci, (val, fmt, col) in enumerate([
        ("",                      None,     WHITE),
        ("COMPOSITE SCORE",       None,     WHITE),
        (comp.get("preFloorComposite"), "0.0000", GOLD_LIGHT),
        ("→",                     None,     GOLD_LIGHT),
        (comp.get("finalComposite"),    "0.0000", GOLD_LIGHT),
        (f"Rating {rr['rating']} — {rr['band']}", None, GOLD_LIGHT),
    ], 1):
        c = ws.cell(row=r, column=ci, value=val)
        c.font = _font(bold=True, color=col, size=11)
        c.fill = _fill(NAVY_MID)
        c.alignment = _align("center" if ci > 2 else "left")
        c.border = _border(NAVY_BORDER)
        if fmt: c.number_format = fmt
    ws.row_dimensions[r].height = 24
    r += 2

    # Disposition
    apply_section_title(ws, r, 2, "DISPOSITION", colspan=5)
    ws.row_dimensions[r].height = 22
    r += 1

    disp_rows = [
        ("Rating",       safe(rr["rating"]),       FMT_TEXT, BLACK_CALC),
        ("Band",         safe(rr["band"]),          FMT_TEXT, BLACK_CALC),
        ("Disposition",  safe(rr["disposition"]),   FMT_TEXT, GREEN),
        ("Floor Rule",   safe(comp.get("floorTriggered"), "None applied"), FMT_TEXT, BLACK_CALC),
    ]
    for i2, (label, val, fmt, vc) in enumerate(disp_rows):
        lc, vc_cell = apply_kv_row(ws, r, label, val, fmt=fmt, shade=i2%2==1, value_color=vc, label_col=2, value_col=4)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # Repayment sources
    apply_section_title(ws, r, 2, "REPAYMENT SOURCES", colspan=5)
    ws.row_dimensions[r].height = 22
    r += 1

    rep = a.get("repaymentSources", {})
    rep_rows = [
        ("Primary Source",            safe(rep.get("primary")),            FMT_TEXT, BLACK_CALC),
        ("Primary Basis",             safe(rep.get("primaryBasis")),       FMT_TEXT, BLACK_CALC),
        ("Secondary Source",          safe(rep.get("secondary")),          FMT_TEXT, BLACK_CALC),
        ("Secondary Assessment",      safe(rep.get("secondaryAssessment")),FMT_TEXT, BLACK_CALC),
        ("Dual-Source Confidence",    safe(rep.get("dualSourceConfidence")),FMT_TEXT,
            GREEN if rep.get("dualSourceConfidence") == "Strong" else AMBER if rep.get("dualSourceConfidence") == "Adequate" else RED),
    ]
    for i2, (label, val, fmt, vc) in enumerate(rep_rows):
        lc, vc_cell = apply_kv_row(ws, r, label, val, fmt=fmt, shade=i2%2==1, value_color=vc, label_col=2, value_col=4)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        vc_cell.alignment = _align("left", wrap=True)
        ws.row_dimensions[r].height = 18
        r += 1

    ws.freeze_panes = "B2"


# ── Tab 6: Lenders ────────────────────────────────────────────────────────────

def build_lenders(wb, a):
    ws = wb.create_sheet("Lenders")
    ws.sheet_view.showGridLines = False

    rr = a["riskRating"]
    lenders = a.get("lenderRouting", [])
    eligible   = [l for l in lenders if l.get("status") == "ELIGIBLE"]
    eliminated = [l for l in lenders if l.get("status") != "ELIGIBLE"]

    for i, w in enumerate([2, 28, 20, 36, 2], 1):
        set_col_width(ws, i, w)

    r = 1
    ws.merge_cells(f"B{r}:D{r}")
    c = ws.cell(row=r, column=2, value=f"LENDER ROUTING — Rating {rr['rating']} ({rr['band']})")
    c.font = _font(bold=True, color=GOLD_LIGHT, size=13)
    c.fill = _fill(NAVY_BASE)
    c.alignment = _align("left", "center")
    ws.row_dimensions[r].height = 28
    r += 2

    # Eligible
    apply_section_title(ws, r, 2, f"ELIGIBLE LENDERS ({len(eligible)})", colspan=3)
    ws.row_dimensions[r].height = 22
    r += 1
    apply_header_row(ws, r, ["", "Institution", "Tier", "Notes"])
    ws.row_dimensions[r].height = 18
    r += 1

    for i, l in enumerate(eligible):
        shade = i % 2 == 1
        bg = "FFE8EFF5" if shade else "FFFFFFFF"
        for ci, (val, col) in enumerate([
            ("",                  BLACK_CALC),
            (l.get("name",""),    BLACK_CALC),
            (l.get("tier",""),    "FF6B7280"),
            (l.get("notes",""),   "FF6B7280"),
        ], 1):
            c = ws.cell(row=r, column=ci, value=val)
            c.font = _font(color=col, size=10, bold=(ci==2))
            c.fill = _fill(bg)
            c.alignment = _align("left")
            c.border = _border("FFD1D5DB")
        # Eligible marker
        ec = ws.cell(row=r, column=2)
        ec.font = _font(color=GREEN, size=10, bold=True)
        ws.row_dimensions[r].height = 17
        r += 1

    r += 1

    # Eliminated
    apply_section_title(ws, r, 2, f"ELIMINATED LENDERS ({len(eliminated)})", colspan=3)
    ws.row_dimensions[r].height = 22
    r += 1
    apply_header_row(ws, r, ["", "Institution", "Tier", "Reason Eliminated"])
    ws.row_dimensions[r].height = 18
    r += 1

    for i, l in enumerate(eliminated):
        shade = i % 2 == 1
        bg = "FFE8EFF5" if shade else "FFFFFFFF"
        for ci, (val, col) in enumerate([
            ("",                  BLACK_CALC),
            (l.get("name",""),    "FF9CA3AF"),
            (l.get("tier",""),    "FF9CA3AF"),
            (l.get("status",""),  "FF9CA3AF"),
        ], 1):
            c = ws.cell(row=r, column=ci, value=val)
            c.font = _font(color=col, size=10)
            c.fill = _fill(bg)
            c.alignment = _align("left")
            c.border = _border("FFD1D5DB")
        ws.row_dimensions[r].height = 17
        r += 1

    ws.freeze_panes = "B2"


# ── Tab 7: Flags ──────────────────────────────────────────────────────────────

def build_flags(wb, a):
    ws = wb.create_sheet("Flags")
    ws.sheet_view.showGridLines = False

    flags = a.get("flags", [])

    for i, w in enumerate([2, 16, 36, 36, 2], 1):
        set_col_width(ws, i, w)

    r = 1
    ws.merge_cells(f"B{r}:D{r}")
    c = ws.cell(row=r, column=2, value=f"FLAGS & MITIGANTS ({len(flags)} identified)")
    c.font = _font(bold=True, color=GOLD_LIGHT, size=13)
    c.fill = _fill(NAVY_BASE)
    c.alignment = _align("left", "center")
    ws.row_dimensions[r].height = 28
    r += 2

    if not flags:
        ws.merge_cells(f"B{r}:D{r}")
        c = ws.cell(row=r, column=2, value="No material flags identified. Credit package presents cleanly.")
        c.font = _font(color=GREEN, size=10)
        c.alignment = _align("left")
        return

    apply_header_row(ws, r, ["", "Severity / Code", "Detail", "Mitigant & Action"])
    ws.row_dimensions[r].height = 18
    r += 1

    for i, flag in enumerate(flags):
        severity = flag.get("severity", "")
        code     = flag.get("code", "")
        title    = flag.get("title", "")
        detail   = flag.get("detail", "")
        mitigant = flag.get("mitigant", "")
        action   = flag.get("action", "")

        sev_color = RED if severity == "CRITICAL" else AMBER
        bg_color  = "FFFDF0EF" if severity == "CRITICAL" else "FFFEFAF0"

        # Title row
        ws.merge_cells(f"B{r}:D{r}")
        c = ws.cell(row=r, column=2, value=f"[{severity}] {code} — {title}")
        c.font = _font(bold=True, color=sev_color, size=10)
        c.fill = _fill(bg_color)
        c.alignment = _align("left", wrap=True)
        c.border = Border(
            left=Side(style="medium", color=sev_color),
            top=Side(style="thin", color="FFD1D5DB"),
        )
        ws.row_dimensions[r].height = 20
        r += 1

        # Detail / Mitigant / Action rows
        for field_label, field_val in [("Detail:", detail), ("Mitigant:", mitigant), ("Required Action:", action)]:
            c_label = ws.cell(row=r, column=2, value=field_label)
            c_label.font = _font(bold=True, color="FF374151", size=9)
            c_label.fill = _fill(bg_color)
            c_label.alignment = _align("left")
            c_label.border = Border(left=Side(style="medium", color=sev_color))

            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
            c_val = ws.cell(row=r, column=3, value=field_val)
            c_val.font = _font(color="FF374151", size=9)
            c_val.fill = _fill(bg_color)
            c_val.alignment = _align("left", wrap=True)
            ws.row_dimensions[r].height = 32
            r += 1

        spacer_row(ws, r, ncols=5)
        ws.row_dimensions[r].height = 8
        r += 1

    ws.freeze_panes = "B2"


# ── Master assembler ──────────────────────────────────────────────────────────


# ── Tab 8: Amortization Schedule ──────────────────────────────────────────────

def build_amortization(wb, a):
    ws = wb.create_sheet("Amortization")
    ws.sheet_view.showGridLines = False

    tx = a["transaction"]
    principal   = tx["loanAmount"]
    annual_rate = tx["illustrativeRate"]
    term        = tx["termMonths"]
    monthly_pmt = tx["monthlyPayment"]
    monthly_r   = annual_rate / 12

    for i, w in enumerate([2, 10, 16, 16, 16, 16, 18, 2], 1):
        set_col_width(ws, i, w)

    r = 1
    ws.merge_cells(f"B{r}:G{r}")
    c = ws.cell(row=r, column=2, value="LOAN AMORTIZATION SCHEDULE")
    c.font = _font(bold=True, color=GOLD_LIGHT, size=13)
    c.fill = _fill(NAVY_BASE)
    c.alignment = _align("left", "center")
    ws.row_dimensions[r].height = 28
    r += 2

    # Parameters summary
    apply_section_title(ws, r, 2, "LOAN PARAMETERS", colspan=6)
    ws.row_dimensions[r].height = 22
    r += 1

    params = [
        ("Loan Amount",        principal,    FMT_CURRENCY, BLUE_INPUT),
        ("Annual Rate",        annual_rate,  FMT_PCT,      BLUE_INPUT),
        ("Term (months)",      term,         FMT_NUMBER,   BLUE_INPUT),
        ("Monthly Payment",    monthly_pmt,  FMT_CURRENCY, BLACK_CALC),
        ("Annual Payment",     monthly_pmt * 12, FMT_CURRENCY, BLACK_CALC),
        ("Total Interest",     monthly_pmt * term - principal, FMT_CURRENCY, BLACK_CALC),
        ("Total Cost of Loan", monthly_pmt * term, FMT_CURRENCY, BLACK_CALC),
    ]
    for i2, (label, val, fmt, vc) in enumerate(params):
        lc, vc_cell = apply_kv_row(ws, r, label, val, fmt=fmt, shade=i2%2==1,
                                    value_color=vc, label_col=2, value_col=4)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # Schedule header
    apply_section_title(ws, r, 2, "MONTHLY SCHEDULE", colspan=6)
    ws.row_dimensions[r].height = 22
    r += 1

    apply_header_row(ws, r, ["", "Month", "Payment", "Principal", "Interest", "Balance", "Cumul. Interest"])
    ws.row_dimensions[r].height = 18
    r += 1

    # Build schedule row by row
    balance = principal
    cumul_interest = 0

    for month in range(1, term + 1):
        interest_pmt  = balance * monthly_r
        principal_pmt = monthly_pmt - interest_pmt
        balance      -= principal_pmt
        cumul_interest += interest_pmt

        # Clamp final month rounding
        if month == term:
            principal_pmt += balance
            balance = 0

        shade = month % 2 == 0

        # Highlight annual summary rows
        is_annual = month % 12 == 0
        bg = NAVY_MID if is_annual else ("FFE8EFF5" if shade else "FFFFFFFF")
        fc = GOLD_LIGHT if is_annual else BLACK_CALC

        for ci, (val, fmt) in enumerate([
            ("",              None),
            (month,           FMT_NUMBER),
            (monthly_pmt,     FMT_CURRENCY),
            (principal_pmt,   FMT_CURRENCY),
            (interest_pmt,    FMT_CURRENCY),
            (max(balance, 0), FMT_CURRENCY),
            (cumul_interest,  FMT_CURRENCY),
        ], 1):
            c = ws.cell(row=r, column=ci, value=round(val, 2) if isinstance(val, float) else val)
            c.font = _font(color=fc, size=9, bold=is_annual)
            c.fill = _fill(bg)
            c.alignment = _align("center" if ci == 2 else "right")
            c.border = _border("FFD1D5DB")
            if fmt: c.number_format = fmt
        ws.row_dimensions[r].height = 14
        r += 1

    # Final totals
    total_row(ws, r, [
        "", "TOTALS",
        round(monthly_pmt * term, 2),
        round(principal, 2),
        round(monthly_pmt * term - principal, 2),
        0,
        round(monthly_pmt * term - principal, 2),
    ], fmts=[None, None, FMT_CURRENCY, FMT_CURRENCY, FMT_CURRENCY, FMT_CURRENCY, FMT_CURRENCY])
    ws.row_dimensions[r].height = 20

    ws.freeze_panes = "B9"


# ── Tab 9: Sensitivity Analysis ───────────────────────────────────────────────

def build_sensitivity(wb, a):
    ws = wb.create_sheet("Sensitivity")
    ws.sheet_view.showGridLines = False

    tx  = a["transaction"]
    inc = a["incomeNormalization"]
    g   = a["gdscr"]

    after_tax_income = inc.get("afterTaxQualifyingIncome") or g.get("afterTaxIncome", 0)
    base_loan        = tx["loanAmount"]
    base_rate        = tx["illustrativeRate"]
    term             = tx["termMonths"]
    existing_ds      = tx["existingAnnualDS"]
    purchase_price   = tx["purchasePrice"]

    for i, w in enumerate([2, 22, 14, 14, 14, 14, 14, 14, 2], 1):
        set_col_width(ws, i, w)

    r = 1
    ws.merge_cells(f"B{r}:H{r}")
    c = ws.cell(row=r, column=2, value="SENSITIVITY ANALYSIS")
    c.font = _font(bold=True, color=GOLD_LIGHT, size=13)
    c.fill = _fill(NAVY_BASE)
    c.alignment = _align("left", "center")
    ws.row_dimensions[r].height = 28
    r += 2

    # Base case summary
    apply_section_title(ws, r, 2, "BASE CASE", colspan=7)
    ws.row_dimensions[r].height = 22
    r += 1

    def monthly_payment(p, annual_r, n):
        if annual_r == 0: return p / n
        mr = annual_r / 12
        return p * (mr * (1 + mr)**n) / ((1 + mr)**n - 1)

    def gdscr_calc(loan, rate):
        pmt   = monthly_payment(loan, rate, term)
        ann_ds = pmt * 12 + existing_ds
        return after_tax_income / ann_ds if ann_ds > 0 else 0

    base_pmt   = monthly_payment(base_loan, base_rate, term)
    base_ann   = base_pmt * 12 + existing_ds
    base_gdscr = after_tax_income / base_ann if base_ann > 0 else 0
    base_ltv   = base_loan / purchase_price

    base_rows = [
        ("After-Tax Qualifying Income",  after_tax_income, FMT_CURRENCY, GREEN_LINK),
        ("Loan Amount",                  base_loan,        FMT_CURRENCY, BLUE_INPUT),
        ("LTV",                          base_ltv,         FMT_PCT,      BLACK_CALC),
        ("Illustrative Rate",            base_rate,        FMT_PCT,      BLUE_INPUT),
        ("Term (months)",                term,             FMT_NUMBER,   BLUE_INPUT),
        ("Monthly Payment",              base_pmt,         FMT_CURRENCY, BLACK_CALC),
        ("Existing Annual DS",           existing_ds,      FMT_CURRENCY, GREEN_LINK),
        ("Total Pro Forma DS",           base_ann,         FMT_CURRENCY, BLACK_CALC),
        ("Base GDSCR",                   base_gdscr,       FMT_MULTIPLE, BLACK_CALC),
    ]
    for i2, (label, val, fmt, vc) in enumerate(base_rows):
        lc, vc_cell = apply_kv_row(ws, r, label, round(val,4) if isinstance(val,float) else val,
                                    fmt=fmt, shade=i2%2==1, value_color=vc, label_col=2, value_col=4)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # ── GDSCR sensitivity: Rate vs Loan Amount ────────────────────────────────
    apply_section_title(ws, r, 2, "GDSCR SENSITIVITY — RATE vs LOAN AMOUNT", colspan=7)
    ws.row_dimensions[r].height = 22
    r += 1

    rates     = [base_rate - 0.01, base_rate - 0.005, base_rate, base_rate + 0.005, base_rate + 0.01, base_rate + 0.015]
    loan_pcts = [0.65, 0.70, 0.75, 0.80, 0.85, 0.90]
    loans     = [round(purchase_price * p / 50000) * 50000 for p in loan_pcts]

    def fmt_currency(v):
        return f"${v/1000000:.2f}M" if v >= 1000000 else f"${v:,.0f}"

    header_vals = ["", "Rate \\ Loan →"] + [fmt_currency(l) for l in loans]
    apply_header_row(ws, r, header_vals)
    ws.cell(row=r, column=3).value = "Rate \\ Loan →"
    ws.row_dimensions[r].height = 18
    r += 1

    # LTV sub-header
    ltv_vals = ["", "LTV →"] + [f"{l/purchase_price*100:.0f}%" for l in loans]
    for ci, val in enumerate(ltv_vals, 1):
        c = ws.cell(row=r, column=ci, value=val)
        c.font = _font(color=STEEL, size=9, italic=True)
        c.fill = _fill(NAVY_MID)
        c.alignment = _align("center" if ci > 2 else "left")
    ws.row_dimensions[r].height = 14
    r += 1

    for ri, rate in enumerate(rates):
        is_base_rate = abs(rate - base_rate) < 0.0001
        for ci, loan in enumerate(loans, 1):
            gdscr_val  = gdscr_calc(loan, rate)
            is_base    = is_base_rate and abs(loan - base_loan) < 1000
            is_viable  = gdscr_val >= 1.25

            bg = NAVY_MID if is_base else ("FFE8F5E9" if gdscr_val >= 1.50 else
                                            "FFFFF8E1" if gdscr_val >= 1.25 else
                                            "FFFDF0EF")
            fc = GREEN if gdscr_val >= 1.50 else AMBER if gdscr_val >= 1.25 else RED
            if is_base: fc = GOLD_LIGHT

            # Rate label (first col of each row)
            if ci == 1:
                lc = ws.cell(row=r, column=2, value=f"{rate*100:.2f}%" + (" ◄ base" if is_base_rate else ""))
                lc.font = _font(color=GOLD_LIGHT if is_base_rate else "FF374151", size=9,
                                bold=is_base_rate)
                lc.fill = _fill(NAVY_MID if is_base_rate else ("FFE8EFF5" if ri%2==1 else "FFFFFFFF"))
                lc.alignment = _align("left")
                lc.border = _border("FFD1D5DB")

            c = ws.cell(row=r, column=ci + 2, value=round(gdscr_val, 2))
            c.font = _font(color=fc, size=10, bold=is_base)
            c.fill = _fill(bg)
            c.alignment = _align("center")
            c.border = _border("FFD1D5DB")
            c.number_format = FMT_MULTIPLE

        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # ── GDSCR sensitivity: Income stress ──────────────────────────────────────
    apply_section_title(ws, r, 2, "INCOME STRESS TEST — GDSCR AT REDUCED INCOME LEVELS", colspan=7)
    ws.row_dimensions[r].height = 22
    r += 1

    apply_header_row(ws, r, ["", "Income Scenario", "Gross Income", "After-Tax Est.", "Total DS", "GDSCR", "Assessment"])
    ws.row_dimensions[r].height = 18
    r += 1

    gross_income   = inc.get("qualifyingIncome", after_tax_income)
    taxes_paid     = inc.get("taxesPaidLowerYear", 0)
    eff_tax_rate   = taxes_paid / gross_income if gross_income > 0 else 0.36
    base_total_ds  = tx["totalProFormaDS"]

    stress_scenarios = [
        ("Base Case (Qualifying Year)",   1.00),
        ("5% Income Reduction",           0.95),
        ("10% Income Reduction",          0.90),
        ("15% Income Reduction",          0.85),
        ("20% Income Reduction",          0.80),
        ("25% Income Reduction",          0.75),
        ("30% Income Reduction — Stress", 0.70),
    ]

    for i2, (label, pct) in enumerate(stress_scenarios):
        stressed_gross    = gross_income * pct
        stressed_after_tax = stressed_gross * (1 - eff_tax_rate)
        stressed_gdscr    = stressed_after_tax / base_total_ds if base_total_ds > 0 else 0
        assessment = ("Strong" if stressed_gdscr >= 1.75 else
                      "Adequate" if stressed_gdscr >= 1.25 else
                      "Marginal" if stressed_gdscr >= 1.0 else
                      "Below threshold")
        is_base = i2 == 0
        shade   = i2 % 2 == 1
        gdscr_color = GREEN if stressed_gdscr >= 1.5 else AMBER if stressed_gdscr >= 1.0 else RED
        if is_base: gdscr_color = GOLD_LIGHT

        data_row(ws, r,
            ["", label, stressed_gross, stressed_after_tax, base_total_ds, stressed_gdscr, assessment],
            fmts=[None, None, FMT_CURRENCY, FMT_CURRENCY, FMT_CURRENCY, FMT_MULTIPLE, None],
            shade=shade,
            colors=["", BLACK_CALC, BLUE_INPUT if is_base else BLACK_CALC,
                    BLACK_CALC, GREEN_LINK, gdscr_color, gdscr_color]
        )
        ws.row_dimensions[r].height = 18
        r += 1

    r += 1

    # Legend
    apply_section_title(ws, r, 2, "COLOR LEGEND", colspan=7)
    ws.row_dimensions[r].height = 22
    r += 1

    legend = [
        ("Green",  "GDSCR ≥ 1.50x — Adequate to Strong",   "FFE8F5E9", GREEN),
        ("Amber",  "GDSCR 1.25x–1.49x — Marginal",          "FFFFF8E1", AMBER),
        ("Red",    "GDSCR < 1.25x — Below threshold",        "FFFDF0EF", RED),
        ("Gold",   "Base case",                              NAVY_MID,   GOLD_LIGHT),
    ]
    for color_name, desc, bg, fc in legend:
        c_label = ws.cell(row=r, column=2, value=color_name)
        c_label.font = _font(color=fc, size=10, bold=True)
        c_label.fill = _fill(bg)
        c_label.border = _border("FFD1D5DB")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        c_desc = ws.cell(row=r, column=3, value=desc)
        c_desc.font = _font(color=fc, size=10)
        c_desc.fill = _fill(bg)
        c_desc.border = _border("FFD1D5DB")
        ws.row_dimensions[r].height = 17
        r += 1

    ws.freeze_panes = "B9"

def generate_workbook(analysis, output_path):
    wb = Workbook()

    build_cover(wb, analysis)
    build_income(wb, analysis)
    build_balance_sheet(wb, analysis)
    build_gdscr(wb, analysis)
    build_scorecard(wb, analysis)
    build_lenders(wb, analysis)
    build_flags(wb, analysis)
    build_amortization(wb, analysis)
    build_sensitivity(wb, analysis)

    # Tab colors
    tab_colors = {
        "Cover":          "C8A96E",
        "Income":         "0B1E3A",
        "Balance Sheet":  "0B1E3A",
        "GDSCR":          "0B1E3A",
        "Scorecard":      "0B1E3A",
        "Lenders":        "0B1E3A",
        "Flags":          "C0392B",
        "Amortization":   "0B1E3A",
        "Sensitivity":    "0B1E3A",
    }
    for sheet_name, color in tab_colors.items():
        if sheet_name in wb.sheetnames:
            wb[sheet_name].sheet_properties.tabColor = color

    wb.save(output_path)
    size_kb = round(__import__("os").path.getsize(output_path) / 1024, 1)
    print(f"Workbook written: {output_path}  ({size_kb} KB, {len(wb.sheetnames)} tabs)")
    return output_path


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python generate_workbook.py <analysis.json> <output.xlsx>")
        sys.exit(1)
    with open(sys.argv[1]) as f:
        data = json.load(f)
    generate_workbook(data, sys.argv[2])
