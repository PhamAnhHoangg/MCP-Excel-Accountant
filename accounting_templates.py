"""Workbook templates for the Excel Accounting MCP.

Every accounting workbook = a 'Guide' sheet + the MUST-HAVE SHEETS in reading order (1. Canvas, 2. ..., each a real tab).
The Guide lists them and highlights the most important ones; the detail (numbers, taxes, payouts, controls) lives in the tabs.
"""
from __future__ import annotations

from datetime import date, datetime

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

import company_profile as P

GUIDE_SHEET = "Guide"
F, RED, NAVY = "Arial", "C00020", "1F3A5F"
VND, PCT, MON, NUM = '#,##0;(#,##0);-', '0.0%', 'mmm-yy', '#,##0.0'
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
FILL_CORE = PatternFill("solid", fgColor="FCE4D6")
FILL_NAVY = PatternFill("solid", fgColor=NAVY)
_thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

DEFAULT_EXTEND = ("When adding a new item (product, customer, month): insert a new row in the MIDDLE of the data range (not at the bottom), "
                  "drag the formulas down from the row next to it, then check that the SUM / SUMPRODUCT formulas now include the new row.")

# (sheet name, what it is for, MUST-HAVE?)
FULL_ORDER = [
    ("Canvas", "One page showing the whole model: customers, value, revenue, costs, margin, cash, tax and owner payout. Key numbers are linked from the other sheets.", True),
    ("Customers & Strategy", "Brand, customer profile, strategy and internal rules: our draft in one column, your real interview answers in the yellow column.", False),
    ("To-do", "What to find out and verify before trusting the numbers: owner, deadline, status.", False),
    ("Assumptions", "Every number the model uses: prices, costs, payment terms, exchange rates, tax and payroll rates with source and date. All other sheets read from here.", True),
    ("Sales policy", "Channels, commission, discount limits, and where each 100 VND of price goes.", False),
    ("Fixed costs", "Monthly fixed costs and one-off spend. Decides break-even and the cash you need every month.", False),
    ("Sales plan", "Units per month, revenue, direct cost, gross profit, with best / base / worst scenario.", False),
    ("Break-even", "Units and revenue per month needed to cover all costs, and the lowest price you can accept.", True),
    ("Costing", "Full cost per unit (material, labour, install, overhead, warranty, finance) and the price that gives your target margin.", True),
    ("Cash flow", "12-month cash: when money comes in and goes out, lowest balance and capital needed.", True),
    ("Tax", "VAT, corporate income tax, payroll taxes and insurance: calculation, payments and the due-date list.", True),
    ("Payouts", "Payroll, bonuses and commission, owner payouts, loans: who is paid what and when.", True),
    ("Agreements & Controls", "Owner / partner agreement, approval limits, cross-checks that must be zero, sign-off and change log.", False),
]


# ------------------------------------------------------------------ helpers
def q(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def plain(addr: str) -> str:
    return addr.replace("$", "")


def style(c, kind="calc", fmt=None, wrap=False):
    color = {"input": "0000FF", "link": "008000"}.get(kind, "000000")
    c.font = Font(name=F, size=10, color=color, bold=kind in ("key", "label", "hdr"))
    if kind == "note":
        c.font = Font(name=F, size=9, italic=True, color="595959")
    c.border = BOX
    if kind == "input":
        c.fill = FILL_IN
    elif kind == "key":
        c.fill = FILL_KEY
    elif kind == "hdr":
        c.fill = FILL_NAVY
        c.font = Font(name=F, size=10, bold=True, color="FFFFFF")
    c.alignment = Alignment(vertical="top" if wrap else "center", wrap_text=wrap,
                            horizontal="center" if kind == "hdr" else None)
    if fmt:
        c.number_format = fmt


def put(ws, addr, v, kind="calc", fmt=None, wrap=False):
    c = ws[addr]
    c.value = v
    style(c, kind, fmt, wrap)
    return c


def head(ws, text, sub=""):
    ws["A1"] = text
    ws["A1"].font = Font(name=F, size=14, bold=True, color=RED)
    ws["A2"] = sub
    ws["A2"].font = Font(name=F, size=9, italic=True, color="595959")
    ws.sheet_view.showGridLines = False


def hdr_row(ws, row, labels, first_col=1):
    for i, t in enumerate(labels):
        put(ws, f"{L(first_col + i)}{row}", t, "hdr", wrap=True)
    ws.row_dimensions[row].height = 30


def bar(ws, row, text, ncols):
    for c in range(1, ncols + 1):
        ws.cell(row=row, column=c).fill = FILL_NAVY
    x = ws.cell(row=row, column=1, value=text)
    x.font = Font(name=F, size=10, bold=True, color="FFFFFF")


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[L(i)].width = w


def month_cols():
    return [L(c) for c in range(2, 14)]  # B..M


def month_header(ws, row, start_ref):
    put(ws, f"A{row}", "Item", "hdr")
    for k, col in enumerate(month_cols()):
        c = put(ws, f"{col}{row}", f"=EDATE({start_ref},{k})", "hdr", MON)
    put(ws, f"N{row}", "Year", "hdr")
    ws.column_dimensions["A"].width = 46
    for col in month_cols():
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["N"].width = 16
    ws.freeze_panes = f"B{row + 1}"


def mrow(ws, r, label, fn, kind="calc", fmt=VND, total="sum", first_kind=None):
    """fn(col, prev_col, k) -> value/formula for month k (0..11). total: 'sum' | 'avg' | None."""
    put(ws, f"A{r}", label, "label" if kind == "key" else "calc")
    cols = month_cols()
    for k, col in enumerate(cols):
        prev = L(1 + k) if k else None
        kd = first_kind if (k == 0 and first_kind) else kind
        put(ws, f"{col}{r}", fn(col, prev, k), kd, fmt)
    if total == "sum":
        put(ws, f"N{r}", f"=SUM(B{r}:M{r})", "key" if kind == "key" else "calc", fmt)
    elif total == "avg":
        put(ws, f"N{r}", f"=AVERAGE(B{r}:M{r})", "calc", fmt)
    else:
        put(ws, f"N{r}", None, "calc")


def idx_row(ws, r):
    put(ws, f"A{r}", "Month no. (1 to 12)", "note")
    for k, col in enumerate(month_cols()):
        put(ws, f"{col}{r}", k + 1, "note", "0")


# --------------------------------------------------------------------- Guide
def add_guide_sheet(wb, title, subtitle, order, questions, extend_note=""):
    """order: list of (sheet name, purpose, must_have). Must-have rows are highlighted."""
    ws = wb.create_sheet(GUIDE_SHEET, 0)
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    widths(ws, [34, 100, 16])
    wrap = Alignment(wrap_text=True, vertical="center")
    r = 1
    ws.cell(row=r, column=1, value=title.upper()).font = Font(name=F, size=15, bold=True, color=RED)
    r += 1
    ws.cell(row=r, column=1, value=(subtitle or "All sample numbers are EXAMPLES. Replace them with real figures.")
            + f"  (Created {datetime.now():%d/%m/%Y})").font = Font(name=F, size=9, italic=True, color="595959")
    r += 2

    def heading(t):
        nonlocal r
        ws.cell(row=r, column=1, value=t).font = Font(name=F, size=11, bold=True, color=RED)
        r += 1

    def line(a, b, c="", fills=(None, None, None), afont=None):
        nonlocal r
        cells = [ws.cell(row=r, column=i, value=v) for i, v in enumerate((a, b, c), 1)]
        for cell, fill in zip(cells, fills):
            cell.font = Font(name=F, size=10, bold=(cell.column != 2))
            cell.border, cell.alignment = BOX, wrap
            if fill:
                cell.fill = fill
        if afont:
            cells[0].font = afont
        ws.row_dimensions[r].height = max(18, 15 * (1 + len(b or "") // 100))
        r += 1
        return cells

    heading("READING ORDER = MUST-HAVE SETUP SHEETS (EACH ONE IS A TAB NEXT TO THIS GUIDE)")
    for i, t in enumerate(["Sheet (click to open)", "What it is for", "Importance"]):
        x = ws.cell(row=r, column=i + 1, value=t)
        x.font, x.fill, x.border = Font(name=F, size=10, bold=True, color="FFFFFF"), FILL_NAVY, BOX
    r += 1
    for i, (name, purpose, core) in enumerate(order, 1):
        fills = (FILL_CORE, FILL_CORE, FILL_CORE) if core else (None, None, None)
        cells = line(f"{i}. {name}", purpose, "MUST-HAVE" if core else "Add-on", fills)
        cells[0].hyperlink = Hyperlink(ref=cells[0].coordinate, location=f"{q(name)}!A1")
        cells[0].font = Font(name=F, size=10, bold=True, underline="single")
    c = ws.cell(row=r, column=1, value="Highlighted rows = the most important sheets. Build and check these first; the add-on sheets can follow.")
    c.font = Font(name=F, size=9, italic=True, color="595959")
    r += 2

    heading("COLOUR CODE")
    line("Blue text, light-yellow fill", "Input cell: you may edit", "", (FILL_IN, None, None), Font(name=F, size=10, color="0000FF"))
    line("Black text", "Formula: do not edit", "", afont=Font(name=F, size=10))
    line("Green text", "Pulls a number from another sheet", "", afont=Font(name=F, size=10, color="008000"))
    line("Bright-yellow fill", "Key result to look at (or a number not yet confirmed, marked VERIFY)", "", (FILL_KEY, None, None))
    r += 1

    heading("ADDING NEW ITEMS")
    c = ws.cell(row=r, column=1, value=extend_note or DEFAULT_EXTEND)
    c.font, c.alignment = Font(name=F, size=9, italic=True, color="595959"), wrap
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.row_dimensions[r].height = 30
    r += 2

    heading("SELF-CHECK QUESTIONS (SCENARIOS)")
    c = ws.cell(row=r, column=1, value="Change the numbers as each question says, watch which sheets move, and explain why. This is how you learn the accounting logic of the model.")
    c.font = Font(name=F, size=9, italic=True, color="595959")
    r += 1
    for i, qn in enumerate(questions, 1):
        c = ws.cell(row=r, column=1, value=f"{i}. {qn}")
        c.font, c.alignment = Font(name=F, size=10), wrap
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 32
        r += 1


# ------------------------------------------------------------- Assumptions
def _rate_row(ref, key, label, value, unit, rid, fmt, extra=""):
    it = ref[rid]
    verify = any(w in it["confidence"].lower() for w in ("verify", "single source"))
    note = it["confidence"] + (" | " + it["note"] if it.get("note") else "") + (" | " + extra if extra else "")
    return (key, label, value, unit, it["source"], it["checked_on"], note, verify, fmt)


def _ex(key, label, value, unit="", fmt=None, note=""):
    return (key, label, value, unit, "EXAMPLE - replace with real figure", "", note, False, fmt)


def assumptions_spec(entity, ref, start):
    S = [("sec", "GENERAL")]
    S += [("entity", "Entity (@B1@ / @B2@ / @B3@)", entity, "", "You", "", "Never mix entities in one file total.", False, None),
          ("start", "First month of the model", start, "date", "You", "", "All month headers follow this date.", False, "mmm-yy"),
          ("vat_incl", "Are prices and costs entered WITH VAT?", "No", "Yes / No", "You", "", "This model assumes prices WITHOUT VAT. VAT is added in the Tax sheet.", False, None)]
    S.append(("sec", "SALES AND PRICE (EXAMPLES)"))
    S += [_ex("price", "Selling price per unit (e.g. 1 m2 of window or door)", 4500000, "VND", VND),
          _ex("units", "Units sold in the first month", 1000, "units", '#,##0'),
          _ex("growth", "Growth in units per month", 0.02, "%", PCT),
          _ex("scenario", "Scenario in use (Best / Base / Worst)", "Base", "", None),
          _ex("f_best", "Scenario factor: Best", 1.2, "x", '0.00'),
          _ex("f_base", "Scenario factor: Base", 1.0, "x", '0.00'),
          _ex("f_worst", "Scenario factor: Worst", 0.7, "x", '0.00'),
          _ex("export_share", "Share of revenue that is exported (0% VAT)", 0.0, "%", PCT, "@B2@: raise this to model export orders.")]
    S.append(("sec", "DIRECT COST PER UNIT (EXAMPLES)"))
    S += [_ex("material", "Material (profile, glass, hardware)", 2300000, "VND", VND),
          _ex("labour", "Direct labour", 450000, "VND", VND),
          _ex("install", "Installation", 300000, "VND", VND),
          _ex("transport", "Transport and packing", 100000, "VND", VND),
          _ex("overhead_pct", "Factory overhead, % of direct cost", 0.08, "%", PCT),
          _ex("warranty_pct", "Warranty reserve, % of price", 0.015, "%", PCT),
          _ex("finance_pct", "Finance cost, % of price", 0.01, "%", PCT),
          _ex("target_margin", "Target margin on price", 0.25, "%", PCT)]
    S.append(("sec", "PAYMENT TERMS AND CASH TIMING (EXAMPLES)"))
    S += [_ex("cust_days", "Customer pays after (days)", 30, "days", '0', "Cash arrives this long after invoice. Retention and advance are not modelled here."),
          _ex("supp_days", "We pay suppliers after (days)", 30, "days", '0'),
          ("cust_lag", "Customer delay in months (calculated)", "=ROUND({cust_days}/30,0)", "months", "Calculated", "", "", False, '0'),
          ("supp_lag", "Supplier delay in months (calculated)", "=ROUND({supp_days}/30,0)", "months", "Calculated", "", "", False, '0'),
          _ex("buffer", "Minimum cash to keep in the bank", 100000000, "VND", VND)]
    S.append(("sec", "PAYROLL AND OWNER PAYOUTS (EXAMPLES)"))
    S += [_ex("payroll", "Gross payroll per month (all staff)", 600000000, "VND", VND),
          _ex("insured_pct", "Insured salary as % of gross payroll", 1.0, "%", PCT, "VERIFY per employee: caps apply above the insured-salary cap."),
          _ex("pit_pct", "Average personal income tax withheld, % of gross payroll", 0.03, "%", PCT, "Estimate only. Compute per employee in the PIT module."),
          _ex("bonus_pct", "Bid-win bonus and commission, % of gross profit", 0.03, "%", PCT),
          _ex("dividend", "Dividend per quarter", 150000000, "VND", VND, "VERIFY tax on dividends with your adviser."),
          _ex("loan", "Loan balance at start", 2000000000, "VND", VND),
          _ex("loan_rate", "Loan interest rate per year", 0.10, "%", PCT),
          _ex("loan_prin", "Loan principal repaid per month", 50000000, "VND", VND)]
    S.append(("sec", "OPENING BALANCES (AGREE TO BANK STATEMENT AND LEDGER)"))
    S += [_ex("cash_open", "Cash in bank at start", 500000000, "VND", VND),
          _ex("recv_open", "Customer receivables at start (collected in month 1)", 0, "VND", VND),
          _ex("pay_open", "Supplier payables at start (paid in month 1)", 0, "VND", VND)]
    S.append(("sec", "EXCHANGE RATES (ENTER TODAY'S BANK RATE AND DATE)"))
    S += [_ex("usd", "USD / VND", 26000, "VND", VND, "Which rate: bank buying / selling / customs. Record the date in the Source column."),
          _ex("eur", "EUR / VND", 30000, "VND", VND, "Used for imported hardware (module C07).")]
    S.append(("sec", "TAX AND PAYROLL RATES (SOURCED; YELLOW = CONFIRM WITH YOUR TAX ADVISER)"))
    S += [_rate_row(ref, "vat", "VAT rate on domestic sales", 0.10, "%", "vat_rates", PCT, "A temporary 2-point cut for eligible items runs to end-2026; check eligibility."),
          _rate_row(ref, "cit", "Corporate income tax rate", 0.20, "%", "cit_standard", PCT, "15% / 17% tiers by prior-year revenue may apply (see cit_small): confirm."),
          _rate_row(ref, "er_si", "Employer social insurance", 0.175, "%", "ins_rates", PCT),
          _rate_row(ref, "er_hi", "Employer health insurance", 0.03, "%", "ins_rates", PCT),
          _rate_row(ref, "er_ui", "Employer unemployment insurance", 0.01, "%", "ins_rates", PCT),
          _rate_row(ref, "er_un", "Employer union fee", 0.02, "%", "ins_rates", PCT),
          ("er_tot", "Employer insurance and union fee, total (calculated)", "={er_si}+{er_hi}+{er_ui}+{er_un}", "%", "Calculated", "", "", False, PCT),
          _rate_row(ref, "ee_si", "Employee social insurance", 0.08, "%", "ins_rates", PCT),
          _rate_row(ref, "ee_hi", "Employee health insurance", 0.015, "%", "ins_rates", PCT),
          _rate_row(ref, "ee_ui", "Employee unemployment insurance", 0.01, "%", "ins_rates", PCT),
          ("ee_tot", "Employee insurance, total (calculated)", "={ee_si}+{ee_hi}+{ee_ui}", "%", "Calculated", "", "", False, PCT),
          _rate_row(ref, "cap_si", "Cap on salary for social and health insurance, per month", 50600000, "VND", "ins_caps", VND, "Not applied per employee in this model."),
          _rate_row(ref, "cap_ui", "Cap on salary for unemployment insurance, per month", 106200000, "VND", "ins_caps", VND),
          _rate_row(ref, "pit_self", "Personal income tax: personal deduction per month", 15500000, "VND", "pit_deductions", VND),
          _rate_row(ref, "pit_dep", "Personal income tax: deduction per dependent per month", 6200000, "VND", "pit_deductions", VND),
          _rate_row(ref, "minw", "Regional minimum wage, Region I, per month (2026)", 5310000, "VND", "min_wage_2026", VND)]
    return S


def build_assumptions(wb, ref, entity):
    ws = wb["Assumptions"]
    head(ws, "ASSUMPTIONS", "Every number the model uses. Blue/yellow cells are yours. Rates have source, date and confidence: yellow in the last column = confirm first.")
    widths(ws, [58, 20, 12, 64, 12, 70])
    hdr_row(ws, 4, ["Item", "Value", "Unit", "Source", "Checked on", "Confidence / note"])
    spec = assumptions_spec(entity, ref, date(datetime.now().year + 1, 1, 1))
    addr, r = {}, 5
    rows = []
    for e in spec:
        if e[0] == "sec":
            rows.append((r, e)); r += 1
        else:
            addr[e[0]] = f"$B${r}"; rows.append((r, e)); r += 1
    verify = []
    for r, e in rows:
        if e[0] == "sec":
            bar(ws, r, e[1], 6)
            continue
        key, label, value, unit, src, chk, note, ver, fmt = e
        put(ws, f"A{r}", label, "calc", wrap=True)
        calc = isinstance(value, str) and value.startswith("=")
        if calc:
            value = value.format(**addr)
        put(ws, f"B{r}", value, "calc" if calc else "input", fmt)
        put(ws, f"C{r}", unit, "calc")
        put(ws, f"D{r}", src, "calc", wrap=True)
        put(ws, f"E{r}", chk, "calc")
        c = put(ws, f"F{r}", note, "calc", wrap=True)
        if ver:
            c.fill = FILL_KEY
            verify.append(label)
        if len(note) > 70 or len(src) > 60:
            ws.row_dimensions[r].height = 30 if len(note) < 140 else 44
    dv = DataValidation(type="list", formula1='"Best,Base,Worst"', allow_blank=False); ws.add_data_validation(dv); dv.add(addr["scenario"].replace("$", ""))
    dv2 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False); ws.add_data_validation(dv2); dv2.add(addr["vat_incl"].replace("$", ""))
    ws.freeze_panes = "B5"
    A = {k: f"Assumptions!{v}" for k, v in addr.items()}
    return A, verify


# ------------------------------------------------------------- other sheets
def build_costing(wb, A):
    ws = wb["Costing"]
    head(ws, "COSTING", "Full cost per unit built from Assumptions, and the price needed for the target margin.")
    widths(ws, [54, 22, 14, 70])
    hdr_row(ws, 4, ["Per unit", "VND", "% of price", "Note"])
    rows = [(5, "Material", f"={A['material']}", "link"), (6, "Direct labour", f"={A['labour']}", "link"),
            (7, "Installation", f"={A['install']}", "link"), (8, "Transport and packing", f"={A['transport']}", "link"),
            (9, "Direct cost per unit", "=SUM(B5:B8)", "key"),
            (10, "Factory overhead", f"=B9*{A['overhead_pct']}", "calc"),
            (11, "Warranty reserve", f"=B14*{A['warranty_pct']}", "calc"),
            (12, "Finance cost", f"=B14*{A['finance_pct']}", "calc"),
            (13, "Full cost per unit", "=B9+B10+B11+B12", "key"),
            (14, "Current selling price", f"={A['price']}", "link"),
            (15, "Margin at current price", "=IFERROR((B14-B13)/B14,0)", "key"),
            (16, "Target margin", f"={A['target_margin']}", "link"),
            (17, "Suggested price for the target margin", "=IFERROR(B13/(1-B16),0)", "key"),
            (18, "Price gap versus current price", "=B17-B14", "calc")]
    for r, lab, f, k in rows:
        put(ws, f"A{r}", lab, "label" if k == "key" else "calc")
        put(ws, f"B{r}", f, k, PCT if r in (15, 16) else VND)
        if r <= 13:
            put(ws, f"C{r}", f"=IFERROR(B{r}/$B$14,0)", "calc", PCT)
    notes = {9: "Material + labour + install + transport", 11: "Based on current price", 15: "(Price - full cost) / price",
             17: "Full cost / (1 - target margin). Warranty and finance are held at the current price."}
    for r, t in notes.items():
        put(ws, f"D{r}", t, "note")


def build_fixed(wb, A):
    ws = wb["Fixed costs"]
    head(ws, "FIXED COSTS", "Costs you pay even when sales are zero. Change the first month; later months copy it (overwrite any month if needed).")
    month_header(ws, 4, A["start"])
    mrow(ws, 5, "Rent and utilities", lambda c, p, k: 60000000 if k == 0 else f"={p}5", "input", first_kind="input")
    for k, col in enumerate(month_cols()[1:], 1):
        ws[f"{col}5"].font = Font(name=F, size=10)
        ws[f"{col}5"].fill = PatternFill(fill_type=None)
    mrow(ws, 6, "Salaries (gross payroll)", lambda c, p, k: f"={A['payroll']}", "link")
    mrow(ws, 7, "Employer insurance and union fee", lambda c, p, k: f"={c}6*{A['insured_pct']}*{A['er_tot']}", "calc")
    mrow(ws, 8, "Depreciation (not cash)", lambda c, p, k: 40000000 if k == 0 else f"={p}8", "input", first_kind="input")
    mrow(ws, 9, "Marketing", lambda c, p, k: 30000000 if k == 0 else f"={p}9", "input", first_kind="input")
    mrow(ws, 10, "Other overhead", lambda c, p, k: 20000000 if k == 0 else f"={p}10", "input", first_kind="input")
    for r in (5, 8, 9, 10):
        for col in month_cols()[1:]:
            ws[f"{col}{r}"].font = Font(name=F, size=10)
            ws[f"{col}{r}"].fill = PatternFill(fill_type=None)
    mrow(ws, 11, "Total fixed costs (profit and loss view)", lambda c, p, k: f"=SUM({c}5:{c}10)", "key")
    mrow(ws, 12, "Of which non-payroll cash costs (rent, marketing, other)", lambda c, p, k: f"={c}5+{c}9+{c}10", "calc")
    mrow(ws, 13, "One-off investment, cash (example: a machine in month 3)", lambda c, p, k: 300000000 if k == 2 else 0, "input")
    put(ws, "A15", "EXAMPLE numbers. Salaries and insurance come from Assumptions; payroll cash is detailed in the Payouts sheet.", "note")


def build_sales_plan(wb, A):
    ws = wb["Sales plan"]
    head(ws, "SALES PLAN", "Units, revenue and gross profit by month for the scenario chosen in Assumptions.")
    month_header(ws, 4, A["start"])
    idx_row(ws, 5)
    mrow(ws, 6, "Planned units (base)", lambda c, p, k: f"={A['units']}*(1+{A['growth']})^({c}$5-1)", "calc", '#,##0')
    mrow(ws, 7, "Scenario factor in use", lambda c, p, k: f'=IF({A["scenario"]}="Best",{A["f_best"]},IF({A["scenario"]}="Worst",{A["f_worst"]},{A["f_base"]}))', "link", '0.00', None)
    mrow(ws, 8, "Units after scenario", lambda c, p, k: f"={c}6*{c}7", "calc", '#,##0')
    mrow(ws, 9, "Selling price per unit", lambda c, p, k: f"={A['price']}", "link", VND, "avg")
    mrow(ws, 10, "Revenue", lambda c, p, k: f"={c}8*{c}9", "key")
    mrow(ws, 11, "Of which domestic (VAT applies)", lambda c, p, k: f"={c}10*(1-{A['export_share']})", "calc")
    mrow(ws, 12, "Of which export (VAT 0%)", lambda c, p, k: f"={c}10*{A['export_share']}", "calc")
    mrow(ws, 13, "Direct cost (material, labour, install, transport)", lambda c, p, k: f"={c}8*Costing!$B$9", "calc")
    mrow(ws, 14, "Warranty and finance cost", lambda c, p, k: f"={c}8*(Costing!$B$11+Costing!$B$12)", "calc")
    mrow(ws, 15, "Gross profit", lambda c, p, k: f"={c}10-{c}13-{c}14", "key")


def build_breakeven(wb, A):
    ws = wb["Break-even"]
    head(ws, "BREAK-EVEN", "How many units each month you must sell to cover every fixed cost, and the lowest price you can accept.")
    widths(ws, [56, 22, 70])
    hdr_row(ws, 4, ["Item", "Value", "Note"])
    rows = [(5, "Selling price per unit", f"={A['price']}", "link", VND, ""),
            (6, "Variable cost per unit (direct + warranty + finance)", "=Costing!B9+Costing!B11+Costing!B12", "link", VND, ""),
            (7, "Contribution per unit", "=B5-B6", "calc", VND, "What each unit leaves to pay fixed costs"),
            (8, "Fixed cost per month (average, profit and loss view)", "=AVERAGE('Fixed costs'!B11:M11)", "link", VND, ""),
            (9, "Loan interest per month (average)", "=AVERAGE(Payouts!B21:M21)", "link", VND, ""),
            (10, "Total fixed cost per month", "=B8+B9", "calc", VND, ""),
            (11, "Break-even units per month", "=IFERROR(B10/B7,0)", "key", '#,##0', "Below this you lose money"),
            (12, "Break-even revenue per month", "=B11*B5", "key", VND, ""),
            (13, "Planned units per month (average, scenario in use)", "=AVERAGE('Sales plan'!B8:M8)", "link", '#,##0', ""),
            (14, "Margin of safety", "=IFERROR(1-B11/B13,0)", "key", PCT, "How far sales can fall before you lose money"),
            (15, "Lowest price at planned volume", "=IFERROR(B6+B10/B13,0)", "key", VND, "Below this price you lose money at the planned volume")]
    for r, lab, f, k, fmt, n in rows:
        put(ws, f"A{r}", lab, "label" if k == "key" else "calc")
        put(ws, f"B{r}", f, k, fmt)
        put(ws, f"C{r}", n, "note")


def build_payouts(wb, A):
    ws = wb["Payouts"]
    head(ws, "PAYOUTS", "Cash that leaves the company to people and lenders: payroll, bonuses, owner payouts, loans.")
    month_header(ws, 4, A["start"])
    bar(ws, 5, "PAYROLL", 14)
    mrow(ws, 6, "Gross payroll", lambda c, p, k: f"='Fixed costs'!{c}6", "link")
    mrow(ws, 7, "Employee insurance withheld", lambda c, p, k: f"={c}6*{A['insured_pct']}*{A['ee_tot']}", "calc")
    mrow(ws, 8, "Personal income tax withheld (estimate)", lambda c, p, k: f"={c}6*{A['pit_pct']}", "calc")
    mrow(ws, 9, "Net pay to employees", lambda c, p, k: f"={c}6-{c}7-{c}8", "calc")
    mrow(ws, 10, "Employer insurance and union fee", lambda c, p, k: f"={c}6*{A['insured_pct']}*{A['er_tot']}", "calc")
    mrow(ws, 11, "Payroll cash out (net pay + insurance + tax)", lambda c, p, k: f"={c}9+{c}7+{c}8+{c}10", "key")
    bar(ws, 12, "BONUSES AND COMMISSION", 14)
    mrow(ws, 13, "Gross profit (from Sales plan)", lambda c, p, k: f"='Sales plan'!{c}15", "link")
    mrow(ws, 14, "Bid-win bonus and commission", lambda c, p, k: f"=MAX(0,{c}13)*{A['bonus_pct']}", "key")
    bar(ws, 15, "OWNER PAYOUTS", 14)
    mrow(ws, 16, "Dividends (paid at quarter end)", lambda c, p, k: f"=IF(MOD(MONTH({c}$4),3)=0,{A['dividend']},0)", "calc")
    mrow(ws, 17, "Shareholder loan repayment (enter if any)", lambda c, p, k: 0, "input")
    bar(ws, 18, "LOANS", 14)
    mrow(ws, 19, "Opening balance", lambda c, p, k: f"={A['loan']}" if k == 0 else f"={p}22", "calc", VND, None)
    mrow(ws, 20, "Principal repaid", lambda c, p, k: f"=MIN({c}19,{A['loan_prin']})", "calc")
    mrow(ws, 21, "Interest", lambda c, p, k: f"={c}19*{A['loan_rate']}/12", "calc")
    mrow(ws, 22, "Closing balance", lambda c, p, k: f"={c}19-{c}20", "calc", VND, None)
    mrow(ws, 23, "Loan cash out (principal + interest)", lambda c, p, k: f"={c}20+{c}21", "key")
    mrow(ws, 25, "TOTAL PAYOUTS", lambda c, p, k: f"={c}11+{c}14+{c}16+{c}17+{c}23", "key")
    put(ws, "A27", "Suppliers and crews are paid through the Cash flow sheet (row 8). Keep invoice, acceptance and payment proof for each payment.", "note")


def build_tax(wb, A):
    ws = wb["Tax"]
    head(ws, "TAX", "VAT, corporate income tax and payroll taxes. Rates come from Assumptions. Unconfirmed rates are yellow there.")
    month_header(ws, 4, A["start"])
    bar(ws, 5, "VAT", 14)
    mrow(ws, 6, "Output VAT on domestic sales", lambda c, p, k: f"='Sales plan'!{c}11*{A['vat']}", "calc")
    mrow(ws, 7, "Export sales (VAT 0%)", lambda c, p, k: f"='Sales plan'!{c}12", "link")
    mrow(ws, 8, "Input VAT on purchases (direct cost)", lambda c, p, k: f"='Sales plan'!{c}13*{A['vat']}", "calc")
    mrow(ws, 9, "Net VAT for the month (output minus input)", lambda c, p, k: f"={c}6-{c}8", "calc")
    mrow(ws, 10, "Position before payment (includes credit carried forward)", lambda c, p, k: f"={c}9" if k == 0 else f"={p}12+{c}9", "calc", VND, None)
    mrow(ws, 11, "VAT paid", lambda c, p, k: f"=MAX(0,{c}10)", "key")
    mrow(ws, 12, "Credit carried forward", lambda c, p, k: f"=MIN(0,{c}10)", "calc", VND, None)
    bar(ws, 13, "CORPORATE INCOME TAX", 14)
    mrow(ws, 15, "Revenue", lambda c, p, k: f"='Sales plan'!{c}10", "link")
    mrow(ws, 16, "Direct, warranty and finance cost", lambda c, p, k: f"='Sales plan'!{c}13+'Sales plan'!{c}14", "link")
    mrow(ws, 17, "Fixed costs (profit and loss view)", lambda c, p, k: f"='Fixed costs'!{c}11", "link")
    mrow(ws, 18, "Loan interest", lambda c, p, k: f"=Payouts!{c}21", "link")
    mrow(ws, 19, "Bonuses and commission", lambda c, p, k: f"=Payouts!{c}14", "link")
    mrow(ws, 20, "Profit before tax", lambda c, p, k: f"={c}15-{c}16-{c}17-{c}18-{c}19", "key")
    mrow(ws, 21, "Add back non-deductible costs (enter)", lambda c, p, k: 0, "input")
    mrow(ws, 22, "Taxable income", lambda c, p, k: f"={c}20+{c}21", "calc")
    mrow(ws, 23, "Cumulative taxable income", lambda c, p, k: f"=SUM($B22:{c}22)", "calc", VND, None)
    mrow(ws, 24, "Cumulative tax due at the CIT rate", lambda c, p, k: f"=MAX(0,{c}23)*{A['cit']}", "calc", VND, None)
    mrow(ws, 25, "CIT paid (provisional, at quarter end)", lambda c, p, k: f"=IF(MOD(MONTH({c}$4),3)=0,MAX(0,{c}24-SUM($A25:{p or 'A'}25)),0)", "key")
    bar(ws, 27, "PAYROLL TAXES AND INSURANCE (FROM PAYOUTS)", 14)
    mrow(ws, 28, "Employer insurance and union fee", lambda c, p, k: f"=Payouts!{c}10", "link")
    mrow(ws, 29, "Employee insurance withheld", lambda c, p, k: f"=Payouts!{c}7", "link")
    mrow(ws, 30, "Personal income tax withheld", lambda c, p, k: f"=Payouts!{c}8", "link")
    mrow(ws, 31, "Total payroll taxes and insurance", lambda c, p, k: f"=SUM({c}28:{c}30)", "key")
    bar(ws, 33, "DUE DATES AND RESPONSIBILITY (SET FROM THE TAX AUTHORITY CALENDAR; VERIFY THE CURRENT DATES)", 14)
    for col_range, t in (("A", "Tax"), ("B:G", "How often and due rule"), ("H:J", "Responsible"), ("K:M", "Status")):
        pass
    put(ws, "A34", "Tax", "hdr"); put(ws, "B34", "How often and due rule", "hdr"); put(ws, "H34", "Responsible", "hdr"); put(ws, "K34", "Status", "hdr")
    ws.merge_cells("B34:G34"); ws.merge_cells("H34:J34"); ws.merge_cells("K34:M34")
    due = ["VAT (monthly or quarterly)", "Corporate income tax: quarterly provisional", "Corporate income tax: annual final",
           "Personal income tax withheld and final", "Insurance and union fee", "Import duty and import VAT (per shipment)",
           "Foreign contractor tax (per payment abroad)", "E-invoice timing check"]
    dv = DataValidation(type="list", formula1='"Not set,Set,Filed"', allow_blank=True); ws.add_data_validation(dv)
    for i, t in enumerate(due):
        r = 35 + i
        put(ws, f"A{r}", t, "calc")
        put(ws, f"B{r}", None, "input"); ws.merge_cells(f"B{r}:G{r}")
        put(ws, f"H{r}", None, "input"); ws.merge_cells(f"H{r}:J{r}")
        put(ws, f"K{r}", "Not set", "input"); ws.merge_cells(f"K{r}:M{r}")
        dv.add(f"K{r}")
    put(ws, "A44", "VAT refunds on exports are not modelled here: the credit carries forward. See module D03 (export VAT refund tracker).", "note")


def build_cashflow(wb, A):
    ws = wb["Cash flow"]
    head(ws, "CASH FLOW", "Profit is not cash. Money arrives after the customer's delay and leaves after supplier terms.")
    month_header(ws, 4, A["start"])
    idx_row(ws, 5)
    mrow(ws, 6, "Opening cash", lambda c, p, k: f"={A['cash_open']}" if k == 0 else f"={p}20", "calc", VND, None)
    mrow(ws, 7, "Receipts from customers (incl. VAT, after the payment delay)",
         lambda c, p, k: f"=IF({c}$5-{A['cust_lag']}>=1,INDEX('Sales plan'!$B$10:$M$10,{c}$5-{A['cust_lag']})+INDEX(Tax!$B$6:$M$6,{c}$5-{A['cust_lag']}),0)+IF({c}$5=1,{A['recv_open']},0)", "calc")
    mrow(ws, 8, "Suppliers and direct costs (incl. input VAT, after supplier terms)",
         lambda c, p, k: f"=IF({c}$5-{A['supp_lag']}>=1,INDEX('Sales plan'!$B$13:$M$13,{c}$5-{A['supp_lag']})+INDEX(Tax!$B$8:$M$8,{c}$5-{A['supp_lag']}),0)+IF({c}$5=1,{A['pay_open']},0)", "calc")
    mrow(ws, 9, "Warranty and finance cost", lambda c, p, k: f"='Sales plan'!{c}14", "link")
    mrow(ws, 10, "Non-payroll fixed costs (rent, marketing, other)", lambda c, p, k: f"='Fixed costs'!{c}12", "link")
    mrow(ws, 11, "Payroll (net pay + insurance + tax)", lambda c, p, k: f"=Payouts!{c}11", "link")
    mrow(ws, 12, "Bonuses and commission", lambda c, p, k: f"=Payouts!{c}14", "link")
    mrow(ws, 13, "VAT paid", lambda c, p, k: f"=Tax!{c}11", "link")
    mrow(ws, 14, "Corporate income tax (provisional)", lambda c, p, k: f"=Tax!{c}25", "link")
    mrow(ws, 15, "Loan principal and interest", lambda c, p, k: f"=Payouts!{c}23", "link")
    mrow(ws, 16, "Owner payouts (dividends, shareholder loan)", lambda c, p, k: f"=Payouts!{c}16+Payouts!{c}17", "link")
    mrow(ws, 17, "One-off investment", lambda c, p, k: f"='Fixed costs'!{c}13", "link")
    mrow(ws, 18, "Total cash out", lambda c, p, k: f"=SUM({c}8:{c}17)", "calc")
    mrow(ws, 19, "Net cash flow", lambda c, p, k: f"={c}7-{c}18", "calc")
    mrow(ws, 20, "Closing cash", lambda c, p, k: f"={c}6+{c}19", "key", VND, None)
    put(ws, "A22", "Lowest closing cash", "label"); put(ws, "B22", "=MIN(B20:M20)", "key", VND)
    put(ws, "A23", "Month of lowest cash", "label"); put(ws, "B23", "=INDEX(B4:M4,MATCH(MIN(B20:M20),B20:M20,0))", "key", MON)
    put(ws, "A24", "Minimum cash to keep", "calc"); put(ws, "B24", f"={A['buffer']}", "link", VND)
    put(ws, "A25", "Capital needed to keep that minimum", "label"); put(ws, "B25", "=MAX(0,B24-B22)", "key", VND)
    put(ws, "A27", "Opening receivables and payables are collected / paid in month 1. Retention money and advances are not modelled here (see module B04 and B03).", "note")


CUST_ROWS = [
    ("section", "BRAND"),
    ("row", "Brand name of each business line", "@B1@: contracting and fabrication. @B2@: export. @B3@: premium system supplier. Confirm the names and which customer each speaks to."),
    ("row", "Slogan", "One short line per brand. Write it yourself after reviewing the positioning below."),
    ("row", "Positioning", "State in one sentence per brand what makes it different (reliability, system performance, made to the buyer's specification). Enter the proven facts only."),
    ("row", "Customer insight", "Investors and contractors want on-time delivery, even quality and a clear price. Architects want systems that meet the design and the technical spec. Home owners want comfort, safety and a warranty."),
    ("row", "Mission", "Deliver doors and systems that are accurate, on time and backed by a clear warranty, for every segment from standard to premium."),
    ("row", "12-month goal", "Set a target for each brand: paying customers, average selling price, cost to win one customer, share of repeat or referred customers. Enter the numbers here, then link them to Assumptions."),
    ("section", "CUSTOMER PROFILE"),
    ("row", "Group 1: Investors and general contractors (@B1@)", "Project owners and main contractors. Buy on schedule, price and track record."),
    ("row", "Group 2: Architects (@B3@)", "Specify the system in the design. Need technical files, samples and test results."),
    ("row", "Group 3: Home owners (@B1@)", "High-end homes. Want safety, comfort, lasting finish and warranty."),
    ("row", "Group 4: Fabricators (@B3@)", "Workshops that buy the system and hardware, then build and install."),
    ("row", "Group 5: Overseas distributors and stores (@B2@)", "Enter the target markets. Buy in volume; need steady supply and competitive price."),
    ("row", "Main pain", "Late delivery, uneven quality, unclear price or warranty. Enter what interviews show."),
    ("row", "How they handle it today", "Enter what interviews show."),
    ("row", "Where they are", "Project tenders, architect networks, social media, trade fairs, building-material stores, B2B marketplaces."),
    ("row", "How much they will pay", "Test two or three price levels with real quotes. Needs verifying."),
    ("section", "STRATEGY"),
    ("row", "Sales channels", "Direct sales team, referrals, architect talks, show-room or site visits, social media and website, distributors and agents, trade fairs."),
    ("row", "How we win the first contact", "Samples, technical catalogue, project references, site visit with a quick estimate."),
    ("row", "How we sell", "Survey, drawings and a clear quote. Price, schedule and warranty agreed before the contract is signed. See Sales policy for commission and discount limits."),
    ("row", "How we keep customers", "Fast reply on service calls, maintenance, handover record, and a count of repeat customers per month."),
    ("section", "INTERNAL"),
    ("row", "Who does what", "Name one owner for each step: sales, technical drawings, quoting, factory, installation, accounting. See Agreements & Controls."),
    ("row", "Bonus policy", "Commission and bid-win bonus rules for sales and the technical team. See Sales policy and Payouts."),
    ("row", "After-sales service", "Set a response-time promise per brand (for example hours to reply, days to visit)."),
    ("row", "Pricing policy", "A price list per brand with discount limits, and the minimum margin that needs a director's approval. See Costing and Sales policy."),
]
CUST_ROWS = P.customers_override(CUST_ROWS)


def build_customers(wb):
    ws = wb["Customers & Strategy"]
    head(ws, "BRANDS, CUSTOMER PROFILE AND STRATEGY",
         "The yellow cells must hold real results from talking to customers, not opinions. Edit the draft column first.")
    widths(ws, [38, 74, 40])
    dark = PatternFill("solid", fgColor="2B2B2B")
    pink = PatternFill("solid", fgColor="F6DCE0")
    for i, t in enumerate(["Item", "Draft suggestion (edit)", "Fill in after customer interviews"]):
        c = put(ws, f"{L(i + 1)}4", t, "hdr", wrap=True)
        c.fill = dark
    ws.row_dimensions[4].height = 22
    r = 5
    for item in CUST_ROWS:
        if item[0] == "section":
            for col in range(1, 4):
                ws.cell(row=r, column=col).fill = pink
                ws.cell(row=r, column=col).border = BOX
            ws.cell(row=r, column=1, value=item[1]).font = Font(name=F, size=10, bold=True)
            ws.row_dimensions[r].height = 18
        else:
            put(ws, f"A{r}", item[1], "calc", wrap=True)
            put(ws, f"B{r}", item[2], "calc", wrap=True)
            put(ws, f"C{r}", None, "input", wrap=True)
            lines = max(-(-len(item[1]) // 44), -(-len(item[2]) // 92), 1)
            ws.row_dimensions[r].height = 13 * lines + 6
        r += 1
    ws.freeze_panes = "A5"


def build_todo(wb, full=True):
    ws = wb["To-do"]
    head(ws, "TO-DO", "What to find out and verify before trusting the numbers.")
    widths(ws, [6, 66, 22, 18, 14, 16, 36])
    hdr_row(ws, 4, ["No.", "Task", "Sheet", "Owner", "Deadline", "Status", "Note"])
    tasks = [("Replace every EXAMPLE number in Assumptions with a real figure and its source", "Assumptions"),
             ("Confirm all yellow rates with your tax adviser (corporate tax tier, reduced VAT, caps)", "Assumptions"),
             ("Enter opening balances from the bank statement and the ledger", "Assumptions"),
             ("Fill the yellow column from real customer interviews (needs, price, channels)", "Customers & Strategy"),
             ("Set channels, commission and discount limits", "Sales policy"),
             ("Fill the real monthly fixed costs", "Fixed costs"),
             ("Fill the due dates from the tax authority calendar", "Tax"),
             ("Agree approval limits and who signs off", "Agreements & Controls"),
             ("Check that all cross-checks show OK", "Agreements & Controls")]
    for i in range(20):
        r = 5 + i
        put(ws, f"A{r}", i + 1, "calc", "0")
        for c, v in zip("BCDEF", (tasks[i][0], tasks[i][1], None, None, "Not started") if i < len(tasks) else (None, None, None, None, "Not started")):
            put(ws, f"{c}{r}", v, "input", wrap=(c == "B"))
        put(ws, f"G{r}", None, "input", wrap=True)
    dv = DataValidation(type="list", formula1='"Not started,Doing,Done,Blocked"', allow_blank=True); ws.add_data_validation(dv); dv.add("F5:F24")
    put(ws, "B26", "Done so far", "label"); put(ws, "C26", '=COUNTIF(F5:F24,"Done")', "key", "0")
    ws.freeze_panes = "A5"


def build_policy(wb, A):
    ws = wb["Sales policy"]
    head(ws, "SALES POLICY", "Channels, commission and discount: how much of each 100 VND of price is really left.")
    widths(ws, [52, 18, 16, 16, 24, 26])
    hdr_row(ws, 4, ["Channel (EXAMPLE)", "Share of revenue", "Commission %", "Max discount %", "Net per 100 VND of list price", "Left after variable cost, per 100"])
    ch = [("Direct to investor or home owner", 0.40, 0.0, 0.05), ("Via general contractor or architect", 0.40, 0.03, 0.08), ("Via distributor or agent (export)", 0.20, 0.05, 0.10)]
    for i, (n, s, cm, d) in enumerate(ch):
        r = 5 + i
        put(ws, f"A{r}", n, "input"); put(ws, f"B{r}", s, "input", PCT); put(ws, f"C{r}", cm, "input", PCT); put(ws, f"D{r}", d, "input", PCT)
        put(ws, f"E{r}", f"=100*(1-D{r})*(1-C{r})", "calc", NUM); put(ws, f"F{r}", f"=E{r}-100*$B$9", "calc", NUM)
    put(ws, "A8", "Total share (must be 100%) and weighted result", "label"); put(ws, "B8", "=SUM(B5:B7)", "key", PCT)
    put(ws, "E8", "=IFERROR(SUMPRODUCT(B5:B7,E5:E7)/B8,0)", "key", NUM); put(ws, "F8", "=IFERROR(SUMPRODUCT(B5:B7,F5:F7)/B8,0)", "key", NUM)
    put(ws, "A9", "Variable cost as % of list price (direct + warranty + finance)", "calc"); put(ws, "B9", "=IFERROR((Costing!B9+Costing!B11+Costing!B12)/Costing!B14,0)", "link", PCT)
    hdr_row(ws, 11, ["Where each 100 VND of list price goes", "VND per 100", "", "", "", ""])
    items = [("Material", "Costing!B5"), ("Direct labour", "Costing!B6"), ("Installation", "Costing!B7"), ("Transport and packing", "Costing!B8"),
             ("Factory overhead", "Costing!B10"), ("Warranty reserve", "Costing!B11"), ("Finance cost", "Costing!B12")]
    for i, (n, ref) in enumerate(items):
        r = 12 + i
        put(ws, f"A{r}", n, "calc"); put(ws, f"B{r}", f"=IFERROR({ref}/Costing!$B$14*100,0)", "link", NUM)
    put(ws, "A19", "Left for fixed costs and profit", "label"); put(ws, "B19", "=100-SUM(B12:B18)", "key", NUM)
    put(ws, "A21", "Commission paid to individuals is taxable income for them: check withholding (module A03).", "note")


def build_controls(wb, entity, module=None):
    ws = wb["Agreements & Controls"]
    head(ws, "AGREEMENTS AND CONTROLS", "Agree the rules before the money moves. Cross-checks must all show OK.")
    widths(ws, [60, 26, 18, 18, 44, 18])
    r = 4
    bar(ws, r, "1. OWNER / PARTNER AGREEMENT (WRITE IT BEFORE THE FIRST SALE)", 6); r += 1
    hdr_row(ws, r, ["Rule", "Agreed text", "Agreed by", "Date", "Note"]); r += 1
    for t in ["Capital each partner puts in", "How profit is split", "Who does what (roles and effort)", "Who bears a loss, and how much",
              "Owner salary versus dividend versus shareholder loan", "When we stop (stop criteria)"]:
        put(ws, f"A{r}", t, "calc")
        for c in "BCDE":
            put(ws, f"{c}{r}", None, "input", wrap=True)
        r += 1
    r += 1
    bar(ws, r, "2. APPROVAL LIMITS", 6); r += 1
    hdr_row(ws, r, ["Payment up to (VND)", "Approver", "Document needed", "", "Note"]); r += 1
    for amt, who, doc in ((5000000, "Department head", "Request + invoice"), (50000000, "Finance manager", "PO + invoice + delivery note"), (500000000, "Director", "Contract + PO + acceptance")):
        put(ws, f"A{r}", amt, "input", VND); put(ws, f"B{r}", who, "input"); put(ws, f"C{r}", doc, "input", wrap=True)
        put(ws, f"D{r}", None, "calc"); put(ws, f"E{r}", "EXAMPLE limits: set your own", "note"); r += 1
    r += 1
    bar(ws, r, "3. CROSS-CHECKS (DIFFERENCE MUST BE ZERO)", 6); r += 1
    hdr_row(ws, r, ["Check", "Difference", "Tolerance", "Status", "What a failure means"]); r += 1
    first = r
    checks = [("Payroll cash out = gross payroll + employer insurance", "=Payouts!N11-(Payouts!N6+Payouts!N10)", 1, "Payroll split is wrong"),
              ("Employer insurance: Fixed costs = Payouts", "='Fixed costs'!N7-Payouts!N10", 1, "Two sheets use different insurance rates"),
              ("Closing cash = opening cash + net cash flow", "='Cash flow'!M20-('Cash flow'!B6+'Cash flow'!N19)", 1, "Cash flow rows do not add up"),
              ("Revenue = domestic + export", "='Sales plan'!N10-'Sales plan'!N11-'Sales plan'!N12", 1, "Export share is broken"),
              ("Channel shares add up to 100%", "='Sales policy'!B8-1", 0.0001, "Channel shares in Sales policy are wrong")]
    for t, f, tol, why in checks:
        put(ws, f"A{r}", t, "calc"); put(ws, f"B{r}", f, "calc", '#,##0.00'); put(ws, f"C{r}", tol, "calc", '0.0000')
        put(ws, f"D{r}", f'=IF(ABS(B{r})<=C{r},"OK","CHECK")', "key"); put(ws, f"E{r}", why, "note"); r += 1
    put(ws, f"A{r}", "All cross-checks", "label"); put(ws, f"D{r}", f'=IF(COUNTIF(D{first}:D{r - 1},"CHECK")=0,"ALL OK","CHECK")', "key")
    all_ok = f"'Agreements & Controls'!D{r}"
    r += 2
    bar(ws, r, "4. SIGN-OFF", 6); r += 1
    hdr_row(ws, r, ["Item", "Name", "Date", "Version", "Note"]); r += 1
    for t in ["Prepared by", "Reviewed by", "Approved by"]:
        put(ws, f"A{r}", t, "calc")
        for c in "BCDE":
            put(ws, f"{c}{r}", None, "input")
        r += 1
    put(ws, f"A{r}", "No sign-off means the file is a draft.", "note"); r += 2
    bar(ws, r, "5. CHANGE LOG (EVERY CHANGE TO AN ASSUMPTION)", 6); r += 1
    hdr_row(ws, r, ["Date and who", "Sheet and cell", "Old value", "New value", "Reason"]); r += 1
    for _ in range(8):
        for c in "ABCDE":
            put(ws, f"{c}{r}", None, "input", wrap=True)
        r += 1
    return all_ok


CANVAS_TITLES = {1: "1. CUSTOMER SEGMENTS", 2: "2. VALUE WE OFFER", 3: "3. CHANNELS", 4: "4. CUSTOMER RELATIONSHIPS", 5: "5. REVENUE STREAMS",
                 6: "6. KEY ACTIVITIES", 7: "7. KEY RESOURCES (what we need)", 8: "8. KEY PARTNERS (who helps us)", 9: "9. COST STRUCTURE"}

# Calm palette for the Canvas: white body (60%), charcoal bars (30%), one soft grey-blue accent for numbers (10%). No bright fills.
CV_BAR, CV_TEXT, CV_LINE, CV_SOFT = "3A3F47", "1F2933", "D5D9DE", "EEF1F5"
_cv_side = Side(style="thin", color=CV_LINE)
CV_BOX = Border(left=_cv_side, right=_cv_side, top=_cv_side, bottom=_cv_side)

CANVAS_FULL = {
    1: ("- Investors and developers of residential and industrial projects\n- General contractors and main contractors (tenders)\n"
        "- Architects and design consultants who specify the system\n- Owners of villas, townhouses and high-end homes\n"
        "- Fabricators that want a premium system (@B3@ customers; @B1@ may be one of them)\n"
        "- Overseas distributors, building-material stores and contractors (@B2@): enter the target markets\n"
        "- Next: other markets or product lines (enter)\n> Detail: Customers & Strategy"),
    2: ("- Doors, windows, glass facades and partitions: design, production, installation and one warranty\n"
        "- A range for every budget: standard (@STD@), mid-range (@JP@) and premium thermal-break (@B3@)\n"
        "- @B3@: premium system with certified performance (enter the proven figures)\n"
        "- Quality hardware, airtight and watertight\n"
        "- Proven residential and industrial references\n"
        "- Export (@B2@): made to the buyer's specification, shipped on schedule\n"
        "- Clear price, schedule and warranty before the contract is signed\n> Detail: Costing, Break-even"),
    3: ("- Direct sales team, referrals, and tenders with general contractors\n- Architect talks, samples and showroom\n"
        "- Social media, website and SEO content\n- Distributors, agents, trade fairs and B2B email abroad\n"
        "- @B3@ sells the system to fabricators only (no fabrication, installation or glass sales)\n> Detail: Sales policy"),
    4: ("- Site survey and technical advice before quoting\n- One project manager per project, regular progress reports\n"
        "- Handover with acceptance record and warranty card\n- Maintenance and fast reply to complaints\n"
        "- Architects: samples, technical files, detail drawings\n- Fabricators: training and system support\n"
        "- Distributors: price list, stock plan, regular contact\n> Detail: Sales policy"),
    6: ("- Survey, design, shop drawings and quoting\n- Bidding and contract negotiation\n- Purchasing profile, glass and imported hardware\n"
        "- Production and quality check in the factory\n- Delivery, installation and site management\n"
        "- Acceptance, collecting payment, warranty\n- Marketing and content\n"
        "- Accounting: monthly close, tax filing, cash forecast"),
    7: ("- Factory, machines and staff (technical, workshop, sales, accounting): enter headcount\n- System licence and hardware supply (@SUP@)\n"
        "- Working capital to buy materials before customers pay\n- Project references and brand\n"
        "- Accounting and tax know-how (this workbook)"),
    8: ("- @SUP@: hardware and technical know-how\n- Profile, glass and accessory suppliers\n- Installation crews and subcontractors\n"
        "- Banks: credit line, guarantees, letters of credit\n- Freight forwarders, customs agent and insurers\n- Accountant and tax adviser\n"
        "- Software and marketing partners"),
}
CANVAS_FULL = P.canvas_override(CANVAS_FULL)


def _cv_dynamic(A):
    """Revenue and cost bullets pull live numbers from the model, so the canvas cannot drift away from the sheets."""
    def t(ref, fmt="#,##0"):
        return f'TEXT({ref},"{fmt}")'
    rev = ['="- Price per unit: "&' + t(A["price"]) + '&" VND (example)"',
           '="- Revenue, 12 months: "&' + t("'Sales plan'!N10") + '&" VND"',
           "- Project payments: advance, progress payments, acceptance, retention",
           "- System and hardware sales to fabricators (@B3@)",
           "- Export orders in USD: enter payment terms (advance or L/C)",
           "- Service and warranty work",
           "> Detail: Sales plan, Cash flow"]
    cost = ['="- Fixed cost per month: "&' + t("'Break-even'!B10") + '&" VND (rent, salaries, insurance, marketing, loan interest)"',
            '="- Direct cost per unit: "&' + t("Costing!B9") + '&" VND (material, labour, install, transport)"',
            '="- Margin per unit after overhead, warranty and finance: "&' + t("Costing!B15", "0.0%"),
            '="- Taxes and insurance, 12 months: "&' + t("Tax!N11+Tax!N25+Tax!N31") + '&" VND"',
            "- Largest single cost: material. Import duty and VAT on imported hardware are cash out before the sale",
            "> Detail: Fixed costs, Costing, Tax"]
    return rev, cost


def _bullets(v):
    return v if isinstance(v, list) else [x for x in str(v).split("\n") if x]


def _joined(lines):
    """All bullets of a box as ONE cell value. Plain text stays text; if any bullet is a live formula, build one formula."""
    if not any(str(x).startswith("=") for x in lines):
        return "\n".join(lines)
    parts = [x[1:] if str(x).startswith("=") else '"' + str(x).replace('"', '""') + '"' for x in lines]
    return "=" + "&CHAR(10)&".join(parts)


def canvas_grid(ws, content, width=52):
    """The company's standard 9-block canvas (layout from the company's own sample): white boxes, charcoal title bars.
    Each box is ONE cell holding all its bullets, with Wrap Text on, so long lines drop to the next line by themselves."""
    widths(ws, [width] * 5)
    PER = 1.05  # characters per width unit, Arial 10
    LINE = 12.75

    def height(n, ncols):
        lines = 0
        for t in _bullets(content.get(n, "")):
            chars = 85 if str(t).startswith("=") else len(str(t))
            lines += max(1, -(-chars // int(width * ncols * PER)))
        return lines * LINE + 8

    # (block, title_row, first_content_row, last_content_row, col_start, col_end)
    spec = [(8, 4, 5, 12, 1, 1), (6, 4, 5, 8, 2, 2), (7, 9, 10, 12, 2, 2), (2, 4, 5, 12, 3, 3), (4, 4, 5, 8, 4, 4),
            (3, 9, 10, 12, 4, 4), (1, 4, 5, 12, 5, 5), (9, 14, 15, 18, 1, 2), (5, 14, 15, 18, 3, 5)]
    for n, tr, r1, r2, c1, c2 in spec:
        for col in range(c1, c2 + 1):
            ws.cell(row=tr, column=col).fill = PatternFill("solid", fgColor=CV_BAR)
        t = ws.cell(row=tr, column=c1, value=CANVAS_TITLES[n])
        t.font = Font(name=F, size=10, bold=True, color="FFFFFF")
        t.alignment = Alignment(horizontal="center", vertical="center")
        if c2 > c1:
            ws.merge_cells(start_row=tr, start_column=c1, end_row=tr, end_column=c2)
        c = ws.cell(row=r1, column=c1, value=_joined(_bullets(content.get(n, ""))))
        c.font = Font(name=F, size=10, color=CV_TEXT)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if r2 > r1 or c2 > c1:
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        for r in range(r1, r2 + 1):
            for col in range(c1, c2 + 1):
                ws.cell(row=r, column=col).border = Border(left=_cv_side if col == c1 else None, right=_cv_side if col == c2 else None,
                                                           bottom=_cv_side if r == r2 else None)
    need = {n: height(n, c2 - c1 + 1) for n, _, _, _, c1, c2 in spec}
    top_a = max(need[6], need[4]) / 4          # rows 5-8
    low_a = max(need[7], need[3]) / 3          # rows 10-12
    h58, h1012 = top_a, low_a
    full = max(need[8], need[2], need[1])
    if 4 * h58 + 22 + 3 * h1012 < full:        # tall boxes (8, 2, 1) need more room: spread the shortfall
        extra = (full - (4 * h58 + 22 + 3 * h1012)) / 7
        h58, h1012 = h58 + extra, h1012 + extra
    for r in (4, 9, 14):
        ws.row_dimensions[r].height = 22
    for r in range(5, 9):
        ws.row_dimensions[r].height = h58
    for r in range(10, 13):
        ws.row_dimensions[r].height = h1012
    for r in range(15, 19):
        ws.row_dimensions[r].height = max(need[9], need[5]) / 4


def key_numbers(ws, start, rows):
    for c in range(1, 6):
        ws.cell(row=start, column=c).fill = PatternFill("solid", fgColor=CV_BAR)
    x = ws.cell(row=start, column=1, value="KEY NUMBERS (LINKED FROM THE OTHER TABS)")
    x.font = Font(name=F, size=10, bold=True, color="FFFFFF")
    ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=5)
    for i, (b, f, fmt, what) in enumerate(rows):
        r = start + 1 + i
        a = ws.cell(row=r, column=1, value=b)
        a.font, a.border = Font(name=F, size=10, color=CV_TEXT), CV_BOX
        v = ws.cell(row=r, column=2, value=f)
        v.font, v.border = Font(name=F, size=10, bold=True, color="008000" if f and str(f).startswith("=") else CV_TEXT), CV_BOX
        v.fill = PatternFill("solid", fgColor=CV_SOFT)
        v.alignment = Alignment(horizontal="right")
        if fmt:
            v.number_format = fmt
        d = ws.cell(row=r, column=3, value=what)
        d.font, d.border = Font(name=F, size=9, italic=True, color="595959"), CV_BOX
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        for cc in (4, 5):
            ws.cell(row=r, column=cc).border = CV_BOX
    return start + 1 + len(rows)


def build_canvas(wb, A, all_ok):
    ws = wb["Canvas"]
    head(ws, "BUSINESS MODEL CANVAS (DRAFT TO EDIT)",
         "See the whole model on one page before touching the numbers. Fill in this order: 1 customers, 2 value, 3 channels, 4 relationships, 5 revenue, 6 activities, 7 resources, 8 partners, 9 costs. Boxes 5 and 9 update from the model.")
    content = dict(CANVAS_FULL)
    content[5], content[9] = _cv_dynamic(A)
    canvas_grid(ws, content)
    rows = [("Units sold, 12 months", "='Sales plan'!N8", '#,##0', "Scenario in use (Assumptions)"),
            ("Price per unit (VND)", f"={A['price']}", VND, "Selling price"),
            ("Net per 100 VND of list price", "='Sales policy'!E8", NUM, "After discount and commission"),
            ("Revenue, 12 months (VND)", "='Sales plan'!N10", VND, ""),
            ("Direct cost per unit (VND)", "=Costing!B9", VND, "Material + labour + install + transport"),
            ("Margin per unit", "=Costing!B15", PCT, "After overhead, warranty and finance"),
            ("Fixed cost per month (VND)", "='Break-even'!B10", VND, "Including loan interest"),
            ("Break-even units per month", "='Break-even'!B11", '#,##0', "Below this you lose money"),
            ("Lowest cash in 12 months (VND)", "='Cash flow'!B22", VND, "Negative means you need funding"),
            ("Capital needed (VND)", "='Cash flow'!B25", VND, "To keep the minimum cash"),
            ("Taxes and insurance paid (VND)", "=Tax!N11+Tax!N25+Tax!N31", VND, "VAT + corporate tax + payroll taxes and insurance"),
            ("Owner payouts (VND)", "=Payouts!N16+Payouts!N17", VND, "Dividends and shareholder loan repaid"),
            ("Profit before tax (VND)", "=Tax!N20", VND, "12 months")]
    end = key_numbers(ws, 20, rows)
    lowest_row = 20 + 1 + [i for i, r in enumerate(rows) if r[0].startswith("Lowest cash")][0]
    for lab, f, off in (("Model health", f'=IF({all_ok}="ALL OK","All cross-checks OK","CHECK: see Agreements & Controls")', 1),
                        ("Cash status", f'=IF(B{lowest_row}<0,"Cash goes negative: raise capital or change terms","Cash stays positive")', 2),
                        ("Entity", f"={A['entity']}", 3)):
        r = end + off
        a = ws.cell(row=r, column=1, value=lab); a.font, a.border = Font(name=F, size=10, bold=True, color=CV_TEXT), CV_BOX
        v = ws.cell(row=r, column=2, value=f); v.font, v.border = Font(name=F, size=10, bold=True, color="008000"), CV_BOX
        v.fill = PatternFill("solid", fgColor=CV_SOFT)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)


def full_questions(A):
    p = plain
    return [f"Lower {p(A['price'])} (selling price) by 10%. What happens to Break-even!B11 and to profit on the Canvas? Why does profit fall by more than 10%?",
            f"Set {p(A['cust_days'])} (customer payment days) from 30 to 90. Does profit before tax change? What happens to the lowest cash ('Cash flow'!B22) and the capital needed?",
            "Double the rent in 'Fixed costs'!B5. How many extra units per month are needed to break even (Break-even!B11)?",
            f"Switch {p(A['scenario'])} to Worst. Which month has the lowest cash and how much capital is needed ('Cash flow'!B23 and B25)?",
            f"Raise {p(A['material'])} (material cost) by 8% and keep the price. What is the new margin (Costing!B15) and what price gives the target margin (Costing!B17)?",
            f"Set {p(A['export_share'])} (export share) to 50%. What happens to output VAT in the Tax sheet and to cash? Why is input VAT still paid?",
            f"Raise {p(A['bonus_pct'])} (bonus % of gross profit) from 3% to 6%. What changes in profit, tax and cash?",
            f"Change {p(A['cit'])} from 20% to 17%. How much corporate tax is saved in the year and in which months is it paid?",
            "A customer cancels after you bought the materials. Where does the loss show up: revenue, cost, VAT or cash? Which sheets change?"]


# -------------------------------------------------------------- full model
def build_full(wb, entity, ref):
    for name, _, _ in FULL_ORDER:
        wb.create_sheet(name)
    A, verify = build_assumptions(wb, ref, entity)
    build_costing(wb, A); build_fixed(wb, A); build_sales_plan(wb, A); build_payouts(wb, A); build_tax(wb, A)
    build_breakeven(wb, A); build_cashflow(wb, A); build_policy(wb, A); build_customers(wb); build_todo(wb)
    all_ok = build_controls(wb, entity)
    build_canvas(wb, A, all_ok)
    return FULL_ORDER, full_questions(A), verify


# ------------------------------------------------------------ module files
def build_module(wb, entity, ref, module, tax_ids):
    """Canvas + Assumptions (module numbers + sourced rates) + module sheets + Agreements & Controls (risk list)."""
    own = [s for s in module["sheets"] if s not in ("Canvas", "Assumptions", "Agreements & Controls")]
    order = [("Canvas", f"{module['why']} One page: purpose, inputs, result and open risks.", True),
             ("Assumptions", "Numbers this file needs (enter first) and the sourced tax / payroll rates with date and confidence. Every other sheet reads from here.", True)]
    for s in own:
        order.append((s, "Working sheet for this module. Replace the example header with the columns you need.", True if len(order) < 5 else False))
    order.append(("Agreements & Controls", "Taxes, payouts and risks to watch for this file, with owner and status; sign-off and change log.", False))
    for n, _, _ in order:
        wb.create_sheet(n)
    # Assumptions
    ws = wb["Assumptions"]
    head(ws, "ASSUMPTIONS", f"Numbers for {module['id']}: {module['name']}. Blue/yellow cells are yours. Yellow in the last column = confirm first.")
    widths(ws, [58, 30, 12, 64, 12, 70])
    hdr_row(ws, 4, ["Item", "Value", "Unit", "Source", "Checked on", "Confidence / note"])
    r = 5
    bar(ws, r, "GENERAL", 6); r += 1
    for lab, v in (("Entity", entity), ("Period start", None), ("Prices include VAT? (Yes / No)", "No"), ("Currency (VND / USD)", "VND")):
        put(ws, f"A{r}", lab, "calc"); put(ws, f"B{r}", v, "input"); [put(ws, f"{c}{r}", None, "calc") for c in "CDEF"]; r += 1
    bar(ws, r, "NUMBERS TO ENTER FIRST (THIS FILE)", 6); r += 1
    for n in module["numbers"]:
        put(ws, f"A{r}", n, "calc"); put(ws, f"B{r}", None, "input"); put(ws, f"C{r}", None, "calc")
        put(ws, f"D{r}", "Enter source: contract / invoice / bank statement", "calc"); put(ws, f"E{r}", None, "calc"); put(ws, f"F{r}", None, "calc"); r += 1
    verify = []
    if tax_ids:
        bar(ws, r, "TAX AND PAYROLL RATES (SOURCED; YELLOW = CONFIRM WITH YOUR TAX ADVISER)", 6); r += 1
        for tid in tax_ids:
            it = ref[tid]
            ver = any(w in it["confidence"].lower() for w in ("verify", "single source"))
            put(ws, f"A{r}", it["topic"], "calc", wrap=True); put(ws, f"B{r}", it["value"], "input", wrap=True); put(ws, f"C{r}", None, "calc")
            put(ws, f"D{r}", it["source"], "calc", wrap=True); put(ws, f"E{r}", it["checked_on"], "calc")
            c = put(ws, f"F{r}", it["confidence"] + (" | " + it["note"] if it.get("note") else ""), "calc", wrap=True)
            ws.row_dimensions[r].height = 58
            if ver:
                c.fill = FILL_KEY; verify.append(it["topic"])
            r += 1
    ws.freeze_panes = "B5"
    # module sheets
    for s in own:
        w = wb[s]
        head(w, s.upper(), f"{module['id']} {module['name']}. Replace the headers below with the columns this sheet needs.")
        widths(w, [14, 22, 40, 12, 20, 14, 36])
        hdr_row(w, 4, ["Date", "Reference", "Description", "Currency", "Amount", "Status", "Note"])
        for i in range(30):
            for j in range(7):
                put(w, f"{L(j + 1)}{5 + i}", ["01/01/2027", "REF-001", "EXAMPLE row: delete or overwrite", "VND", 0, "Open", None][j] if i == 0 else None, "input", VND if j == 4 else None)
        w.freeze_panes = "A5"
    # controls (risk list)
    c = wb["Agreements & Controls"]
    head(c, "AGREEMENTS AND CONTROLS", f"Taxes, payouts and risks to watch for {module['id']}. Yellow = not confirmed.")
    widths(c, [6, 90, 14, 22, 36])
    bar(c, 4, "TAXES, PAYOUTS AND RISKS TO WATCH", 5)
    hdr_row(c, 5, ["No.", "Risk / tax / payout", "Status", "Owner", "Note"])
    dv = DataValidation(type="list", formula1='"Open,Checked,Not relevant"', allow_blank=True); c.add_data_validation(dv)
    for i, wtxt in enumerate(module["watch"]):
        r = 6 + i
        put(c, f"A{r}", i + 1, "calc", "0"); cell = put(c, f"B{r}", wtxt, "calc", wrap=True)
        if "verify" in wtxt.lower():
            cell.fill = FILL_KEY
        put(c, f"C{r}", "Open", "input"); put(c, f"D{r}", None, "input"); put(c, f"E{r}", None, "input", wrap=True)
        dv.add(f"C{r}"); c.row_dimensions[r].height = 30
    last = 5 + len(module["watch"])
    r = last + 2
    for lab in ("Prepared by", "Reviewed by", "Approved by"):
        put(c, f"B{r}", lab, "label"); put(c, f"C{r}", None, "input"); put(c, f"D{r}", None, "input"); r += 1
    put(c, f"B{r}", "No sign-off means the file is a draft. Log every assumption change below (date, who, cell, old, new, reason).", "note")
    # canvas: the same 9-block layout, prefilled with this module's purpose and inputs
    cv = wb["Canvas"]
    head(cv, f"CANVAS: {module['id']} {module['name'].upper()}", "Same 9-block layout as the company model. Blue text = your answers. Edit before use.")
    prompts = {1: "Who uses this file? (names / roles)", 2: module["why"], 3: "Where does the data come from? (bank, ERP, invoices, customs)",
               4: "Who reviews the result and how often?", 5: "What money does this file protect or recover?",
               6: "What do you do in this file each period?", 7: "Inputs needed:\n- " + "\n- ".join(module["numbers"]),
               8: "Who must give you data or sign off? (accountant, bank, customs agent)", 9: "What does it cost to keep this file right? (time, tools, penalties if wrong)"}
    canvas_grid(cv, prompts)
    key_numbers(cv, 20, [("Risks still open", f"=COUNTIF('Agreements & Controls'!C6:C{last},\"Open\")", '0', "Rows with status Open in Agreements & Controls"),
                         ("Key result 1", None, VND, "Enter the formula that answers the decision"),
                         ("Key result 2", None, VND, "Enter the formula that answers the decision")])
    scen = list(module["scenarios"]) + [
        "Which input in Assumptions would change your decision the most if it were 20% wrong? How would you find out the real value?",
        "Which risk in Agreements & Controls could cost the most cash, and who owns it?",
        "Which numbers are yellow (not confirmed)? What could go wrong if you used them as they are?"]
    return order, scen, verify
