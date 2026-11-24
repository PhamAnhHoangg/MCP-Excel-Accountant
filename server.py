"""Excel Accounting MCP server.

Safe, auditable Excel tools for accounting work.
- All file access is confined to one workspace folder (EXCEL_MCP_ROOT).
- Every write makes a timestamped backup first and appends to an audit log.
- Add new tools at the bottom of the TOOLS section; bump VERSION when you do.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
from datetime import datetime, date
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

try:  # mcp >= 2
    from mcp.server.mcpserver import MCPServer as _Server
except ImportError:  # mcp 1.x
    from mcp.server.fastmcp import FastMCP as _Server

VERSION = "0.8.4"
ROOT = Path(os.environ.get("EXCEL_MCP_ROOT", Path.home() / "excel-workspace")).expanduser().resolve()
BACKUP_DIR = ROOT / ".backups"
AUDIT_LOG = ROOT / ".audit.jsonl"
MAX_READ_CELLS = 5000
ROOT.mkdir(parents=True, exist_ok=True)

mcp = _Server("excel-accounting")


# ----------------------------------------------------------------- helpers
def _path(file: str, must_exist: bool = True) -> Path:
    """Resolve `file` inside ROOT; refuse anything that escapes it."""
    p = (ROOT / file).resolve()
    if p != ROOT and ROOT not in p.parents:
        raise ValueError(f"Path outside workspace is not allowed: {file}")
    if p.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Only .xlsx / .xlsm files are supported")
    if must_exist and not p.exists():
        raise FileNotFoundError(f"File not found in workspace: {file}")
    return p


def _backup(p: Path) -> str | None:
    if not p.exists():
        return None
    BACKUP_DIR.mkdir(exist_ok=True)
    dest = BACKUP_DIR / f"{datetime.now():%Y%m%d-%H%M%S}__{p.name}"
    shutil.copy2(p, dest)
    return dest.name


def _audit(action: str, file: str, detail: Any = None, backup: str | None = None) -> None:
    rec = {"ts": datetime.now().isoformat(timespec="seconds"), "action": action,
           "file": file, "backup": backup, "detail": detail}
    with AUDIT_LOG.open("a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")


def _jsonable(v: Any) -> Any:
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    return v


def _sheet(wb, name: str):
    if name not in wb.sheetnames:
        raise ValueError(f"Sheet '{name}' not found. Available: {wb.sheetnames}")
    return wb[name]


def _df(file: str, sheet: str, header_row: int = 1) -> pd.DataFrame:
    return pd.read_excel(_path(file), sheet_name=sheet, header=header_row - 1)


# ------------------------------------------------------------------- TOOLS
@mcp.tool()
def server_info() -> dict:
    """Server version, workspace folder and tool health. Call first when unsure."""
    return {"version": VERSION, "workspace": str(ROOT),
            "files": len(list(ROOT.glob("*.xls*"))), "backups": len(list(BACKUP_DIR.glob("*"))) if BACKUP_DIR.exists() else 0}


@mcp.tool()
def list_workbooks(subfolder: str = "") -> list[str]:
    """List .xlsx/.xlsm files in the workspace (optionally inside a subfolder)."""
    base = (ROOT / subfolder).resolve()
    if base != ROOT and ROOT not in base.parents:
        raise ValueError("Path outside workspace is not allowed")
    return sorted(str(p.relative_to(ROOT)) for p in base.rglob("*")
                  if p.suffix.lower() in {".xlsx", ".xlsm"} and BACKUP_DIR not in p.parents)


@mcp.tool()
def list_sheets(file: str) -> list[dict]:
    """Sheet names with used dimensions."""
    wb = openpyxl.load_workbook(_path(file), read_only=False)
    return [{"sheet": ws.title, "dimensions": ws.dimensions, "rows": ws.max_row, "cols": ws.max_column}
            for ws in wb.worksheets]


@mcp.tool()
def read_range(file: str, sheet: str, cell_range: str = "A1:J30", formulas: bool = False) -> dict:
    """Read a range (e.g. 'A1:F40'). formulas=True returns formulas instead of cached values."""
    wb = openpyxl.load_workbook(_path(file), data_only=not formulas)
    ws = _sheet(wb, sheet)
    cells = ws[cell_range]
    if not isinstance(cells, tuple):
        cells = ((cells,),)
    elif cells and not isinstance(cells[0], tuple):
        cells = (cells,)
    n = sum(len(r) for r in cells)
    if n > MAX_READ_CELLS:
        raise ValueError(f"Range has {n} cells; limit is {MAX_READ_CELLS}. Read in smaller chunks or use summary tools.")
    rows = [[_jsonable(c.value) for c in r] for r in cells]
    return {"sheet": sheet, "range": cell_range, "rows": rows}


@mcp.tool()
def write_cells(file: str, sheet: str, cells: dict[str, Any]) -> dict:
    """Write values/formulas to cells, e.g. {"B2": 1500000, "C2": "=B2*0.1"}. Backs up first."""
    p = _path(file)
    bk = _backup(p)
    wb = openpyxl.load_workbook(p)
    ws = _sheet(wb, sheet)
    for addr, val in cells.items():
        ws[addr] = val
    wb.save(p)
    _audit("write_cells", file, {"sheet": sheet, "cells": cells}, bk)
    return {"written": len(cells), "backup": bk}


@mcp.tool()
def append_rows(file: str, sheet: str, rows: list[list[Any]]) -> dict:
    """Append rows after the last used row. Backs up first."""
    p = _path(file)
    bk = _backup(p)
    wb = openpyxl.load_workbook(p)
    ws = _sheet(wb, sheet)
    for r in rows:
        ws.append(r)
    wb.save(p)
    _audit("append_rows", file, {"sheet": sheet, "count": len(rows)}, bk)
    return {"appended": len(rows), "last_row": ws.max_row, "backup": bk}


import sys
sys.path.insert(0, str(Path(__file__).parent))
import accounting_templates as T  # noqa: E402
import company_profile as P  # noqa: E402

GUIDE_SHEET = T.GUIDE_SHEET
GENERIC_NAMES = {"summary", "sheet1", "sheet2", "data", "main", "overview", "misc", "other"}
DEFAULT_QUESTIONS = [
    "If revenue drops 10%, which sheet changes first and by how much? Why does profit fall by more than 10%?",
    "Which month has the lowest cash balance? How much capital is needed to avoid going negative, even if the company is profitable?",
    "A customer has received goods but not paid yet: how do revenue, profit and cash flow each change differently?",
]


@mcp.tool()
def create_workbook(file: str, sheets: dict[str, list[list[Any]]] | None = None, title: str = "", subtitle: str = "",
                    sheet_guide: dict[str, str] | None = None, must_have: list[str] | None = None,
                    scenario_questions: list[str] | None = None, extend_note: str = "", guide: bool = True,
                    template: str = "", module_id: str = "", entity: str = "",
                    tax_refs: list[str] | None = None) -> dict:
    """Create a new workbook. Refuses to overwrite. Layout rule: first tab = 'Guide', then the real working tabs in
    reading order, starting 'Canvas' then 'Assumptions'. The Guide LISTS those tabs (click to open) and highlights
    only the most important ones; all detail (numbers, taxes, payouts, controls, risks) lives INSIDE the tabs.
    Never use a generic tab such as 'Summary'.
    Three ways to build:
      template="full"  -> complete 13-tab company model (Canvas, Customers & Strategy, To-do, Assumptions, Sales policy,
                          Fixed costs, Sales plan, Break-even, Costing, Cash flow, Tax, Payouts, Agreements & Controls),
                          all linked by formulas, with sourced VN rates. `entity` = the legal entity name (default: first entity in the private company profile).
      module_id="D03"  -> catalog module (see accounting_catalog): Canvas, Assumptions (module numbers + sourced rates),
                          the module's own tabs, Agreements & Controls (tax / payout / risk watch-list). Override rates with `tax_refs`.
      sheets={...}     -> custom: {"Name": [[header], [row]...]}. Pass `sheet_guide` {tab: purpose} for EVERY tab and
                          `must_have` = the 3-6 most important tab names. Put 'Canvas' first and 'Assumptions' second.
    Always pass `title`, `subtitle` and 5-8 file-specific `scenario_questions` (what-if changes quoting real cell addresses).
    guide=False only for scratch files."""
    p = _path(file, must_exist=False)
    if p.exists():
        raise FileExistsError(f"{file} already exists; choose another name")
    p.parent.mkdir(parents=True, exist_ok=True)
    entity = entity or P.default_entity()
    warnings: list[str] = []
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ref = {i["id"]: i for i in _load("vn_tax_reference.json")["items"]}
    unknown = [i for i in (tax_refs or []) if i not in ref]
    if unknown:
        raise ValueError(f"Unknown tax_refs {unknown}. Use tax_reference to list ids.")
    verify: list[str] = []
    if template:
        if template != "full":
            raise ValueError("template must be 'full'")
        order, questions, verify = T.build_full(wb, entity, ref)
        questions = scenario_questions or questions
    elif module_id:
        found = [m for m in _load("accounting_catalog.json")["modules"] if m["id"].lower() == module_id.lower()]
        if not found:
            raise ValueError(f"Unknown module_id {module_id}. Use accounting_catalog to list modules.")
        module = found[0]
        ids = tax_refs if tax_refs is not None else _module_tax_refs(module)
        order, questions, verify = T.build_module(wb, entity, ref, module, ids)
        questions = scenario_questions or questions
        title = title or f"{module['id']} {module['name']}"
    else:
        if not sheets:
            raise ValueError("Pass template='full', a module_id, or sheets={...}")
        for name, rows in sheets.items():
            if name.lower() in GENERIC_NAMES:
                warnings.append(f"tab '{name}' has a generic name: use a real, specific name")
            ws = wb.create_sheet(name[:31])
            for r in rows:
                ws.append(r)
            for i, col in enumerate(ws.columns, 1):
                width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 50)
        names = list(sheets)
        if names[0].lower() != "canvas":
            warnings.append("first tab should be 'Canvas' (one page showing the whole file)")
        if len(names) < 2 or names[1].lower() != "assumptions":
            warnings.append("second tab should be 'Assumptions' (every input, with source and date)")
        sg = sheet_guide or {}
        if not sheet_guide:
            warnings.append("sheet_guide missing: Guide shows '(no description yet)' for every tab")
        if not must_have:
            warnings.append("must_have missing: pass the 3-6 most important tab names so the Guide can highlight them")
        mh = {x.lower() for x in (must_have or names[:2])}
        order = [(n, sg.get(n, "(no description yet - write this tab's purpose)"), n.lower() in mh) for n in names]
        questions = scenario_questions or DEFAULT_QUESTIONS
        if not scenario_questions:
            warnings.append("scenario_questions missing: only 3 generic questions were added; write file-specific ones")
    if guide:
        T.add_guide_sheet(wb, title or Path(file).stem.replace("-", " ").replace("_", " "), subtitle, order, questions, extend_note)
    from openpyxl.worksheet.properties import PageSetupProperties
    for w in wb.worksheets:  # print every tab landscape, one page wide
        w.page_setup.orientation = "landscape"
        w.page_setup.fitToWidth, w.page_setup.fitToHeight = 1, 0
        w.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    if verify:
        warnings.append(f"{len(verify)} rate(s) are yellow (not fully confirmed): {verify}")
    P.apply_to_workbook(wb)  # replace @TOKENS@ with the private profile's real names
    wb.save(p)
    _audit("create_workbook", file, {"sheets": wb.sheetnames, "template": template, "module": module_id})
    return {"created": file, "sheets": wb.sheetnames, "warnings": warnings}


@mcp.tool()
def check_conventions(file: str) -> dict:
    """Check a workbook against the company setup rules: guide sheet first and describing every sheet,
    hard-coded numbers inside formulas (e.g. =B5*1.05), and cross-sheet formulas. Read-only."""
    import re
    wb = openpyxl.load_workbook(_path(file))
    names = wb.sheetnames
    res: dict[str, Any] = {"guide_first": names[0] == GUIDE_SHEET, "sheets": names}
    undescribed, hard = [], []
    if GUIDE_SHEET in names:
        gtxt = " ".join(str(c.value) for row in wb[GUIDE_SHEET].iter_rows() for c in row if c.value)
        undescribed = [n for n in names if n != GUIDE_SHEET and n not in gtxt]
    res["sheets_not_in_guide"] = undescribed
    res["canvas_second"] = len(names) > 1 and names[1].lower() == "canvas"
    res["generic_tab_names"] = [n for n in names if n.lower() in GENERIC_NAMES]
    pat = re.compile(r"(?<![A-Za-z$\d.:!'])(\d+\.?\d*)(?![\d:!A-Za-z(])")
    ok = {"0", "1", "2", "12", "100", "1000", "365", "30", "0.5"}
    for ws in wb.worksheets:
        if ws.title == GUIDE_SHEET:
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    body = re.sub(r'"[^"]*"', "", v)          # drop text literals
                    body = re.sub(r"CHAR\(\d+\)", "", body)                 # line break in text
                    body = re.sub(r"EDATE\(([^,()]+),\d+\)", "", body)   # month offset in a date header
                    body = re.sub(r"MOD\(MONTH\([^()]*\),\d+\)", "", body)  # quarter-end test
                    body = re.sub(r"\$?[A-Z]{1,3}\$?\d+", "", body)  # drop cell refs
                    bad = [n for n in pat.findall(body) if n not in ok]
                    if bad:
                        hard.append({"cell": f"{ws.title}!{c.coordinate}", "formula": v, "numbers": bad})
    res["hardcoded_in_formulas"] = hard[:50]
    res["hardcoded_count"] = len(hard)
    res["verdict"] = "ok" if (res["guide_first"] and not undescribed and not hard and res["canvas_second"]
                              and not res["generic_tab_names"]) else "needs attention"
    return res


@mcp.tool()
def find_values(file: str, query: str, sheet: str = "") -> list[dict]:
    """Case-insensitive text search across cells. Returns up to 100 hits."""
    wb = openpyxl.load_workbook(_path(file), data_only=True)
    q, hits = query.lower(), []
    for ws in ([_sheet(wb, sheet)] if sheet else wb.worksheets):
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None and q in str(c.value).lower():
                    hits.append({"sheet": ws.title, "cell": c.coordinate, "value": _jsonable(c.value)})
                    if len(hits) >= 100:
                        return hits
    return hits


@mcp.tool()
def group_sum(file: str, sheet: str, group_col: str, value_col: str, header_row: int = 1) -> list[dict]:
    """Sum `value_col` grouped by `group_col` (column headers). Good for cost by project / category."""
    df = _df(file, sheet, header_row)
    for c in (group_col, value_col):
        if c not in df.columns:
            raise ValueError(f"Column '{c}' not found. Columns: {list(df.columns)}")
    out = df.groupby(group_col, dropna=False)[value_col].sum().sort_values(ascending=False).reset_index()
    return out.to_dict("records")


@mcp.tool()
def reconcile(file_a: str, sheet_a: str, key_a: str, amount_a: str,
              file_b: str, sheet_b: str, key_b: str, amount_b: str,
              tolerance: float = 0.5, header_row_a: int = 1, header_row_b: int = 1) -> dict:
    """Match two lists by key (invoice no., reference) and compare amounts.
    Returns matched count, amount differences, items only in A, items only in B. Read-only."""
    a = _df(file_a, sheet_a, header_row_a)[[key_a, amount_a]].copy()
    b = _df(file_b, sheet_b, header_row_b)[[key_b, amount_b]].copy()
    a.columns, b.columns = ["key", "amt_a"], ["key", "amt_b"]
    a["key"], b["key"] = a["key"].astype(str).str.strip(), b["key"].astype(str).str.strip()
    a, b = a.groupby("key", as_index=False).sum(), b.groupby("key", as_index=False).sum()
    m = a.merge(b, on="key", how="outer", indicator=True)
    both = m[m["_merge"] == "both"].copy()
    both["diff"] = both["amt_a"] - both["amt_b"]
    return {
        "matched_equal": int((both["diff"].abs() <= tolerance).sum()),
        "amount_differences": both[both["diff"].abs() > tolerance][["key", "amt_a", "amt_b", "diff"]].to_dict("records"),
        "only_in_a": m[m["_merge"] == "left_only"][["key", "amt_a"]].to_dict("records"),
        "only_in_b": m[m["_merge"] == "right_only"][["key", "amt_b"]].to_dict("records"),
        "total_a": float(a["amt_a"].sum()), "total_b": float(b["amt_b"].sum()),
    }


@mcp.tool()
def aging_report(file: str, sheet: str, date_col: str, amount_col: str,
                 as_of: str = "", party_col: str = "", header_row: int = 1) -> dict:
    """Receivable/payable aging buckets (0-30, 31-60, 61-90, 90+) from an invoice-date or due-date column.
    as_of = 'YYYY-MM-DD' (default today). Optional party_col gives per-customer/supplier breakdown."""
    df = _df(file, sheet, header_row)
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    ref = pd.Timestamp(as_of) if as_of else pd.Timestamp.today().normalize()
    df["days"] = (ref - df[date_col]).dt.days
    bins = [-10**6, 30, 60, 90, 10**6]
    labels = ["0-30", "31-60", "61-90", "90+"]
    df["bucket"] = pd.cut(df["days"], bins=bins, labels=labels)
    total = df.groupby("bucket", observed=False)[amount_col].sum().reindex(labels, fill_value=0)
    res = {"as_of": str(ref.date()), "buckets": {k: float(v) for k, v in total.items()},
           "total": float(df[amount_col].sum()), "rows_without_date": int(df[date_col].isna().sum())}
    if party_col:
        pv = df.pivot_table(index=party_col, columns="bucket", values=amount_col, aggfunc="sum",
                            fill_value=0, observed=False).reindex(columns=labels, fill_value=0)
        pv["total"] = pv.sum(axis=1)
        res["by_party"] = pv.sort_values("total", ascending=False).reset_index().to_dict("records")
    return res


@mcp.tool()
def check_workbook(file: str) -> dict:
    """Audit a workbook: formula errors (#REF!, #DIV/0!...), formulas with no cached value, hard-coded numbers in formula columns is NOT checked."""
    p = _path(file)
    wf, wv = openpyxl.load_workbook(p), openpyxl.load_workbook(p, data_only=True)
    errors, uncached, nformulas = [], [], 0
    for ws in wf.worksheets:
        wsv = wv[ws.title]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    nformulas += 1
                    v = wsv[c.coordinate].value
                    if v is None:
                        uncached.append(f"{ws.title}!{c.coordinate}")
                    elif isinstance(v, str) and v.startswith("#"):
                        errors.append({"cell": f"{ws.title}!{c.coordinate}", "error": v, "formula": c.value})
    return {"formulas": nformulas, "errors": errors[:100],
            "uncached_count": len(uncached), "uncached_sample": uncached[:20],
            "hint": "uncached = file never recalculated; run recalculate()" if uncached else "ok"}


@mcp.tool()
def open_workbook(file: str, app: str = "default") -> dict:
    """Open a workbook from the workspace in the user's own spreadsheet program (Excel, LibreOffice, or the system default).
    Works only when this server runs on the user's computer (Claude Desktop). `app`: "default", "excel" or "libreoffice".
    Never writes to the file. If Excel already has it open, save and close it first or use File > Reload."""
    p = _path(file)
    app = app.lower().strip()
    if app not in {"default", "excel", "libreoffice"}:
        raise ValueError('app must be "default", "excel" or "libreoffice"')
    plat = sys.platform
    try:
        if plat.startswith("win"):
            if app == "default":
                os.startfile(str(p))  # type: ignore[attr-defined]
            else:
                exe = {"excel": "excel", "libreoffice": "soffice"}[app]
                subprocess.Popen(["cmd", "/c", "start", "", exe, str(p)])
        elif plat == "darwin":
            cmd = ["open"] if app == "default" else ["open", "-a", {"excel": "Microsoft Excel", "libreoffice": "LibreOffice"}[app]]
            subprocess.Popen(cmd + [str(p)])
        else:
            if not (os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY")):
                return {"ok": False, "reason": "No screen on this machine (cloud or headless). Run this server on the user's own computer."}
            exe = shutil.which("soffice") or shutil.which("libreoffice") if app in {"libreoffice", "excel"} else shutil.which("xdg-open")
            if not exe:
                return {"ok": False, "reason": "No program found to open the file (install LibreOffice or xdg-utils)."}
            subprocess.Popen([exe, str(p)])
    except Exception as e:  # launcher missing, app not installed, etc.
        return {"ok": False, "reason": f"Could not start the program: {e}"}
    _audit("open_workbook", file, {"app": app})
    return {"ok": True, "file": str(p), "app": app}


@mcp.tool()
def recalculate(file: str) -> dict:
    """Recalculate all formulas using LibreOffice (if installed) so cached values are up to date. Backs up first."""
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return {"ok": False, "reason": "LibreOffice not installed on this machine; open the file in Excel and save."}
    p = _path(file)
    bk = _backup(p)
    out = p.parent / ".recalc"
    out.mkdir(exist_ok=True)
    subprocess.run([exe, "--headless", "--convert-to", "xlsx", "--outdir", str(out), str(p)],
                   check=True, capture_output=True, timeout=120)
    shutil.move(str(out / p.name), p)
    shutil.rmtree(out, ignore_errors=True)
    _audit("recalculate", file, None, bk)
    return {"ok": True, "backup": bk}


@mcp.tool()
def list_backups(file: str = "") -> list[str]:
    """List backups (newest first), optionally only for one file name."""
    if not BACKUP_DIR.exists():
        return []
    names = sorted((p.name for p in BACKUP_DIR.iterdir()), reverse=True)
    return [n for n in names if not file or n.endswith("__" + Path(file).name)]


@mcp.tool()
def restore_backup(backup_name: str, file: str) -> dict:
    """Restore `file` from a backup made by this server. Current version is backed up first."""
    src = (BACKUP_DIR / backup_name).resolve()
    if BACKUP_DIR.resolve() not in src.parents or not src.exists():
        raise ValueError("Unknown backup name")
    p = _path(file, must_exist=False)
    bk = _backup(p)
    shutil.copy2(src, p)
    _audit("restore_backup", file, {"from": backup_name}, bk)
    return {"restored": file, "from": backup_name, "previous_saved_as": bk}



# ------------------------------------------------- catalog & tax reference
DATA_DIR = Path(__file__).parent / "data"

_PIT, _INS, _MW = ["pit_deductions", "pit_brackets", "pit_unused_leave"], ["ins_rates", "ins_caps"], ["min_wage_2026", "min_wage_2027"]
_CIT = ["cit_standard", "cit_small", "cit_new_sme", "cit_st_fund"]
_AREA_TAX = {"A": ["vat_rates", "cit_standard", "cit_small"], "B": ["vat_rates", "einvoice_penalty"], "C": ["vat_rates"],
             "D": ["vat_rates", "vat_export_refund", "vat_on_the_spot", "us_ad_cvd_extrusions"], "E": ["cit_standard", "cit_small"],
             "F": ["vat_rates", "vat_reduction", "cit_standard", "cit_small", *_PIT, *_INS, "einvoice_penalty", "license_fee"],
             "G": [*_PIT, *_INS, *_MW], "H": ["cit_standard", "cit_small"], "I": []}
_MODULE_TAX = {"F01": ["vat_rates", "vat_reduction", "einvoice_penalty"], "F02": _CIT, "F04": _PIT, "F05": ["einvoice_penalty", "vat_rates"],
               "F07": _CIT, "F09": ["license_fee", "einvoice_penalty"], "D03": ["vat_rates", "vat_export_refund", "vat_on_the_spot"],
               "D05": ["us_ad_cvd_extrusions"], "G02": [*_INS, *_MW], "G07": _MW, "G01": [*_PIT, *_INS], "G03": [*_PIT, "ins_rates"],
               "G05": ["pit_unused_leave", "pit_deductions"], "E04": ["cit_standard"], "C07": ["vat_rates"], "I01": [], "I06": []}


def _module_tax_refs(module: dict) -> list[str]:
    return _MODULE_TAX.get(module["id"], _AREA_TAX.get(module["id"][0], []))


def _load(name: str) -> dict:
    text = (DATA_DIR / name).read_text(encoding="utf-8")
    return json.loads(P.sub(text))  # @TOKENS@ in shared data -> private real names


@mcp.tool()
def accounting_catalog(area: str = "", module_id: str = "", priority: int = 0, brand: str = "") -> dict:
    """The company's library of accounting workbooks (64+ modules in 9 areas: A Planning, B Projects, C Ops, D Export,
    E Group, F Tax, G Payroll, H Cash/funding, I Reporting). Use it BEFORE creating any workbook: pick the module,
    reuse its `sheets`, `numbers` (inputs), `watch` (tax/payout risks) and `scenarios` (seed what-if questions).
    Filters: area (letter A-I or name part), module_id (e.g. 'D03'), priority (1=build first), brand (SHW/EXP/ITA/OWNER)."""
    cat = _load("accounting_catalog.json")
    mods = cat["modules"]
    if module_id:
        mods = [m for m in mods if m["id"].lower() == module_id.lower()]
    if area:
        a = area.lower()
        mods = [m for m in mods if (m["id"].lower().startswith(a) if len(a) == 1 else a in m["area"].lower())]
    if priority:
        mods = [m for m in mods if m["priority"] == priority]
    if brand:
        mods = [m for m in mods if m["brands"] in ("ALL", brand.upper())]
    if not module_id:  # compact listing
        mods = [{k: m[k] for k in ("id", "name", "priority", "brands", "status")} for m in mods]
    return {"areas": cat["areas"], "count": len(mods), "modules": mods}


@mcp.tool()
def tax_reference(topic: str = "") -> dict:
    """Vietnam tax / payroll / trade reference numbers WITH source URL, effective date and confidence.
    Put these in the workbook's Assumptions sheet and copy source + checked date next to each rate.
    Anything marked VERIFY must be confirmed with the company's tax adviser before use. Warns if data is older than 120 days."""
    ref = _load("vn_tax_reference.json")
    items = ref["items"]
    if topic:
        t = topic.lower()
        items = [i for i in items if t in (i["id"] + " " + i["topic"] + " " + i["value"]).lower()]
    age = (datetime.now() - datetime.fromisoformat(ref["checked_on"])).days
    return {"checked_on": ref["checked_on"], "age_days": age,
            "stale_warning": "Reference older than 120 days: re-check sources before relying on it" if age > 120 else None,
            "count": len(items), "items": items}

# --------------------------------------------------------- add new tools above

if __name__ == "__main__":
    mcp.run()  # stdio
