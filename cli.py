"""Command-line launcher: use the Excel Accounting tools without Claude Desktop.

Examples (PowerShell, inside the repo folder):
  python cli.py info
  python cli.py new-full SH-Window-model.xlsx --open
  python cli.py modules
  python cli.py new-module D03-vat-refund.xlsx D03 --open
  python cli.py check SH-Window-model.xlsx
  python cli.py recalc SH-Window-model.xlsx
  python cli.py open SH-Window-model.xlsx --app excel
Files live in the workspace folder (EXCEL_MCP_ROOT, default ~/excel-workspace).
"""
from __future__ import annotations

import argparse
import json
import sys

import server as S


def show(x) -> None:
    print(json.dumps(x, indent=2, ensure_ascii=False, default=str))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Excel Accounting tools without Claude Desktop")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("info", help="show version and workspace folder")
    sub.add_parser("list", help="list workbooks in the workspace")
    sub.add_parser("modules", help="list catalog modules (id, name, priority)")
    for name, helptext in [("check", "check errors + company conventions"), ("recalc", "recalculate formulas (needs LibreOffice)")]:
        p = sub.add_parser(name, help=helptext); p.add_argument("file")
    p = sub.add_parser("open", help="open a workbook in Excel / LibreOffice / default")
    p.add_argument("file"); p.add_argument("--app", default="default", choices=["default", "excel", "libreoffice"])
    p = sub.add_parser("new-full", help="build the 13-tab company model")
    p.add_argument("file"); p.add_argument("--entity", default=""); p.add_argument("--open", action="store_true")
    p = sub.add_parser("new-module", help="build a catalog module workbook")
    p.add_argument("file"); p.add_argument("module_id"); p.add_argument("--entity", default=""); p.add_argument("--open", action="store_true")
    a = ap.parse_args(argv)

    if a.cmd == "info":
        show(S.server_info())
    elif a.cmd == "list":
        show(S.list_workbooks())
    elif a.cmd == "modules":
        cat = S.accounting_catalog()
        mods = cat.get("modules", cat)
        show(mods)
    elif a.cmd == "check":
        show({"errors": S.check_workbook(a.file), "conventions": S.check_conventions(a.file)})
    elif a.cmd == "recalc":
        show(S.recalculate(a.file))
    elif a.cmd == "open":
        show(S.open_workbook(a.file, a.app))
    elif a.cmd in {"new-full", "new-module"}:
        kw = {"template": "full"} if a.cmd == "new-full" else {"module_id": a.module_id}
        res = S.create_workbook(a.file, entity=a.entity, **kw)
        show(res)
        if a.open:
            show(S.open_workbook(a.file))
    return 0


if __name__ == "__main__":
    sys.exit(main())
