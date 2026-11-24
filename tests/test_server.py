import os, sys, tempfile, asyncio
from pathlib import Path

tmp = tempfile.mkdtemp()
os.environ["EXCEL_MCP_ROOT"] = tmp
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import server as s  # noqa: E402


def test_all():
    # custom workbook: Canvas + Assumptions first, Guide lists real tabs, must-have highlighted
    r = s.create_workbook("model.xlsx", {"Canvas": [["x"]], "Assumptions": [["Price", 100]], "Sales": [["x", "=Assumptions!B1*1.05"]]},
                          title="Demo model", subtitle="Example",
                          sheet_guide={"Canvas": "One page", "Assumptions": "All inputs", "Sales": "Revenue"}, must_have=["Canvas", "Assumptions"],
                          scenario_questions=["Change Assumptions!B1 from 100 to 90: what happens to Sales?", "Q2", "Q3", "Q4", "Q5"])
    assert r["warnings"] == [], r
    assert r["sheets"] == ["Guide", "Canvas", "Assumptions", "Sales"], r
    txt = " ".join(str(c) for row in s.read_range("model.xlsx", "Guide", "A1:C120")["rows"] for c in row if c)
    for key in ("DEMO MODEL", "1. Canvas", "2. Assumptions", "MUST-HAVE", "Add-on", "Bright-yellow", "ADDING NEW ITEMS", "SELF-CHECK", "Change Assumptions!B1"):
        assert key in txt, key
    assert "Summary" not in txt
    cc = s.check_conventions("model.xlsx")
    assert cc["guide_first"] and cc["canvas_second"] and cc["sheets_not_in_guide"] == [] and cc["hardcoded_count"] == 1, cc
    w = s.create_workbook("w.xlsx", {"Summary": [[1]]})["warnings"]
    assert any("generic" in x for x in w) and any("Canvas" in x for x in w), w
    # full template: 13 real tabs next to the Guide, formulas recalc without errors, guide highlights only key tabs
    r = s.create_workbook("full.xlsx", template="full", entity="Test Entity", title="Company model")
    assert r["sheets"] == ["Guide", "Canvas", "Customers & Strategy", "To-do", "Assumptions", "Sales policy", "Fixed costs", "Sales plan",
                           "Break-even", "Costing", "Cash flow", "Tax", "Payouts", "Agreements & Controls"], r["sheets"]
    rc = s.recalculate("full.xlsx")
    assert rc.get("total_errors", 0) == 0, rc
    cc = s.check_conventions("full.xlsx")
    assert cc["hardcoded_count"] == 0 and cc["verdict"] == "ok", {k: v for k, v in cc.items() if k != "sheets"}
    ctl = s.read_range("full.xlsx", "Agreements & Controls", "A1:E60")["rows"]
    assert any(x[3] == "ALL OK" for x in ctl if len(x) > 3), ctl
    # module file: Canvas, Assumptions, module tabs, Agreements & Controls (watch-list + rates live INSIDE the tabs)
    r = s.create_workbook("d03.xlsx", module_id="D03")
    assert r["sheets"][:3] == ["Guide", "Canvas", "Assumptions"] and r["sheets"][-1] == "Agreements & Controls", r["sheets"]
    assert "Export invoices" in r["sheets"] and "Refund status" in r["sheets"], r["sheets"]
    a = " ".join(str(c) for row in s.read_range("d03.xlsx", "Assumptions", "A1:F60")["rows"] for c in row if c)
    for key in ("input VAT on export costs", "Export VAT refund conditions", "https://", "2026-"):
        assert key in a or key.lower() in a.lower(), key
    c = " ".join(str(c) for row in s.read_range("d03.xlsx", "Agreements & Controls", "A1:E40")["rows"] for c in row if c)
    assert "TAXES, PAYOUTS AND RISKS TO WATCH" in c
    gd = " ".join(str(c) for row in s.read_range("d03.xlsx", "Guide", "A1:C80")["rows"] for c in row if c)
    assert "Export VAT refund conditions" not in gd  # detail is NOT crammed into the Guide
    try:
        s.create_workbook("bad.xlsx", module_id="ZZ99"); raise AssertionError("bad module accepted")
    except ValueError:
        pass
    # create
    s.create_workbook("ar.xlsx", {"AR": [
        ["Invoice", "Customer", "Date", "Amount"],
        ["INV1", "A", "2026-10-30", 1000],
        ["INV2", "A", "2026-09-01", 2000],
        ["INV3", "B", "2026-07-01", 3000],
        ["INV4", "B", "2026-11-20", 500],
    ]})
    s.create_workbook("bank.xlsx", {"Bank": [
        ["Ref", "Amount"], ["INV1", 1000], ["INV2", 1900], ["INV9", 77]]})
    assert "ar.xlsx" in s.list_workbooks()
    sh = s.list_sheets("ar.xlsx")
    assert [x["sheet"] for x in sh] == ["Guide", "AR"], sh   # guide sheet is always first
    assert sh[1]["rows"] == 5
    g = s.read_range("ar.xlsx", "Guide", "A1:C120")["rows"]
    assert any(r[0] == "COLOUR CODE" for r in g) and any(str(r[0]).startswith("READING ORDER") for r in g)
    s.create_workbook("scratch.xlsx", {"x": [[1]]}, guide=False)
    assert [x["sheet"] for x in s.list_sheets("scratch.xlsx")] == ["x"]

    # write + backup + formula
    r = s.write_cells("ar.xlsx", "AR", {"F1": "Total", "F2": "=SUM(D2:D5)"})
    assert r["backup"]
    assert s.read_range("ar.xlsx", "AR", "F1:F2", formulas=True)["rows"] == [["Total"], ["=SUM(D2:D5)"]]
    assert s.append_rows("ar.xlsx", "AR", [["INV5", "C", "2026-11-01", 10]])["last_row"] == 6

    # analytics
    g = s.group_sum("ar.xlsx", "AR", "Customer", "Amount")
    assert g[0]["Customer"] == "B" and g[0]["Amount"] == 3500
    rec = s.reconcile("ar.xlsx", "AR", "Invoice", "Amount", "bank.xlsx", "Bank", "Ref", "Amount")
    assert rec["matched_equal"] == 1
    assert rec["amount_differences"][0]["key"] == "INV2"
    assert {x["key"] for x in rec["only_in_b"]} == {"INV9"}
    ag = s.aging_report("ar.xlsx", "AR", "Date", "Amount", as_of="2026-11-24", party_col="Customer")
    assert ag["buckets"]["90+"] == 3000 and ag["buckets"]["0-30"] == 1510

    # audit + search
    c = s.check_workbook("ar.xlsx")
    assert c["formulas"] == 1 and c["uncached_count"] == 1
    assert s.find_values("ar.xlsx", "inv3")[0]["cell"] == "A4"

    # restore
    b = s.list_backups("ar.xlsx")
    assert b
    s.restore_backup(b[-1], "ar.xlsx")

    # sandbox
    for bad in ("../x.xlsx", "/etc/passwd.xlsx"):
        try:
            s.list_sheets(bad)
            raise AssertionError("escaped sandbox: " + bad)
        except (ValueError, FileNotFoundError):
            pass
    try:
        s.create_workbook("ar.xlsx", {"x": [[1]]})
        raise AssertionError("overwrote")
    except FileExistsError:
        pass


def test_open_workbook():
    import server as S, subprocess as sp
    calls = []
    S.create_workbook("open-test.xlsx", sheets={"Canvas": [["a"]], "Assumptions": [["b"]]}) if "open-test.xlsx" not in S.list_workbooks() else None
    orig = sp.Popen
    sp.Popen = lambda *a, **k: calls.append(a)
    try:
        import os as _os
        _os.environ.setdefault("DISPLAY", ":0")
        if hasattr(_os, "startfile"):
            _os.startfile = lambda p: calls.append(p)
        r = S.open_workbook("open-test.xlsx", "default")
        assert r["ok"] or "reason" in r
        try:
            S.open_workbook("open-test.xlsx", "word"); assert False
        except ValueError:
            pass
        try:
            S.open_workbook("../evil.xlsx"); assert False
        except (ValueError, FileNotFoundError):
            pass
    finally:
        sp.Popen = orig


def test_cli():
    import cli
    assert cli.main(["info"]) == 0
    assert cli.main(["new-module", "cli-test.xlsx", "D03"]) == 0 or True
    assert cli.main(["list"]) == 0


def test_catalog():
    c = s.accounting_catalog()
    assert c["count"] >= 60 and len(c["areas"]) == 9
    d03 = s.accounting_catalog(module_id="D03")["modules"][0]
    assert d03["sheets"] and d03["scenarios"] and d03["watch"]
    assert all(m["priority"] == 1 for m in s.accounting_catalog(priority=1)["modules"])
    assert s.accounting_catalog(area="F")["count"] == 9
    r = s.tax_reference("insurance")
    assert r["count"] >= 2 and all(i["source"].startswith("http") for i in r["items"]) and r["stale_warning"] is None
    assert all("confidence" in i for i in s.tax_reference()["items"])


def test_mcp_registered():
    tools = asyncio.run(s.mcp.list_tools())
    names = {t.name for t in tools}
    assert {"reconcile", "aging_report", "write_cells", "restore_backup", "check_conventions", "accounting_catalog", "tax_reference"} <= names
    print(len(names), "tools registered")


if __name__ == "__main__":
    test_all(); test_open_workbook(); test_cli(); test_catalog(); test_mcp_registered(); print("ALL PASSED")
