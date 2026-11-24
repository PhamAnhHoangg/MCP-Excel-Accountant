# Excel Accounting MCP (v0.8.4)

Connects Claude to your company's Excel files for accounting work. Safe by design:
- Only touches files inside ONE workspace folder (`EXCEL_MCP_ROOT`, default `~/excel-workspace`).
- Every write makes a backup in `.backups/` and a line in `.audit.jsonl`.
- Never overwrites an existing file when creating; restore any version with `restore_backup`.

## Install (5 minutes)
1. Install Python 3.10+ and run: `pip install -r requirements.txt`
2. Put the Excel files you want Claude to use in the workspace folder (`~/excel-workspace`, or set `EXCEL_MCP_ROOT` to a shared company folder).
3. Add to Claude Desktop config (`claude_desktop_config.json`), then restart Claude Desktop:
```json
{
  "mcpServers": {
    "excel-accounting": {
      "command": "python3",
      "args": ["/FULL/PATH/excel-accounting-mcp/server.py"],
      "env": { "EXCEL_MCP_ROOT": "/FULL/PATH/excel-workspace" }
    }
  }
}
```
   On Windows use `python` and paths like `C:\\Users\\you\\excel-workspace`.
4. Add the skill: zip the `skills/accounting-excel` folder and upload it in Settings > Capabilities > Skills.
5. Test: ask "Run server_info and list my workbooks".

## Terminal shortcut `cl` (type, no window)
Add to your PowerShell profile (`notepad $PROFILE`):
```powershell
function cl { param([string]$n="Excel",[string]$p)
  $e = if ($p) { python C:\Users\<you>\MCP-Excel-Accountant\launcher.py --encoded $n $p } else { python C:\Users\<you>\MCP-Excel-Accountant\launcher.py --encoded $n }
  if ($e) { powershell -NoLogo -ExecutionPolicy Bypass -EncodedCommand $e } }
```
Then in any terminal: `cl Excel` starts that profile's Claude Code in the same window; `cl Excel "build the company model"` also types the first message. `/exit` returns to your normal prompt.

## Claude Launcher (one button per account / task)
`pythonw launcher.py` (or double-click `Claude-Launcher.bat`) opens a small window. Each profile gets a button that opens a PowerShell window already set to that profile's Claude account folder (`CLAUDE_CONFIG_DIR`), workspace folder and MCP, runs any start-up commands, then starts Claude Code, optionally with a first message typed for you. "Task" buttons start Claude with a ready-made prompt. Edit `%USERPROFILE%\claude-launcher\profiles.json` (button "Edit profiles", then "Reload") to add accounts such as CAD: give each its own `config_dir`, `workspace` and `mcp`. Sign in once per profile inside its terminal; the login stays in that profile's folder.

## No Claude Desktop? Use the command line or Windows setup
Inside the repo folder in PowerShell: `powershell -ExecutionPolicy Bypass -File .\setup_windows.ps1` installs the packages, creates `~/excel-workspace`, runs the tests and prints the exact connection commands for Claude Code and Claude Desktop. Without any Claude app: `python cli.py new-full SH-Window-model.xlsx --open` (also `info`, `list`, `modules`, `new-module <file> <ID>`, `check`, `recalc`, `open --app excel|libreoffice`). Always run commands from inside the repo folder (`cd C:\Users\<you>\MCP-Excel-Accountant`).

## Opening files on your own computer
This server runs LOCALLY (started by Claude Desktop), so it can reach the software on your machine. `open_workbook(file, app)` opens a workspace file in Excel, LibreOffice or your default program (Windows `start`, Mac `open`, Linux `xdg-open`). It never writes to the file. Typical flow: Claude builds or edits the file in your workspace folder, then calls `open_workbook`; you read and edit it in your own software. If Excel already has the file open, close it first or use File > Reload. In a cloud session there is no screen, so the tool returns a clear "not available" message.

## Privacy: company data stays out of the repository
Real company, brand and supplier names and business detail live ONLY in `company_profile.json` (git-ignored; see `company_profile.py`). The repository holds generic placeholders written as `@B1@`, `@B2@`, `@B3@`, `@SUP@`; the server replaces them with the private profile's names when it builds a workbook. Pictures, workbooks, zips and backups are git-ignored too. `python tests/test_privacy.py` fails if a real name from your profile appears in any file git would publish. Never share `company_profile.json`; back it up yourself.

### Automatic guard (install once after cloning)
`git config core.hooksPath .githooks` makes git run `tests/test_privacy.py` before every commit and refuse the commit if `company_profile.json` is staged or a name from your private profile appears in any file that would be published. (`git commit --no-verify` skips it; do not use that.)

## Changelog
- **0.8.4**: `.githooks/pre-commit` privacy guard.
- **0.8.3**: Privacy: real names and business detail moved to the private `company_profile.json`; generic placeholders in code, catalog, skill and docs; `tests/test_privacy.py`; `.gitignore` covers profile, workbooks, pictures, zips.
- **0.8.2**: "Customers & Strategy" tab rebuilt to the company layout: three columns (Item | Draft suggestion | Fill in after customer interviews, yellow), pink section bars BRAND / CUSTOMER PROFILE / STRATEGY / INTERNAL, row heights fitted to text.
- **0.8.1**: `launcher.py --encoded NAME [PROMPT]` and the `cl` PowerShell function: start a profile in the current terminal, no window.
- **0.8.0**: `launcher.py` + `Claude-Launcher.bat`: small Windows launcher, one button per profile and per task; builds the PowerShell start-up script (account folder, workspace, MCP registration, skill copy, start-up commands, first prompt). Test: `tests/test_launcher.py`.
- **0.7.1**: `cli.py` (use the tools from the command line, no Claude app needed) and `setup_windows.ps1` (one-step Windows setup that prints the Claude Code / Claude Desktop connection commands).
- **0.7.0**: New tool `open_workbook` (19 tools): opens a workspace file in Excel, LibreOffice or the system default on the user's own computer. Same sandbox as every tool; read-only; logged in the audit file.
- **0.6.8**: Skill rule 13a: send a picture of every tab with each workbook, explaining that the Claude preview clips cells.
- **0.6.7**: Back to ONE cell per Canvas box (all bullets in the cell, Wrap Text on, heights fitted), as the user prefers; the in-app preview ignores wrap, so deliver a picture of the Canvas alongside the file when the user wants to read it in chat. Replaces the one-line-per-row layout of 0.6.6.
- **0.6.6**: Canvas shows all text in every viewer, including the Claude app preview (which ignores Wrap Text and shows only the first line of a multi-line cell). One short line per row inside each box; long bullets continue on an indented next line; uniform row heights; no merged text cells. Key-numbers strip now starts below the grid wherever it ends.
- **0.6.5**: Canvas boxes are back to ONE cell per box (all bullets in one cell, Wrap Text on, row heights fitted to the text), as the user asked. Block 5 and 9 are a single formula that joins live numbers and text.
- **0.6.4**: Canvas line spacing tightened (wider columns, row heights fitted to text).
- **0.6.3**: Canvas text boxes no longer use merged cells: one bullet per row inside each box, wrapped, with row heights set from text length. Fixes boxes that showed as one long unwrapped line (or empty) in Google Sheets and previews. Block capacity: 8 bullets for blocks 8, 6, 2, 4, 1 (blocks 7 and 3: 6), 7 for blocks 9 and 5.
- **0.6.2**: Canvas redesign: calm palette (white boxes, charcoal bars, soft grey number panel) instead of yellow fills; every box filled with specific bulleted content for the three brands; revenue and cost boxes pull live numbers from the model; each box points to the tab holding its detail.
- **0.6.1**: The Canvas tab is now the company's standard 9-block business-model canvas (shared by the full model and every catalog module) with a linked key-numbers strip below.
- **0.6.0**: Layout redesign. The Guide is now a short index: numbered, clickable list of the workbook's REAL tabs (first "Canvas", then "Assumptions"...), with only the most important highlighted MUST-HAVE. All detail moved inside the tabs. `create_workbook(template="full")` builds a 13-tab linked company model (formulas, cross-checks, sourced VN rates); `create_workbook(module_id=...)` builds Canvas + Assumptions + module tabs + Agreements & Controls. New `must_have` argument; generic tab names such as "Summary" are flagged; `check_conventions` also checks Canvas is the first tab after the Guide. Code in `accounting_templates.py`.
- **0.5.0**: Guide "Must-have setup requirements" now has 30 standard rules in 5 groups (structure, numbers to set up first, taxes, payouts, control). `create_workbook(module_id=...)` adds the module's numbers-to-enter, tax/payout watch-list and the matching rates with source URL, date and confidence (yellow = verify). Priority-1 modules have 3+ specific watch items each.
- **0.4.0**: `data/accounting_catalog.json` (64 workbook modules in 9 areas: planning, projects, ops, export, group, tax, payroll, cash, reporting) and `data/vn_tax_reference.json` (18 sourced Vietnam tax/payroll/trade figures with confidence flags). New tools `accounting_catalog`, `tax_reference` (18 tools). Edit the JSON to add modules or refresh rates; keep `checked_on` current.
- **0.3.0**: Guide sheet (English) now has reading order, must-have setup rules, colour code, how-to-extend and scenario questions. New tool `check_conventions`. 16 tools.
- **0.2.0**: `create_workbook` now always adds a first sheet "Guide" (how to use + colour legend). Edit `_add_guide_sheet` in `server.py` to change the text for all future files.
- **0.1.0**: first release, 15 tools.

## Tools (19)
| Group | Tools |
|---|---|
| Explore | `server_info`, `list_workbooks`, `list_sheets`, `read_range`, `find_values` |
| Open in your software | `open_workbook` |
| Edit (auto-backup) | `write_cells`, `append_rows`, `create_workbook`, `recalculate` |
| Accounting analysis | `group_sum`, `reconcile`, `aging_report` |
| Safety | `check_workbook`, `check_conventions`, `accounting_catalog`, `tax_reference`, `list_backups`, `restore_backup` |

## Upgrade roadmap (add one at a time, test each)
1. **VAT summary**: input/output VAT by month from an invoice sheet.
2. **Cash-flow template**: monthly receipts, payments, closing balance with formulas.
3. **Project P&L**: revenue minus costs per project (per-project and per-order results).
4. **Payroll 3P/KPI check**: validate the payroll sheet against rules.
5. **Report pack**: one tool that builds the monthly director report workbook.
6. **Approvals**: require confirmation text before large writes.

## How to upgrade
1. Add a new `@mcp.tool()` function above the marker at the bottom of `server.py`.
2. Add a test in `tests/test_server.py`, run `python3 tests/test_server.py`.
3. Bump `VERSION`, add the workflow to `skills/accounting-excel/SKILL.md`, restart Claude Desktop.

## Limits (be aware)
- Charts, pivot tables and macros in existing files may be lost when openpyxl saves. Keep a copy of complex files; backups protect you, but test on a copy first.
- Formula results appear only after the file is recalculated (`recalculate` needs LibreOffice, or just open and save in Excel).
- Only .xlsx / .xlsm. Old .xls must be converted.
