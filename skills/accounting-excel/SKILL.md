---
name: accounting-excel
description: Accounting work in Excel via the excel-accounting MCP — reconcile bank/invoice lists, AR/AP aging, cost by project, cash-flow reports, workbook audits. Use whenever the user asks to check, reconcile, summarize, build or fix accounting spreadsheets for the company and its related entities.
---

# Accounting Excel skill

## Ground rules (never skip)
1. Start with `server_info`, then `list_workbooks` and `list_sheets`. Read headers with `read_range` before using any tool that needs column names.
2. Never guess columns, units or currency. If unclear (VND vs USD, VAT included or not), ask once.
3. Writes always back up automatically. After a write, say which cells changed and the backup name.
4. Do arithmetic with tools (`group_sum`, `reconcile`, `aging_report`) or Excel formulas, never in your head.
5. Prefer live formulas (`=SUM(...)`) over pasted numbers so the user can audit. Run `recalculate` then `check_workbook` after adding formulas.
6. Report numbers with thousand separators and the currency. Flag anything uncertain as uncertain.
7. Data is confidential: do not copy figures outside the workspace.

7. WORKBOOK LAYOUT. Tab 1 = "Guide"; the next tabs are the real working tabs in reading order, always "Canvas" first then "Assumptions". The Guide only LISTS those tabs (clickable, numbered) and highlights the 3-6 MUST-HAVE ones; all detail (numbers to enter, taxes, payouts, controls, risk watch-list, sourced rates) lives INSIDE the tabs, never crammed into the Guide. Never create a generic tab such as "Summary". English only. Use `create_workbook(template="full")` for the 13-tab company model (Canvas, Customers & Strategy, To-do, Assumptions, Sales policy, Fixed costs, Sales plan, Break-even, Costing, Cash flow, Tax, Payouts, Agreements & Controls) or `create_workbook(module_id=...)` for a catalog module; for a custom file pass `sheets`, `sheet_guide` for every tab and `must_have`. The "Canvas" tab is ALWAYS the company's 9-block business-model canvas (blocks 8 Partners, 6 Activities, 7 Resources, 2 Value, 4 Relationships, 3 Channels, 1 Segments, 9 Costs, 5 Revenue, laid out as in the company sample), followed by a strip of key numbers linked from the other tabs. Do not invent another canvas layout; the user will extend this one later. Canvas colours are deliberately calm (white boxes, charcoal title bars, one soft grey panel for numbers); never fill canvas boxes with bright or yellow colours. Fill each box with 5-8 specific, bulleted points (names, products, markets, terms), not generic words; boxes 5 and 9 pull live numbers from the model. Each canvas box is ONE cell holding all its bullets (newline separated) with Wrap Text on and row heights fitted to the text; do not split a box into one cell per bullet or one line per row. The Claude app preview ignores Wrap Text, so after building a canvas also render it (LibreOffice to PDF, crop the Canvas page to PNG) and send the picture so the user can read it in chat. ALWAYS pass `title`, `subtitle` and 5-8 `scenario_questions`. Yellow = not confirmed: tell the user.
8. Scenario questions must be specific to the file: quote real cell addresses ("Change Assumptions!B25 from 2 to 1: does break-even change? Why does profit change?"), ask for cause-and-effect, and cover accounting basics where relevant: profit vs cash, break-even, cost per unit, working capital, tax threshold, scenario best/base/worst. Offer to give the answer key if the user asks.
9. Cell colours in every new workbook: inputs = blue text on light-yellow fill, formulas = black, links to other sheets = green, key results = bright-yellow fill. Put an "Assumptions" sheet first after the Guide; other sheets reference it, no hard-coded numbers in formulas. Mark sample data "EXAMPLE".
10. Before delivering a new workbook run `check_workbook` (errors) and `check_conventions` (guide + hard-coded numbers); fix what they report.

11. NEW WORKBOOK = CATALOG MODULE. Call `accounting_catalog` first (area / priority / brand, then `module_id`). `create_workbook(module_id=...)` builds Canvas, Assumptions (module numbers + sourced rates), the module's own tabs and Agreements & Controls (tax / payout / risk watch-list with owner and status). Then add formulas to the module tabs, and write file-specific `scenario_questions` from the module seeds. If no module fits, build a custom file and tell the user which module to add to the catalog.
12. RATES COME FROM `tax_reference`, never from memory. Put each rate in the Assumptions sheet as a blue input with the source URL and "checked on" date beside it. Anything marked VERIFY or "single source": show a visible "VERIFY WITH TAX ADVISER" note next to the number. If `stale_warning` is set, say so before using it. If a rate is not in the reference, search the web, cite it, and tell the user it is not yet in the reference.
13a. PICTURES OF EVERY TAB. The Claude app preview clips cells (shows only the first line and cuts long text), so after building or changing a workbook, render EVERY tab to a picture (LibreOffice headless to PDF, then `pdftoppm -png`, crop each tab's page) and send the pictures with the file. Tell the user plainly: "The Claude preview clips the text in boxes, so I am sending a picture of each tab to read here; the Excel file itself shows everything when you open it in Excel or Google Sheets."
13b. OPEN IT FOR THE USER. When this server runs on the user's own computer (Claude Desktop), after saving a workbook call `open_workbook(file)` so it opens in their Excel or LibreOffice. If it returns ok=false (cloud or no screen), send the file and the tab pictures instead. Never call it on a file the user has not asked to see.
13. Never claim a tax conclusion as final. Give the calculation and the assumption; the accountant or tax adviser confirms the treatment.

## Workflows

**Build a module from the catalog**: `accounting_catalog(module_id=...)` -> `tax_reference(topic)` for every rate involved -> `create_workbook` (Guide, Canvas, Assumptions first) -> add formulas -> `recalculate` -> `check_workbook` + `check_conventions` -> deliver with the list of numbers still marked VERIFY.

**Bank / invoice reconciliation**: read both headers → `reconcile` (tolerance 0.5 for VND rounding, adjust if needed) → report matched count, amount differences, items only in A, only in B → suggest likely cause for each group (timing, partial payment, missing entry).

**Receivable / payable aging**: confirm the date column is invoice date or due date → `aging_report` with `party_col` → list top 5 overdue parties and the 90+ total → draft a payment-reminder message only if asked.

**Cost by project / category**: `group_sum` → compare with the budget column if present → highlight over-budget items.

**Cash-flow report**: read the sheets with receipts and payments → build a monthly table with `create_workbook` (opening balance, inflows, outflows, closing balance as formulas) → `check_workbook`.

**Workbook audit before sending to a director**: `check_workbook` → list errors with cell addresses → fix only after the user approves.

## Company conventions (edit these as the company's rules become clear)
- Default currency: VND. Export business: USD, state the exchange rate used.
- Accounting regime: confirm with the user (TT200 or TT133) before classifying accounts.
- Sheet naming: `YYYY-MM_<topic>`; one workbook per legal entity.

## Output style
Short summary first (3–5 lines), then a table, then recommended next actions. Do not pad.
