"""Generates data/accounting_catalog.json and data/vn_tax_reference.json for the plugin."""
import json, sys
from pathlib import Path

OUT = Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)

AREAS = {
    "A": "Planning & business model", "B": "Projects & contracts (@B1@)", "C": "Procurement, inventory & factory",
    "D": "Export (@B2@)", "E": "Group: 3 entities & owner", "F": "Tax & compliance (Vietnam)",
    "G": "Payroll & payouts", "H": "Cash, funding & owner payouts", "I": "Reporting & control",
}
mods = []
def M(id, name, why, numbers, watch, sheets, scen, prio, brands="ALL"):
    mods.append(dict(id=id, area=AREAS[id[0]], name=name, why=why, numbers=numbers, watch=watch, sheets=sheets,
                     scenarios=scen, priority=prio, brands=brands, status="planned"))

# ---------------- A Planning
M("A01","Business model canvas & unit economics","One page: who pays, for what, how much margin per unit.",
  ["price per m2/set","direct cost per m2/set","gross margin %"],["margin by brand differs: contractor vs system licensor vs exporter"],
  ["Canvas","Assumptions","Unit economics"],["Raise material cost 8% and keep price: what is the new margin per m2?"],2)
M("A02","Assumptions master (prices, costs, FX, tax rates)","Single source of truth; every other file reads from it.",
  ["price list","cost list","FX rates (USD, EUR, CAD, AUD)","tax rates with source+date"],["tax rates change yearly: store source and date next to each rate"],
  ["Assumptions","Rates & sources","Change log"],["Change the VAT rate cell: which sheets move?"],1)
M("A03","Sales policy: channels, commission, discount","Commission and discount are real costs that shrink margin.",
  ["commission %","discount limits","channel split of revenue"],["commission paid to individuals is taxable income (check withholding)"],
  ["Policy","Commission calc","Margin after commission"],["Discount 5% more to win a bid: how many extra projects needed to recover margin?"],2)
M("A04","Fixed-cost budget & monthly overhead","Fixed cost decides break-even and the cash you need each month.",
  ["rent","salaries","insurance","depreciation","interest","marketing"],["fixed cost must be covered even in months with no sales"],
  ["Fixed costs","Monthly budget","Funding need"],["No sales for 2 months: how much cash is burned?"],1)
M("A05","Sales forecast & scenarios (best / base / worst)","Plans must survive a bad quarter.",
  ["new customers/month","win rate","average ticket","cycle time"],["revenue timing vs cash timing"],
  ["Forecast","Scenarios","Summary"],["Win rate falls from 30% to 20%: which month does cash go negative?"],2)
M("A06","Break-even & pricing per m2 / per set","Know the minimum price that still pays all costs.",
  ["variable cost/unit","fixed cost/month","target margin"],["price quoted below full cost looks like a sale but loses money"],
  ["Inputs","Break-even","Price ladder"],["Fixed cost doubles: how many m2 per month to break even?"],1)

# ---------------- B Projects
M("B01","Quotation / BOQ costing per m2","Build the price from real quantities: profile kg, glass m2, hardware, labour, install, overhead, margin.",
  ["profile kg/m2","glass m2","hardware sets","labour hours","install cost","overhead %","margin %"],["waste/cutting loss must be in the cost","VAT shown separately on quotes"],
  ["Inputs","BOQ","Cost build-up","Quote summary"],["Aluminium price +10%: what is the quote increase to hold margin?"],1,"SHW")
M("B02","Project job costing & P&L per project","Shows which projects make money after ALL costs (not just materials).",
  ["contract value","material","labour","subcontract","install","site cost","warranty reserve"],["cost overruns found late; allocate shared factory cost fairly"],
  ["Project list","Cost ledger","Project P&L"],["Project delayed 2 months: how much extra site cost and finance cost?"],1,"SHW")
M("B03","Billing schedule & milestone payments","Advance, progress and acceptance payments decide your cash, not the contract total.",
  ["advance %","progress %","acceptance %","payment terms days"],["issue the invoice at the right time (invoice timing penalties exist from Jan 2026)"],
  ["Contract terms","Billing schedule","Cash timing"],["Client pays 30 days late: what happens to cash in the month and to profit?"],1,"SHW")
M("B04","Retention / warranty holdback & release tracker","Part of the contract is held until warranty ends; it is profit you have not collected yet.",
  ["retention %","release date","guarantee instead of cash?"],["(verify) terms are contract-specific; record per contract"],
  ["Retention list","Release calendar"],["Retention is 5% on 20 projects: how much cash is locked for 12 months?"],2,"SHW")
M("B05","Receivable aging & collection","Late customers are the #1 reason profitable contractors run out of cash.",
  ["invoice date","due date","amount","customer"],["bad-debt provisions: tax deductibility needs documents (verify rules)"],
  ["Invoices","Aging","Top overdue","Collection log"],["A 90+ day invoice of 500m: what is the cash and profit impact if written off?"],1,"SHW")
M("B06","Subcontractor & installation crew payments","Crew payments are large, often cash, and need documents.",
  ["rate per m2/day","advance","completion %","acceptance"],["(verify) withholding when paying individuals; non-cash payment rules for large invoices"],
  ["Crew list","Work done","Payment due","Documents checklist"],["Crew rate +10%: effect on project margin?"],2,"SHW")
M("B07","Bid bond & performance guarantee tracker","Guarantees tie up cash or credit lines and cost fees.",
  ["guarantee amount","fee %","expiry","cash margin deposit"],["expired guarantee not returned = money stuck"],
  ["Guarantees","Fees","Expiry calendar"],["5 bids at once need bonds: how much credit line is used?"],2,"SHW")
M("B08","WIP & revenue recognition","Decides when profit is booked: by % complete or at handover. Accounting vs tax may differ.",
  ["cost to date","estimated total cost","% complete","billed to date"],["(verify) tax timing of revenue vs accounting timing"],
  ["Projects","% complete","Over/under billing"],["Estimated cost rises 10% mid-project: how does booked profit change?"],2,"SHW")
M("B09","Warranty provision & after-sales cost","Warranty claims are a future cost caused by today's sales.",
  ["claim rate %","avg claim cost","warranty years"],["(verify) provision deductibility conditions"],
  ["Claims history","Provision calc"],["Claim rate doubles: what reserve do you need?"],3,"SHW")
M("B10","Contract risk register","Penalty clauses, delay damages, price escalation, scope change.",
  ["penalty %/day","price escalation clause?","payment security"],["unrecoverable aluminium price rises if contract has no escalation clause"],
  ["Contracts","Risks","Mitigation"],["Aluminium +15% after signing and no escalation clause: loss per project?"],3,"SHW")

# ---------------- C Ops
M("C01","Material price tracker (aluminium, glass, hardware)","Aluminium and EUR prices move; quotes go stale.",
  ["supplier price","date","LME/market reference","FX"],["price validity of quotes; FX on EUR hardware"],
  ["Prices","Trend","Alerts"],["Price rises 6% next month: which open quotes lose margin?"],2)
M("C02","Purchase orders & supplier payables","Know what you owe and when; protect early-payment discounts and supplier relations.",
  ["PO","delivery","invoice","terms","paid"],["VAT input deduction needs valid invoice and payment evidence (verify thresholds)"],
  ["POs","Payables aging","Payment plan"],["Suppliers shorten terms from 45 to 30 days: cash effect?"],1)
M("C03","Inventory & stock count","Stock is cash sitting on shelves; count differences are losses.",
  ["opening","receipts","issues","closing","costing method"],["stock loss deductibility needs documents (verify)"],
  ["Items","Movements","Count vs book"],["Slow stock of 300m for 6 months: carrying cost?"],2)
M("C04","Cutting yield, scrap & recovery","Profile cutting loss is a real cost; scrap sale is income with VAT rules.",
  ["standard yield %","actual yield %","scrap kg","scrap price"],["VAT treatment of scrap/by-products was clarified from 1 Jan 2026"],
  ["Cutting plan","Yield","Scrap sales"],["Yield drops 3 points: cost per m2?"],2)
M("C05","Factory overhead & product costing","Electricity, rent, supervisors, machine time must be allocated to products.",
  ["overhead pool","allocation base","machine hours"],["under-allocated overhead makes quotes too cheap"],
  ["Overhead","Allocation","Product cost"],["Factory utilisation falls 20%: cost per unit?"],2)
M("C06","Fixed assets & depreciation register","Machines and vehicles: depreciation is deductible only within rules.",
  ["cost","useful life","method","in-service date"],["(verify) depreciation limits for vehicles; minimum asset value; useful-life ranges"],
  ["Assets","Depreciation","Disposals"],["Buy a new CNC machine of 1.5bn: monthly depreciation and cash impact?"],2)
M("C07","Imported hardware landed cost (@SUP@, EUR)","True cost = price + freight + insurance + import duty + import VAT + customs fees + FX.",
  ["EUR price","EUR/VND","freight","import duty %","import VAT %","clearance fees"],["import VAT is cash out before sales; duty classification (HS code); FX loss"],
  ["Shipment inputs","Landed cost","Cost per unit"],["EUR rises 4% against VND: effect on unit cost and price list?"],1,"ITA")
M("C08","Capacity vs fixed-cost absorption","Idle capacity is cost without revenue.",
  ["capacity","utilisation","fixed cost"],["do not accept low-margin jobs just to fill capacity without checking contribution"],
  ["Capacity","Utilisation","Absorption"],["Utilisation 60% vs 85%: unit cost difference?"],3)

# ---------------- D Export
M("D01","Export order costing & margin per container","Margin in USD after freight, insurance, packing, duties and FX.",
  ["FOB/CIF/DDP price","freight","packing","insurance","FX"],["Incoterms decide who pays freight/duty; margin must be per container"],
  ["Order","Cost to serve","Margin per container"],["Freight rises 25%: is a CIF order still profitable?"],1,"EXP")
M("D02","FX exposure & gains/losses","You sell in USD/CAD/AUD, cost in VND/EUR: FX moves profit.",
  ["open receivables by currency","open payables by currency","rates"],["realised vs unrealised FX; year-end revaluation (verify tax treatment)"],
  ["Exposure","Gain/loss","Hedge ideas"],["VND strengthens 3%: effect on USD receivables?"],1,"EXP")
M("D03","Export VAT 0% & VAT refund tracker","Exports are 0% VAT; input VAT can be refunded but needs complete documents and time.",
  ["input VAT on export costs","export value","documents status","refund claimed/received"],["refund conditions and documents changed in 2025-2026; cash is stuck until refund"],
  ["Export invoices","Input VAT","Documents","Refund status"],["Refund takes 6 months instead of 2: how much cash is locked?"],1,"EXP")
M("D04","Customs & export document checklist","Missing or inconsistent documents block shipment, refund and payment.",
  ["contract","invoice","packing list","customs declaration","bank payment proof","C/O"],["document names/numbers must match across set"],
  ["Shipments","Checklist","Exceptions"],["Packing list weight differs from invoice: what documents break?"],2,"EXP")
M("D05","Destination tariffs & trade remedies tracker","Buyer-country duties can wipe out your price advantage.",
  ["HS code","duty rate","anti-dumping/countervailing status","origin rules"],["(verify) US additional tariffs; AD/CVD case outcomes; Canada/Australia/Cambodia rules"],
  ["Markets","Duty table","Alerts"],["US adds 10% duty: new price needed to keep margin?"],2,"EXP")
M("D06","Foreign buyer credit risk, deposits & L/C","Foreign customers can pay late or not at all, and recovery is hard.",
  ["deposit %","payment terms","L/C fees","credit limit"],["use deposits, L/C or credit insurance for new buyers"],
  ["Buyers","Exposure","Overdue"],["A buyer with 200k USD overdue 90 days: provision and cash?"],2,"EXP")
M("D07","Distributor / agent commission & rebate","Retail channel (e.g. home-improvement chains) demands rebates, returns allowances.",
  ["rebate %","returns %","marketing fund","commission"],["rebates reduce net revenue; record them accurately"],
  ["Channel terms","Rebates","Net margin"],["Chain asks 3% rebate: net margin?"],3,"EXP")

# ---------------- E Group
M("E01","Intercompany transactions & reconciliation","@B1@ buys from @B3@; @B3@ may sell to @B2@. Both sides must match.",
  ["sales between entities","payables/receivables","price basis"],["unmatched intercompany balances = errors and audit questions"],
  ["Transactions","Matching","Differences"],["@B3@ invoices 1bn but @B1@ recorded 900m: where is the gap?"],1)
M("E02","Transfer pricing & related-party disclosure","Related companies must trade at arm's-length prices and disclose.",
  ["related-party sales/purchases","margin benchmark","documents"],["(verify) disclosure forms and thresholds; owner-controlled entities are related parties"],
  ["Related parties","Pricing test","Disclosure checklist"],["@B3@ sells to @B1@ at a 5% discount versus others: any tax risk?"],2)
M("E03","Consolidated group view & owner dashboard","Owner needs one view across 3 entities with intercompany removed.",
  ["P&L per entity","cash per entity","intercompany elimination"],["do not double count intercompany revenue"],
  ["Entity P&L","Eliminations","Group summary"],["Remove intercompany: what is real group revenue?"],2)
M("E04","Licence/royalty & foreign-payment tax","Payments to foreign parties (royalty, licence, services) may trigger foreign contractor tax and treaty questions.",
  ["payment type","amount","withholding rate","treaty basis"],["(verify) foreign contractor tax on payments to the Italian group and DTA relief documents"],
  ["Foreign payments","Tax calc","Documents"],["Pay 20k EUR licence fee: what tax is due on top?"],2,"ITA")

# ---------------- F Tax
M("F01","VAT: input / output & invoice matching","Monthly or quarterly VAT: output minus deductible input; every input invoice must be valid.",
  ["output VAT","input VAT","carry-forward","reduced-rate items"],["valid invoice + payment evidence; temporary rate reduction through end of 2026 for eligible items (verify eligibility)"],
  ["Sales invoices","Purchase invoices","VAT calc","Mismatches"],["Supplier invoice issued but not declared by supplier: effect on your deduction?"],1)
M("F02","CIT: quarterly provisional & annual finalisation","Corporate income tax from 2025: tiered rates by revenue (conditions apply).",
  ["accounting profit","non-deductible items","taxable income","rate","prepayments"],["(verify) which rate your entity qualifies for; provisional payment accuracy; loss carry-forward"],
  ["P&L","Adjustments","CIT calc","Payments"],["Revenue crosses 50bn: how much extra CIT next year?"],1)
M("F03","Non-deductible expense & document audit","Expenses without valid documents are added back to taxable income.",
  ["expense list","document status","payment method"],["(verify) non-cash payment threshold; invoices in company name; personal expenses"],
  ["Expenses","Document check","Add-back"],["200m of expenses without invoices: extra CIT?"],1)
M("F04","PIT withholding & annual finalisation","Employee salary tax with new 2026 deductions and brackets; also commissions and crews.",
  ["gross pay","insurance","personal deduction","dependents","taxable income"],["new brackets and deductions apply to the 2026 tax year"],
  ["Employees","PIT calc","Year-end"],["Employee with 2 dependents earns 30m/month: PIT before and after the new law?"],1)
M("F05","E-invoice compliance register","Wrong timing, cancelled and adjusted invoices cause penalties.",
  ["invoice no.","issue date","delivery/acceptance date","status"],["invoice timing penalties updated from 16 Jan 2026"],
  ["Invoices","Timing check","Cancelled/adjusted"],["Invoice issued after delivery month: what is the risk?"],2)
M("F06","Tax calendar & deadline tracker","Missing a deadline costs penalties and interest.",
  ["filing dates","payment dates","responsible person"],["monthly/quarterly VAT, quarterly CIT, PIT, insurance, annual finalisations"],
  ["Calendar","Responsibilities","Reminders"],["Which filings fall in the same week next month?"],1)
M("F07","Tax losses, incentives & funds","SME rates, new-company exemptions, science & technology fund, loss carry-forward.",
  ["loss by year","eligible incentives","fund appropriation"],["(verify) eligibility conditions; fund can be set aside up to 20% of taxable income under the new CIT law"],
  ["Losses","Incentives","Fund"],["A loss year followed by profit: tax saved from carry-forward?"],3)
M("F08","Tax inspection readiness file","A prepared folder cuts inspection stress and penalty risk.",
  ["documents by tax type","contracts","bank statements"],["keep evidence for VAT, CIT, PIT, transfer pricing"],
  ["Checklist","Evidence index","Gaps"],["Which 10 documents would you need first in an inspection?"],3)
M("F09","Other taxes & fees","Licence fee, import duty, foreign contractor tax, land/property, environment fees.",
  ["fee type","base","rate","due date"],["business licence fee exemption from 2026 (verify scope); others to verify"],
  ["Taxes & fees","Calendar"],["Which fees apply to the factory site?"],3)

# ---------------- G Payroll
M("G01","Payroll with 3P / KPI","Pay for Position, Person, Performance; ties salary to results.",
  ["base pay","KPI weight","KPI result","allowances"],["variable pay still counts for tax and insurance depending on how it is structured (verify base for insurance)"],
  ["Roster","3P calc","KPI","Payslip"],["Sales KPI 80% vs 100%: pay and company cost difference?"],1)
M("G02","Insurance & union fee calculator with caps","Employer pays about 23.5% on top of salary in 2026 (within caps); employee pays 10.5%.",
  ["insured salary","caps","rates","regional minimum wage"],["caps changed on 1 Jul 2026; minimum wage changes 1 Jan 2027"],
  ["Rates & caps","Per employee","Company total"],["Hire 5 workers at 12m/month: total monthly employer cost?"],1)
M("G03","Bonus, commission & bid-win payout","Technical bid-win bonuses and sales commissions are taxable pay and real cost.",
  ["bonus pool","rate","eligibility","tax treatment"],["include in PIT; decide whether part counts for insurance (verify)"],
  ["Rules","Payout calc","Tax effect"],["Win a 5bn bid: bonus cost and net to the employee?"],1)
M("G04","Overtime, night shift & factory pay","Overtime premiums have special tax treatment; hours must be recorded.",
  ["hours","premium %","night shift","piece rate"],["(verify) exempt portion of overtime premium for PIT"],
  ["Timesheet","OT calc","Cost"],["Peak season +30 OT hours per worker: cost?"],2)
M("G05","13th-month / Tet bonus, severance & unused leave","Year-end payouts need cash planning; unused-leave pay is now exempt from PIT.",
  ["bonus rule","severance rule","leave balance"],["unused annual-leave payment exempt from PIT under the new law"],
  ["Accruals","Payout plan","Cash needed"],["Tet bonus 1 month: how much cash must be saved by December?"],2)
M("G06","Headcount plan & cost per head","Total cost per employee, not just salary.",
  ["salary","insurance","union fee","equipment","training"],["each new hire adds 23.5% employer cost on insured pay"],
  ["Headcount","Cost per head","Plan"],["Add 3 sales staff: break-even revenue needed?"],3)
M("G07","Labour contract & regional minimum wage check","Salaries must not fall below regional minimums; they rose in Jan 2026 and again in Jan 2027.",
  ["region","minimum wage","pay by employee"],["check every salary against the minimum each year"],
  ["Employees","Minimum check","Exceptions"],["Minimum wage +7% in 2027: which employees fall below?"],2)

# ---------------- H Cash & funding
M("H01","Cash-flow: 13-week and 12-month","Profit is not cash: timing of collection and payment decides survival.",
  ["receipts by week","payments by week","opening balance","minimum cash"],["tax payments, insurance, bonuses, import VAT are lumpy cash outflows"],
  ["Inputs","13-week","12-month","Alerts"],["Which week is the lowest cash? What funding fills it?"],1)
M("H02","Working capital & credit line utilisation","Receivables + inventory - payables = cash locked in operations.",
  ["receivable days","inventory days","payable days","credit limit"],["bank covenants and guarantee limits"],
  ["Working capital","Credit line"],["Customers pay 20 days later: how much extra credit is used?"],1)
M("H03","Loan & interest schedule","Loan interest is deductible only within limits.",
  ["principal","rate","term","repayment profile"],["(verify) interest deductibility cap under current decrees"],
  ["Loans","Schedule","Interest tax effect"],["Rate +2%: monthly interest and profit?"],2)
M("H04","Owner payouts: salary, dividends, shareholder loans","How the owner takes money out changes tax and cash.",
  ["salary","dividend","loan to/from company"],["(verify) PIT on dividends; shareholder loan interest and documentation"],
  ["Payout options","Tax comparison","Cash impact"],["Take 1bn as salary vs dividend vs loan: tax and cash difference?"],1,"OWNER")
M("H05","Profit-sharing / partnership agreement calculator","Capital, effort, profit split between partners written down before starting.",
  ["capital shares","effort shares","profit rule","exit rule"],["write the exit and loss rules before the first sale"],
  ["Agreement terms","Profit split","Scenarios"],["Project loses 200m: who bears how much?"],2,"OWNER")
M("H06","Equity & capital contribution register","Who contributed what and when; needed for legal and tax records.",
  ["capital contributions","dates","ownership %"],["legal registration of changes (verify)"],
  ["Contributions","Ownership"],["Add a new partner at 20%: dilution?"],3,"OWNER")
M("H07","Personal finance & investment tracker (owner)","Keep personal and company money separate; track net worth and investments.",
  ["assets","liabilities","investments","income/expenses"],["do not mix personal and company payments"],
  ["Net worth","Investments","Budget"],["Market falls 20%: effect on net worth?"],3,"OWNER")

# ---------------- I Reporting & control
M("I01","Monthly management report pack","One workbook the director reads each month: P&L, cash, receivables, backlog by brand.",
  ["revenue","gross margin","cash","AR/AP","backlog"],["compare to budget and last year"],
  ["P&L by brand","Cash","AR/AP","Backlog"],["What changed most versus last month and why?"],1)
M("I02","Budget vs actual variance","Explains differences, not just shows them.",
  ["budget","actual","variance","reason"],["flag variances over a set threshold"],
  ["Budget","Actual","Variance"],["Materials 12% over budget: price or quantity?"],2)
M("I03","KPI dashboard by department","Sales, technical, site, factory, procurement, HR, accounting KPIs.",
  ["KPI list","targets","actuals"],["few KPIs per department"],
  ["KPI list","Dashboard"],["Which department missed its KPI and what did it cost?"],2)
M("I04","Approval limits & internal control checklist","Who may approve what amount; separation between ordering, receiving and paying.",
  ["approval limits","petty cash limit","signers"],["fraud and error risk is highest in cash payments and returns"],
  ["Limits","Checklist","Exceptions"],["One person orders, receives and pays: what can go wrong?"],2)
M("I05","Month-end & year-end closing checklist","A fixed routine prevents missing accruals and cut-off errors.",
  ["accruals","prepayments","stock count","reconciliations"],["cut-off of revenue and costs"],
  ["Checklist","Reconciliations","Sign-off"],["A December invoice booked in January: what changes?"],2)
M("I06","Document retention & file naming","Retention rules and consistent names make audits and searches fast.",
  ["document types","retention years","naming rule"],["(verify) legal retention periods"],
  ["Retention table","Naming rule"],["Can you find the 2024 contract with client X in 1 minute?"],3)


EXTRA = {
 "A02": ["Label every price and cost as VAT-inclusive or VAT-exclusive", "Keep a change log: who changed which rate, when, why"],
 "A04": ["Employer insurance and union fee (about 23.5% of insured salary) belong in fixed cost", "Spread yearly costs (audit, licences, insurance premiums) monthly so no month surprises you"],
 "A06": ["A price below full cost looks like a sale but is paid for later by a cash shortage", "Include warranty reserve and finance cost in the minimum price"],
 "B01": ["Set a quote validity period: aluminium and euro prices move", "Show VAT separately (10%, or the reduced rate only if the item qualifies - verify)", "Add contingency for installation risk and site delays"],
 "B02": ["Allocate shared factory and site costs by a written rule", "Book cost to the project when incurred, not when paid", "Hold a warranty reserve per project"],
 "B03": ["Issue the invoice at the milestone or acceptance date: wrong timing is penalised from 16 Jan 2026", "(verify) VAT timing on advance payments", "Cash arrives by milestone, so forecast cash from this schedule, not from contract value"],
 "B05": ["Chase before 30 days overdue, escalate at 60, stop work at 90 (set your own rule)", "(verify) bad-debt provision is deductible only with required legal documents", "Retention is not overdue until its release date"],
 "C02": ["Match purchase order, delivery note and invoice before paying", "Supplier invoice must be a valid e-invoice in the company name for VAT deduction", "Compare early-payment discount with the cost of the cash"],
 "C07": ["(verify) import duty depends on HS code and origin: confirm the rate for aluminium-system hardware from Italy", "Import VAT paid at customs is deductible only with complete documents", "Customs exchange rate vs bank rate: the difference is an FX cost", "Freight, insurance and broker fees belong in landed cost"],
 "D01": ["Incoterms (FOB / CIF / DDP) decide who pays freight, insurance and duty", "Margin must include packing, container loading, bank charges and rebates", "Costs are in VND/EUR but revenue is in USD/CAD/AUD: show the FX effect"],
 "D02": ["Record the rate at invoice date, payment date and year-end (revaluation)", "Show FX gain/loss separately from operating profit", "Match USD costs with USD revenue where possible (natural hedge)"],
 "D03": ["(verify) 0% VAT needs: sales contract, commercial invoice, customs declaration and bank payment proof", "Late or missing documents mean the input VAT stays as a cost", "The supplier of your materials must have declared VAT on its invoices for you to be refunded", "Refund can take months: show locked cash in the cash-flow file"],
 "E01": ["Same price, date and amount in both companies' books", "Month-end cut-off: delivered vs invoiced", "(verify) VAT invoice is required even between related companies"],
 "F01": ["Reconcile the VAT return to the ledger every period", "(verify) invoices from risky suppliers can lose the deduction", "Input VAT on non-business use is not deductible"],
 "F02": ["Check which revenue tier applies (based on prior-year revenue)", "Add back fines, personal expenses, costs without valid invoices; (verify) limits on interest and depreciation", "(verify) provisional payments must be close to the final tax to avoid late interest"],
 "F03": ["Invoices must be in the company name and match the payment", "Review recurring risky items: entertainment, gifts, personal fuel", "(verify) non-cash payment threshold for large invoices"],
 "F04": ["Withhold on every pay run and finalise yearly", "(verify) withholding on commissions and crew paid as individuals depends on contract type", "Dependents must be registered before the deduction applies"],
 "F06": ["Set each date from the tax authority calendar (verify current dates)", "A missed deadline means penalty and late-payment interest", "Name an owner and a backup for every filing"],
 "G01": ["Structure matters: variable pay may or may not count for insurance (verify)", "Cap KPI payouts to protect cash", "Write the KPI formula into the policy or contract"],
 "G02": ["Employer pays about 23.5% and employee 10.5%, within caps", "Caps changed on 1 Jul 2026 (base salary 2.53m): update every employee", "Unemployment cap uses the regional minimum wage, not the base salary"],
 "G03": ["Bonus is taxable income for the employee", "(verify) whether the bonus counts for insurance depends on how it is structured", "Cap the pool as a % of project gross profit and pay after cash is collected"],
 "H01": ["Include tax, insurance and bonus payment dates", "Include import VAT outflows and VAT refund inflows", "Show a minimum cash buffer in every week"],
 "H02": ["Credit limit vs guarantees already issued", "Check bank covenants, if any", "Watch the gap between receivable days and payable days"],
 "H04": ["(verify) salary, dividend and shareholder loan are taxed differently", "A shareholder loan needs a written agreement", "Do not pay the owner's personal expenses from the company"],
 "I01": ["Use the same definitions every month", "Reconcile to bank and ledger before reporting", "Show every large variance with its reason"],
}
for m in mods:
    m["watch"] = m["watch"] + EXTRA.get(m["id"], [])

(OUT/"accounting_catalog.json").write_text(json.dumps({"version":"1.0","areas":AREAS,"modules":mods}, ensure_ascii=False, indent=1), encoding="utf-8")

# ---------------- Tax reference (verified against sources on 2026-11-24)
CH = "2026-11-24"
R = []
def T(id, topic, value, effective, source, confidence, note=""):
    R.append(dict(id=id, topic=topic, value=value, effective=effective, source=source, confidence=confidence, note=note, checked_on=CH))
T("cit_standard","CIT standard rate","20%","Law 67/2025/QH15, from 1 Oct 2025 (tax year 2025 on)","https://taxsummaries.pwc.com/Vietnam/Corporate/Taxes-on-corporate-income","verified (3+ sources)")
T("cit_small","CIT rate for small revenue","15% if prior-year revenue <= VND 3bn; 17% if VND 3-50bn","Law 67/2025; Decree 320/2025 (from 15 Dec 2025); Circular 20/2026","https://www.vietnam-briefing.com/doing-business-guide/vietnam/taxation-and-accounting/country-wise-tax-structure/corporate-income-tax","verified (3 sources) but conditions apply","Not every enterprise qualifies; confirm with your tax adviser.")
T("cit_new_sme","CIT exemption for newly established SMEs","100% exemption for first 3 years (Decree 20/2026; Resolution 198/2025)","2026","https://www.gvw.com/en/news/blog/detail/corporate-income-tax-cit-incentives-for-small-and-medium-sized-enterprises-smes-from-2026","single source: VERIFY","Eligibility conditions not confirmed.")
T("cit_st_fund","Science & technology fund appropriation","Up to 20% of annual taxable income (was 10%)","Law 67/2025","https://www.roedl.com/en/insights/vietnam-new-corporate-income-tax-law/","single source: VERIFY")
T("vat_rates","VAT rates","0% export; 5% listed items; 10% standard","Law 48/2024 from 1 Jul 2025","https://www.vietnam-briefing.com/news/vietnam-new-vat-law-key-compliance-guidance.html/","verified (3+ sources)")
T("vat_reduction","Temporary VAT reduction","2 percentage points off the 10% rate for eligible goods/services until end of 2026","until 31 Dec 2026","https://www.vietnam-briefing.com/news/vietnam-vat-rates-and-applicability-in-2025-a-brief-guide.html/","single source: VERIFY eligibility","Check which of your products qualify before using 8%.")
T("vat_export_refund","Export VAT refund conditions","Seller must have declared and paid VAT on input invoices the buyer claims; Law 149/2025 changed refund rules from 1 Jan 2026","from 1 Jul 2025 / 1 Jan 2026","https://www.ey.com/en_vn/technical/tax/tax-and-law-updates/amendments-to-the-law-on-value-added-tax","verified (2 sources)","Document requirements: confirm current list with customs/tax adviser.")
T("vat_on_the_spot","0% VAT for on-the-spot import-export","Allowed from 1 Jul 2025","1 Jul 2025","https://www.vietnam-briefing.com/news/vietnam-new-vat-law-key-compliance-guidance.html/","single source: VERIFY")
T("pit_deductions","PIT personal and dependent deduction","VND 15.5m/month personal; VND 6.2m/month per dependent","2026 tax year (Res. 110/2025)","https://luatvietnam.vn/legal-news/vietnams-new-personal-income-tax-regime-what-foreign-workers-and-investors-need-to-know-4729-110522-article.html","verified (4+ sources)")
T("pit_brackets","PIT brackets (salary, residents)","5 progressive brackets; 5% up to VND 10m/month taxable; top rate 35% per draft","Law 109/2025 (1 Jul 2026), salary rules from 2026 tax year","https://vietnamnet.vn/en/vietnam-revises-personal-income-tax-law-easing-burden-for-millions-2471255.html","bracket count & 5% first band verified; full table VERIFY","Get the full bracket table from Decree 253/2026 or Circular 87/2026 before building PIT formulas.")
T("pit_unused_leave","PIT on unused annual-leave payout","Exempt","Law 109/2025","https://www.pwc.com/vn/en/publications/2025/251212-new-personal-income-tax-law.pdf","single source: VERIFY")
T("ins_rates","Mandatory insurance + union fee","Employer: SI 17.5%, HI 3%, UI 1%, union fee 2% = 23.5%. Employee: SI 8%, HI 1.5%, UI 1% = 10.5%","2026","https://www.secondtalent.com/resources/vietnam-payroll-benefits-tax/","verified (3 sources)","Foreign employees have different rules (e.g. UI not applicable).")
T("ins_caps","Insurance caps","SI/HI/union cap = 20 x base salary: VND 50.6m/month from 1 Jul 2026 (was 46.8m). UI cap = 20 x regional minimum: VND 106.2m (Region I)","1 Jul 2026 (Decree 161/2026)","https://support.remote.com/hc/en-us/articles/29318874191245-What-are-the-income-tax-and-salary-deduction-rates-in-Vietnam","verified (3+ sources)")
T("min_wage_2026","Regional minimum wage 2026","Region I VND 5.31m, Region III 4.14m, Region IV 3.70m per month (+7.2%)","1 Jan 2026 (Decree 293/2025)","https://www.bakermckenzie.com/en/insight/publications/2026/03/vietnam-change-to-the-regional-minimum-wage-in-2026","verified (3 sources)","Region II (4.73m) from a government news report; confirm your factory's region.")
T("min_wage_2027","Regional minimum wage 2027","Region I VND 5.7m per month (+7.3%)","1 Jan 2027 (Decree 399/2026)","https://www.crowe.com/vn/news/regional-minimum-wages-applicable-from-1-january-2027","verified (2 sources); other regions VERIFY")
T("license_fee","Business licence fee","Exempt from 1 Jan 2026","1 Jan 2026","https://baonghean.vn/en/hang-loat-chinh-sach-moi-quan-trong-co-hieu-luc-tu-thang-1-2026-10317457.html","single source (text truncated): VERIFY scope")
T("einvoice_penalty","E-invoice timing penalties","Issuing invoices at the wrong time is penalised under Decree 310/2025","16 Jan 2026","https://baonghean.vn/en/hang-loat-chinh-sach-moi-quan-trong-co-hieu-luc-tu-thang-1-2026-10317457.html","single source: VERIFY amounts")
T("us_ad_cvd_extrusions","US AD/CVD case on aluminium extrusions (incl. Vietnam)","US ITC final vote: industry not materially injured, so no AD duties ordered on extrusions from the 14 countries","Oct 2025 (verify year)","https://vietnamnet.vn/en/us-not-impose-anti-dumping-duties-on-aluminum-extrusions-from-vietnam-2337653.html","single source: VERIFY","Other US tariffs (Section 232, reciprocal tariffs, fence/door HS codes) were NOT checked. Verify before pricing US orders.")
(OUT/"vn_tax_reference.json").write_text(json.dumps({"checked_on":CH,"items":R}, ensure_ascii=False, indent=1), encoding="utf-8")
print(len(mods), "modules;", len(R), "reference items")
