/**
 * Fully synthetic demo dataset. Every ledger name, code and amount here is invented.
 * It builds sheet grids in the same shape as the two real workbooks, so the demo runs through the
 * very same parsers as an uploaded file. It is also the main fixture for the unit tests.
 *
 * Built-in scenarios: a regroup across two provisions (the "gratuity" pattern), a pair of payables
 * with offsetting variances, a ledger renamed between the two files (one xlsb-only, one Copy-only),
 * local GAAP captions that Ind AS reclassifies (fixed deposits, deposits, capital creditors, bank
 * charges), and Ind AS 116 / 109 / 19 / 12 adjustments that net to nil on the balance sheet.
 */
import type { Grid } from '../engine/types'

type Amt = [number, number, number] // rupees for 31-03-2026, 31-03-2025, 01-04-2024 (FY 2023-24 close)

interface DemoLedger {
  code: string
  name: string
  primary: string
  parent: string
  face: string
  note: string
  category: string
  sub: string
  line: string
  amt: Amt
  /** Copy's local GAAP when it differs from the signed books. */
  copy?: Amt
}

const L = (
  code: string, name: string, primary: string, face: string, note: string,
  category: string, line: string, amt: Amt, copy?: Amt,
): DemoLedger => ({ code, name, primary, parent: primary, face, note, category, sub: category, line, amt, copy })

const LEDGERS: DemoLedger[] = [
  // ---------------------------------------------------------------- balance sheet
  L('D1001', 'DEMO LEDGER A - PLANT AND MACHINERY', 'Fixed Assets', 'Property Plant and Equipment', 'Plant and machinery', 'Property, Plant and Equipment', 'Gross carrying amount', [5200000, 4800000, 4500000]),
  L('D1002', 'DEMO LEDGER B - ACC DEP PLANT', 'Fixed Assets', 'Property Plant and Equipment', 'Plant and machinery', 'Property, Plant and Equipment', 'Accumulated depreciation', [-1800000, -1300000, -900000]),
  L('D1003', 'DEMO LEDGER C - CAPITAL WIP', 'Fixed Assets', 'Capital work in progress', 'Capital work in progress', 'Capital work in progress', 'Capital work-in-progress', [300000, 600000, 0]),
  L('D1101', 'DEMO LEDGER D - RAW MATERIAL STOCK', 'Current Assets', 'Inventories', 'Raw materials', 'Inventories', 'Raw materials', [4000000, 3500000, 3000000]),
  L('D1102', 'DEMO LEDGER E - FINISHED GOODS', 'Current Assets', 'Inventories', 'Finished goods', 'Inventories', 'Finished goods', [1500000, 1200000, 1000000]),
  L('D1201', 'DEMO CUSTOMER 1', 'Sundry Debtors', 'Trade receivables', 'Unsecured, considered good', 'Trade receivables - current', 'Unsecured, considered good', [6000000, 5000000, 4000000]),
  L('D1202', 'DEMO CUSTOMER 2', 'Sundry Debtors', 'Trade receivables', 'Unsecured, considered good', 'Trade receivables - current', 'Unsecured, considered good', [3500000, 3000000, 2500000]),
  L('D1203', 'DEMO CUSTOMER 3', 'Sundry Debtors', 'Trade receivables', 'Unsecured, considered good', 'Trade receivables - current', 'Unsecured, considered good', [1200000, 900000, 600000]),
  L('D1301', 'DEMO LEDGER F - BANK CURRENT ACCOUNT', 'Bank Accounts', 'Cash and cash equivalents', 'Balances with banks', 'Cash and cash equivalents', 'Balances with banks', [0, 0, 0]), // plug, see below
  L('D1302', 'DEMO LEDGER G - FIXED DEPOSIT', 'Bank Accounts', 'Cash and cash equivalents', 'Other bank balances', 'Other bank balances', 'Fixed deposits (3-12 months)', [2000000, 1800000, 1500000]),
  L('D1401', 'DEMO LEDGER H - RENT SECURITY DEPOSIT', 'Loans & Advances (Asset)', 'Long term loans and advances', 'Security deposits', 'Other financial assets - non current', 'Security and other deposits', [600000, 600000, 600000]),
  L('D1402', 'DEMO LEDGER I - GST INPUT CREDIT', 'Duties & Taxes', 'Short term loans and advances', 'Balances with government authorities', 'Other current assets', 'Balances with government authorities', [800000, 700000, 600000]),
  L('D1403', 'DEMO LEDGER J - ADVANCE TAX NET', 'Current Assets', 'Short term loans and advances', 'Advance tax', 'Income Tax Assets', 'Advance tax (net of provision)', [200000, 150000, 100000]),
  L('D1404', 'DEMO LEDGER K - PREPAID EXPENSES', 'Current Assets', 'Other current assets', 'Prepaid expenses', 'Other current assets', 'Prepaid expenses', [100000, 80000, 60000]),
  L('D2001', 'DEMO LEDGER L - SHARE CAPITAL', 'Capital Account', 'Share capital', 'Equity shares', 'Equity Share Capital', 'Equity share capital', [-5000000, -5000000, -5000000]),
  L('D2002', 'DEMO LEDGER M - RETAINED EARNINGS', 'Capital Account', 'Reserves and surplus', 'Surplus in P&L', 'Other Equity', 'Retained earnings', [0, 0, 0]), // computed
  L('D2101', 'DEMO VENDOR 1', 'Sundry Creditors', 'Trade payables due to others', 'Due to others', 'Trade Payables - current', 'Dues of creditors other than MSME', [-4500000, -4000000, -3500000]),
  L('D2102', 'DEMO VENDOR 2', 'Sundry Creditors', 'Trade payables due to others', 'Due to others', 'Trade Payables - current', 'Dues of creditors other than MSME', [-3000000, -2600000, -2200000]),
  L('D2103', 'DEMO VENDOR 3', 'Sundry Creditors', 'Trade payables due to others', 'Due to others', 'Trade Payables - current', 'Dues of creditors other than MSME', [-400000, -350000, -300000], [-410000, -350000, -300000]),
  L('D2104', 'DEMO VENDOR 4', 'Sundry Creditors', 'Trade payables due to others', 'Due to others', 'Trade Payables - current', 'Dues of creditors other than MSME', [-200000, -150000, -100000], [-190000, -150000, -100000]),
  L('D2105', 'DEMO VENDOR 5', 'Sundry Creditors', 'Trade payables due to others', 'Due to MSME', 'Trade Payables - current', 'Dues of micro and small enterprises', [-800000, -700000, -600000]),
  L('D2201', 'DEMO LEDGER N - SALARY PAYABLE', 'Current Liabilities', 'Other current liabilities', 'Employee dues', 'Other current liabilities', 'Employee dues', [-400000, -350000, -300000]),
  L('D2202', 'DEMO LEDGER O - TDS PAYABLE', 'Duties & Taxes', 'Other current liabilities', 'Statutory dues', 'Other current liabilities', 'Statutory dues', [-120000, -100000, -80000]),
  L('D2203', 'DEMO LEDGER P - CAPITAL CREDITORS', 'Current Liabilities', 'Other current liabilities', 'Creditors for capital goods', 'Other financial liabilities - current', 'Creditors for capital goods', [-250000, -150000, 0]),
  L('D2301', 'DEMO LEDGER Q - GRATUITY PAYABLE', 'Provisions', 'Long term provisions', 'Provision for employee benefits', 'Provisions - non current', 'Provision for gratuity', [-300000, -240000, -180000], [0, -240000, -180000]),
  L('D2303', 'DEMO LEDGER S - PROVISION FOR EXPENSES', 'Current Liabilities', 'Short term provisions', 'Provision for expenses', 'Provisions - current', 'Provision for expenses', [-150000, -120000, -100000]),
  L('D2304', 'DEMO LEDGER T - DEFERRED TAX LIABILITY', 'Current Liabilities', 'Deferred tax liabilities Net', 'Deferred tax liabilities', 'Deferred tax liabilities, net', 'Deferred tax liabilities (net)', [-320000, -270000, -230000]),
  // ---------------------------------------------------------------- profit and loss
  L('D3001', 'DEMO LEDGER U - SALES DOMESTIC', 'Sales Accounts', 'Revenue from operations', 'Sale of products', 'Revenue From Operations', 'Sale of products', [-32000000, -28000000, -24000000]),
  L('D3002', 'DEMO LEDGER V - SALES EXPORT', 'Sales Accounts', 'Revenue from operations', 'Sale of products', 'Revenue From Operations', 'Sale of products', [-8000000, -6000000, -4000000]),
  L('D3101', 'DEMO LEDGER W - INTEREST ON DEPOSITS', 'Indirect Incomes', 'Other Income', 'Interest income', 'Other Income', 'Interest income on bank deposits', [-240000, -210000, -180000]),
  L('D4001', 'DEMO LEDGER X - PURCHASE OF RAW MATERIAL', 'Purchase Accounts', 'Cost of material consumed', 'Purchases', 'Cost of materials consumed', 'Cost of materials consumed', [23000000, 20000000, 17000000]),
  L('D4002', 'DEMO LEDGER Y - CHANGE IN STOCK', 'Purchase Accounts', 'Change in Inventories of work in progress and finished goods', 'Change in inventories', 'Changes in inventories of finished goods, Stock in Trade and work in progress', 'Changes in inventories', [-300000, -200000, 100000]),
  L('D4101', 'DEMO LEDGER Z - SALARIES AND WAGES', 'Indirect Expenses', 'Employee benefit expenses', 'Salaries and wages', 'Employee benefits expense', 'Salaries and wages', [4800000, 4200000, 3600000]),
  L('D4102', 'DEMO LEDGER AA - GRATUITY EXPENSE', 'Indirect Expenses', 'Employee benefit expenses', 'Contribution to funds', 'Employee benefits expense', 'Gratuity', [260000, 220000, 180000]),
  L('D4201', 'DEMO LEDGER AB - BANK CHARGES', 'Indirect Expenses', 'Finance costs', 'Bank charges', 'Other expenses', 'Bank charges', [120000, 110000, 90000]),
  L('D4301', 'DEMO LEDGER AC - DEPRECIATION', 'Indirect Expenses', 'Depreciation and amortization expenses', 'Depreciation', 'Depreciation and amortization expense', 'Depreciation on property, plant and equipment', [500000, 400000, 350000]),
  L('D4401', 'DEMO LEDGER AD - RENT FACTORY', 'Indirect Expenses', 'Other expenses', 'Rent', 'Other expenses', 'Rent', [600000, 600000, 600000]),
  L('D4402', 'DEMO LEDGER AE - POWER AND FUEL', 'Indirect Expenses', 'Other expenses', 'Power and fuel', 'Other expenses', 'Power and fuel', [1800000, 1600000, 1400000]),
  L('D4403', 'DEMO LEDGER AF - REPAIRS', 'Indirect Expenses', 'Other expenses', 'Repairs and maintenance', 'Other expenses', 'Repairs and maintenance', [450000, 380000, 300000]),
  L('D4404', 'DEMO LEDGER AG - FOREIGN EXCHANGE', 'Indirect Expenses', 'Other expenses', 'Exchange difference', 'Other expenses', 'Net loss on foreign exchange', [120000, -80000, 50000]),
  L('D4501', 'DEMO LEDGER AH - CURRENT TAX', 'Indirect Expenses', 'Tax Expenses', 'Current tax', 'Tax expenses', 'Current tax', [1400000, 1200000, 1000000]),
  L('D4502', 'DEMO LEDGER AI - DEFERRED TAX', 'Indirect Expenses', 'Tax Expenses', 'Deferred tax', 'Tax expenses', 'Deferred tax', [50000, 40000, 30000]),
]

/** The "gratuity" pattern: one ledger that the Ind AS workbook splits over two rows. */
const SPLIT = {
  code: 'D2302',
  name: 'DEMO LEDGER R - PROVISION FOR GRATUITY',
  primary: 'Current Liabilities',
  face: 'Short term provisions',
  note: 'Provision for employee benefits',
  amt: [-50000, -40000, -30000] as Amt,
  // Copy rows: non-current part and current part
  copyNC: [-300000, 0, 0] as Amt,
  copyC: [-50000, -40000, -30000] as Amt,
}

/** A ledger renamed between the two files. */
const RENAMED = {
  signed: { code: 'D2106', name: 'DEMO VENDOR 6 (NEW CODE)', amt: [-75000, 0, 0] as Amt },
  copy: { code: 'D2199', name: 'DEMO VENDOR 6 (OLD CODE)', amt: [-75000, 0, 0] as Amt },
}

/** Ind AS adjustment rows: [ref, layer, narration, category, line item, jv per period]. */
type AdjRow = [string, 'A' | 'P', string, string, string, Amt]
const ADJUSTMENTS: AdjRow[] = [
  ['1', 'A', 'Ind AS 116 Right-of-use asset - demo factory lease', 'Right of Use Assets', 'Right-of-use assets', [1000000, 1500000, 2000000]],
  ['2', 'A', 'Ind AS 116 Lease liability (non-current)', 'Lease liabilities non-current', 'Lease liabilities', [-550000, -1000000, -1400000]],
  ['3', 'A', 'Ind AS 116 Lease liability (current)', 'Lease liabilities current', 'Lease liabilities', [-535000, -560000, -600000]],
  ['4', 'A', 'Ind AS 109 Security deposit - write-down to amortised cost', 'Other financial assets - non current', 'Security and other deposits', [-123500, -136000, -150000]],
  ['5', 'A', 'Ind AS 109 Prepaid rent - day-one difference on the deposit', 'Other current assets', 'Prepaid rent (Ind AS 109)', [72000, 96000, 120000]],
  ['6', 'A', 'Ind AS 12 Deferred tax on the Ind AS adjustments', 'Deferred tax liabilities, net', 'Deferred tax on Ind AS adjustments', [34354, 25168, 7550]],
  ['7', 'A', 'Ind AS 101 Retained earnings - transition adjustment', 'Other Equity', 'Ind AS transition adjustment', [22450, 22450, 22450]],
  ['8', 'A', 'Retained earnings - Ind AS P&L adjustments FY 2024-25 (contra of layer P)', 'Other Equity', 'Ind AS retained earnings movement', [22449, 22449, 0]],
  ['9', 'A', 'Retained earnings - Ind AS P&L adjustments FY 2025-26 (contra of layer P)', 'Other Equity', 'Ind AS retained earnings movement', [-17585, 0, 0]],
  ['10', 'A', 'Ind AS 19 Remeasurement reserve (OCI) - gross', 'Other Equity', 'Remeasurement of defined benefit plans (OCI)', [100000, 40000, 0]],
  ['11', 'A', 'Ind AS 12 Tax on OCI remeasurement', 'Other Equity', 'Remeasurement of defined benefit plans (OCI)', [-25168, -10067, 0]],
  ['12', 'P', 'Ind AS 116 Rent on the demo factory lease reversed', 'Other expenses', 'Rent', [-600000, -600000, 0]],
  ['13', 'P', 'Ind AS 116 Depreciation on right-of-use asset', 'Depreciation and amortization expense', 'Depreciation on right-of-use assets', [500000, 500000, 0]],
  ['14', 'P', 'Ind AS 116 Interest on lease liability', 'Finance costs', 'Interest on lease liabilities', [125000, 160000, 0]],
  ['15', 'P', 'Ind AS 109 Unwinding of discount on security deposit', 'Other Income', 'Interest income - unwinding of deposit', [-12500, -14000, 0]],
  ['16', 'P', 'Ind AS 109 Amortisation of prepaid rent', 'Other expenses', 'Rent', [24000, 24000, 0]],
  ['17', 'P', 'Ind AS 19 Remeasurement reclassified to OCI', 'Employee benefits expense', 'Less: remeasurement reclassified to OCI', [-60000, -40000, 0]],
  ['18', 'P', 'Ind AS 12 Deferred tax on Ind AS adjustments (P&L)', 'Tax expenses', 'Deferred tax', [-9186, -17618, 0]],
  ['19', 'P', 'Ind AS 12 Tax on remeasurement moved to OCI', 'Tax expenses', 'Deferred tax', [15101, 10067, 0]],
]

const PERIOD_HEADERS = ['31-March-2026', '31-March-2025', 45382] // 45382 = 31-03-2024 (Excel serial)
const OPENING_RETAINED: number = -4000000

/** Signed TB per period with the bank plug and retained earnings filled in. */
function signedAmounts(): { ledgers: DemoLedger[] } {
  const ledgers = LEDGERS.map((l) => ({ ...l, amt: [...l.amt] as Amt }))
  const isPL = (l: DemoLedger) => l.code.startsWith('D3') || l.code.startsWith('D4')
  const re = ledgers.find((l) => l.code === 'D2002')!
  const bank = ledgers.find((l) => l.code === 'D1301')!
  // Opening retained earnings roll forward with profit (no dividends): RE(k) = RE(k-1) + P&L(k-1).
  const plSum = (i: number) => ledgers.filter(isPL).reduce((a, l) => a + l.amt[i], 0)
  re.amt[2] = OPENING_RETAINED
  re.amt[1] = re.amt[2] + plSum(2)
  re.amt[0] = re.amt[1] + plSum(1)
  for (let i = 0; i < 3; i++) {
    const others = ledgers.filter((l) => l !== bank).reduce((a, l) => a + l.amt[i], 0) + SPLIT.amt[i] + RENAMED.signed.amt[i]
    bank.amt[i] = -others
  }
  return { ledgers }
}

export function demoSignedGrid(): Grid {
  const { ledgers } = signedAmounts()
  const grid: Grid = [
    [null, 'Debits Balance - Positive(+) Credit Balance - Negative(-)', null, null, null, null],
    ['Demo company (synthetic)', null, null, null, '', 0],
    ['Name of Company', 'System Primary Grouping', 'System Parent Grouping', 'Final Codes', 'Name of Ledger', PERIOD_HEADERS[0],
      'Face Grouping_Current Year', 'Note Grouping_Current Year', 'Sub-Note Grouping_Current Year', PERIOD_HEADERS[1],
      'Face Grouping_\nPrevious Year', 'Note Grouping_\nPrevious Year', PERIOD_HEADERS[2], 'Face Grouping_Preceding Previous Year'],
  ]
  const push = (code: string, name: string, primary: string, face: string, note: string, amt: Amt) =>
    grid.push(['', primary, primary, null, `${code}/${name}`, amt[0], face, note, null, amt[1], face, note, amt[2], face])
  for (const l of ledgers) push(l.code, l.name, l.primary, l.face, l.note, l.amt)
  push(SPLIT.code, SPLIT.name, SPLIT.primary, SPLIT.face, SPLIT.note, SPLIT.amt)
  push(RENAMED.signed.code, RENAMED.signed.name, 'Sundry Creditors', 'Trade payables due to others', 'Due to others', RENAMED.signed.amt)
  return grid
}

interface CopyLine { ref: string; code: string; ledger: string; category: string; sub: string; line: string; layer: 'G' | 'A' | 'P'; local: Amt; jv: Amt }

function copyLines(): CopyLine[] {
  const { ledgers } = signedAmounts()
  const out: CopyLine[] = []
  let n = 1
  const zero: Amt = [0, 0, 0]
  for (const l of ledgers) {
    out.push({ ref: String(n++), code: l.code, ledger: `${l.code}/${l.name}`, category: l.category, sub: l.sub, line: l.line, layer: 'G', local: l.copy ?? l.amt, jv: zero })
  }
  out.push({ ref: String(n++), code: SPLIT.code, ledger: `${SPLIT.code}/${SPLIT.name}`, category: 'Provisions - non current', sub: 'Provisions', line: 'Provision for gratuity', layer: 'G', local: SPLIT.copyNC, jv: zero })
  out.push({ ref: String(n++), code: SPLIT.code, ledger: `${SPLIT.code}/${SPLIT.name}`, category: 'Provisions - current', sub: 'Provisions', line: 'Provision for gratuity', layer: 'G', local: SPLIT.copyC, jv: zero })
  out.push({ ref: String(n), code: RENAMED.copy.code, ledger: `${RENAMED.copy.code}/${RENAMED.copy.name}`, category: 'Trade Payables - current', sub: 'Trade payables', line: 'Dues of creditors other than MSME', layer: 'G', local: RENAMED.copy.amt, jv: zero })
  // A duplicate zero row, as audit working rows often appear.
  out.push({ ref: 'AUDIT', code: 'D4301', ledger: 'D4301/DEMO LEDGER AC - DEPRECIATION', category: 'Depreciation and amortization expense', sub: 'Depreciation', line: 'Depreciation on property, plant and equipment', layer: 'G', local: zero, jv: zero })
  for (const [ref, layer, narr, category, line, jv] of ADJUSTMENTS) {
    out.push({ ref: `ADJ-${ref}`, code: layer === 'A' ? 'ADJ-BS' : 'ADJ-PL', ledger: narr, category, sub: category, line, layer, local: zero, jv })
  }
  return out
}

export function demoCopyTBGrid(): Grid {
  const lines = copyLines()
  const grid: Grid = [
    ['Demo company (synthetic) - Ind AS workings'],
    ['Trial Balance (source sheet)', null, null, null, null, null, null, 'FY 2025-26', null, null, 'FY 2024-25', null, null, 'FY 2023-24', null, null],
    ['Amounts are in rupees.', null, null, null, null, null, null, 'Balances as at 31 March 2026', null, null, 'Balances as at 31 March 2025', null, null, 'Balances as at 1 April 2024 (Ind AS transition date)', null, null],
    ['#', 'Account code', 'Ledger (code / name)', 'Statement category', 'Sub-heading', 'Line item', 'Layer', 'Local GAAP', 'Ind AS JV', 'Ind AS', 'Local GAAP', 'Ind AS JV', 'Ind AS', 'Local GAAP', 'Ind AS JV', 'Ind AS'],
  ]
  for (const l of lines) {
    const cols: unknown[] = [l.ref, l.code, l.ledger, l.category, l.sub, l.line, l.layer]
    for (let i = 0; i < 3; i++) cols.push(l.local[i], l.jv[i], l.local[i] + l.jv[i])
    grid.push(cols)
    if (l.ref === 'AUDIT') grid.push(['IND AS ADJUSTMENT VOUCHERS (synthetic)'])
  }
  // A memo line that carries a layer letter but no account code: the parser must skip it.
  grid.push([null, null, 'Memo: P&L restatement', null, null, null, 'P', null, 12345, 12345])
  return grid
}

// ---------------------------------------------------------------- statements (computed like the workbook's SUMIFS)

function copyStatementValues() {
  const lines = copyLines()
  const by = (pred: (l: CopyLine) => boolean, i: number, useJv = true) =>
    lines.filter(pred).reduce((a, l) => a + l.local[i] + (useJv ? l.jv[i] : 0), 0)
  const cat = (c: string) => (l: CopyLine) => l.category === c
  const ga = (c: string) => (l: CopyLine) => l.category === c && (l.layer === 'G' || l.layer === 'A')
  const gp = (c: string) => (l: CopyLine) => l.category === c && (l.layer === 'G' || l.layer === 'P')
  const PL_CATS = ['Revenue From Operations', 'Other Income', 'Cost of materials consumed', 'Changes in inventories of finished goods, Stock in Trade and work in progress', 'Employee benefits expense', 'Finance costs', 'Depreciation and amortization expense', 'Other expenses', 'Tax expenses']
  const result = [0, 1, 2].map((i) => {
    const bs = (c: string) => by(ga(c), i)
    const pl = (c: string) => by(gp(c), i)
    const plLocal = PL_CATS.reduce((a, c) => a + by((l) => cat(c)(l) && l.layer === 'G', i), 0)
    const ppe = bs('Property, Plant and Equipment'), cwip = bs('Capital work in progress'), rou = bs('Right of Use Assets'), ncofa = bs('Other financial assets - non current')
    const inv = bs('Inventories'), rec = bs('Trade receivables - current'), cash = bs('Cash and cash equivalents'), ob = bs('Other bank balances'), ctax = bs('Income Tax Assets'), oca = bs('Other current assets')
    const sc = -bs('Equity Share Capital'), oe = -(bs('Other Equity') + plLocal)
    const nclease = -bs('Lease liabilities non-current'), ncprov = -bs('Provisions - non current'), dtl = -bs('Deferred tax liabilities, net')
    const clease = -bs('Lease liabilities current'), tp = -bs('Trade Payables - current'), cofl = -bs('Other financial liabilities - current'), ocl = -bs('Other current liabilities'), cprov = -bs('Provisions - current')
    const assets = ppe + cwip + rou + ncofa + inv + rec + cash + ob + ctax + oca
    const equity = sc + oe
    const eql = equity + nclease + ncprov + dtl + clease + tp + cofl + ocl + cprov
    const rev = -pl('Revenue From Operations'), oi = -pl('Other Income')
    const exp = ['Cost of materials consumed', 'Changes in inventories of finished goods, Stock in Trade and work in progress', 'Employee benefits expense', 'Finance costs', 'Depreciation and amortization expense', 'Other expenses'].map(pl)
    const pbt = rev + oi - exp.reduce((a, b) => a + b, 0)
    const tax = pl('Tax expenses')
    const pat = pbt - tax
    const oci = i === 0 ? -44899 : i === 1 ? -29933 : 0
    return { ppe, cwip, rou, ncofa, inv, rec, cash, ob, ctax, oca, sc, oe, nclease, ncprov, dtl, clease, tp, cofl, ocl, cprov, assets, equity, eql, rev, oi, exp, pbt, tax, pat, oci, tci: pat + oci }
  })
  return result
}

const lakhs = (v: number) => v / 1e5

export function demoBalanceSheetGrid(): Grid {
  const v = copyStatementValues()
  const row = (label: string, pick: (x: (typeof v)[number]) => number): unknown[] => [label, null, null, lakhs(pick(v[0])), null, lakhs(pick(v[1])), null, lakhs(pick(v[2]))]
  return [
    ['Demo company (synthetic)'],
    ['Balance Sheet as at 31 March 2026'],
    ['(All amounts in ₹ lakhs, unless otherwise stated)'],
    ['Particulars', null, 'Note No.', '31 Mar 2026', null, '31 Mar 2025', null, '1 Apr 2024'],
    ['ASSETS'],
    row('   Property, plant and equipment', (x) => x.ppe),
    row('   Capital work-in-progress', (x) => x.cwip),
    row('   Right-of-use assets', (x) => x.rou),
    row('   Financial assets — Other financial assets (non-current)', (x) => x.ncofa),
    row('   Inventories', (x) => x.inv),
    row('   Financial assets — Trade receivables', (x) => x.rec),
    row('   Financial assets — Cash and cash equivalents', (x) => x.cash),
    row('   Financial assets — Bank balances other than cash and cash equivalents', (x) => x.ob),
    row('   Current tax assets (net)', (x) => x.ctax),
    row('   Other current assets', (x) => x.oca),
    row('Total assets', (x) => x.assets),
    row('   Equity share capital', (x) => x.sc),
    row('   Other equity', (x) => x.oe),
    row('Total equity', (x) => x.equity),
    row('   Financial liabilities — Lease liabilities (non-current)', (x) => x.nclease),
    row('   Provisions (non-current)', (x) => x.ncprov),
    row('   Deferred tax liabilities (net)', (x) => x.dtl),
    row('   Financial liabilities — Lease liabilities (current)', (x) => x.clease),
    row('   Financial liabilities — Trade payables', (x) => x.tp),
    row('   Financial liabilities — Other financial liabilities (current)', (x) => x.cofl),
    row('   Other current liabilities', (x) => x.ocl),
    row('   Provisions (current)', (x) => x.cprov),
    row('Total equity and liabilities', (x) => x.eql),
  ]
}

export function demoProfitLossGrid(): Grid {
  const v = copyStatementValues()
  const row = (label: string, pick: (x: (typeof v)[number]) => number): unknown[] => [label, null, lakhs(pick(v[0])), null, lakhs(pick(v[1]))]
  const names = ['Cost of materials consumed', 'Changes in inventories of finished goods and WIP', 'Employee benefits expense', 'Finance costs', 'Depreciation and amortisation expense', 'Other expenses']
  return [
    ['Demo company (synthetic)'],
    ['Statement of Profit and Loss for the year ended 31 March 2026'],
    ['(All amounts in ₹ lakhs, unless otherwise stated)'],
    ['Particulars', 'Note', 'FY 2025-26', null, 'FY 2024-25'],
    row('   Revenue from operations', (x) => x.rev),
    row('   Other income', (x) => x.oi),
    ...names.map((n, i) => row(`   ${n}`, (x) => x.exp[i])),
    row('Profit before tax', (x) => x.pbt),
    row('Tax expense', (x) => x.tax),
    row('Profit for the year (A)', (x) => x.pat),
    ['Other comprehensive income (B)'],
    row('Items that will not be reclassified to profit or loss — remeasurement of defined benefit plans, net of tax', (x) => x.oci),
    row('Total comprehensive income for the year (A+B)', (x) => x.tci),
  ]
}

export function demoEntriesGrid(): Grid {
  const grid: Grid = [
    ['Demo company (synthetic) - Ind AS adjustment vouchers'],
    ['All amounts in ₹ lakhs'],
    ['Ref', 'TB row', 'Voucher', 'Account / narration', 'Statement line', 'Layer', '31-03-2026', '31-03-2025', '01-04-2024', 'Ind AS reference', 'Why the entry exists (plain English)'],
  ]
  const tb = demoCopyTBGrid()
  const why: Record<string, string> = {
    'Ind AS 116': 'The factory lease is brought on balance sheet: a right-of-use asset and a lease liability replace the rent expense.',
    'Ind AS 109': 'The interest-free deposit is carried at amortised cost; the day-one difference is prepaid rent.',
    'Ind AS 19': 'Actuarial remeasurement of the gratuity obligation goes to other comprehensive income, not profit.',
    'Ind AS 12': 'Deferred tax on the temporary differences created by the adjustments above.',
    'Ind AS 101': 'Opening Ind AS balance sheet at the transition date.',
  }
  let section = ''
  tb.forEach((row, i) => {
    const layer = row[6]
    if (layer !== 'A' && layer !== 'P') return
    if (!row[1]) return
    const narr = String(row[2])
    const std = /Ind AS \d+/.exec(narr)?.[0] ?? 'Retained earnings'
    if (std !== section) {
      grid.push([std])
      section = std
    }
    grid.push([row[0], i + 1, row[1], narr, row[5], layer, lakhs(Number(row[8])), lakhs(Number(row[11])), lakhs(Number(row[14])), std, why[std] ?? 'Carries the profit and loss adjustments into retained earnings.'])
  })
  return grid
}

export const DEMO_COMPANY = 'Demo Manufacturing Pvt Ltd (synthetic data)'
