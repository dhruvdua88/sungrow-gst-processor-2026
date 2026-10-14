/**
 * Bridge engine: from the joined ledgers and the Ind AS JV layer, build for each period
 *  - the balance sheet and P&L bridge tables (Local GAAP | Reclass | Ind AS adjustment | Ind AS),
 *  - the profit waterfall (Local GAAP PAT -> adjustments -> Ind AS PAT -> OCI -> TCI),
 *  - the equity waterfall (Local GAAP equity -> adjustments by standard -> Ind AS equity),
 *  - tie-out checks against the Ind AS workbook's own face statements.
 *
 * Internally everything is signed (debit +, credit −). Bridge lines carry presentation-sign amounts:
 * assets and expenses positive, equity, liabilities and income positive (credit flipped).
 */
import { CAPTION_BY_ID, CAPTIONS, CREDIT_SECTIONS, classifyIndAsCategory, normaliseText } from './captions'
import type { Adjustment, JoinedLedger, JoinResult } from './join'
import { VARIANCE_TOLERANCE } from './join'
import type { Paise } from './money'
import type { Caption, CopyModel, Period, PeriodKey, SectionId, StatementFigure } from './types'

export interface BridgeLine {
  caption: Caption
  /** Presentation sign. */
  local: Paise
  reclass: Paise
  adj: Paise
  indas: Paise
  /** Copy face-statement figure for this caption, when one could be matched. */
  reference: Paise | null
  ledgerCount: number
  adjustmentRows: number[]
}

export type TieStatus = 'tie' | 'rounding' | 'diff' | 'na'

export interface Tie {
  id: string
  label: string
  engine: Paise | null
  reference: Paise | null
  diff: Paise | null
  status: TieStatus
  note?: string
}

export interface Step {
  id: string
  label: string
  amount: Paise
  kind: 'start' | 'step' | 'total'
  detail?: string
  rows?: number[]
  tone?: 'neutral' | 'warn'
}

export interface LineItemNode {
  id: string
  captionId: string
  category: string
  lineItem: string
  /** Signed. */
  local: Paise
  jv: Paise
  ledgerKeys: string[]
  adjustmentRows: number[]
}

export interface PeriodBridge {
  period: Period
  prior: Period | null
  bsLines: BridgeLine[]
  plLines: BridgeLine[]
  totals: {
    local: StatementTotals
    indas: StatementTotals
  }
  plWaterfall: Step[] | null
  equityWaterfall: Step[]
  ties: Tie[]
  lineItems: LineItemNode[]
}

export interface StatementTotals {
  assets: Paise
  equity: Paise
  liabilities: Paise
  unmappedBS: Paise
  income: Paise
  expenses: Paise
  pbt: Paise
  tax: Paise
  pat: Paise
  oci: Paise | null
  tci: Paise | null
}

const ROUNDING_LIMIT: Paise = 100_000 // ₹1,000

export function tieStatus(diff: Paise | null): TieStatus {
  if (diff === null) return 'na'
  const a = Math.abs(diff)
  if (a <= VARIANCE_TOLERANCE) return 'tie'
  if (a <= ROUNDING_LIMIT) return 'rounding'
  return 'diff'
}

export function presentation(section: SectionId, signed: Paise): Paise {
  return CREDIT_SECTIONS.has(section) ? -signed : signed
}

const isPL = (c: Caption | undefined) => c?.statement === 'PL'

interface Sums {
  byLocal: Map<string, Paise>
  byIndAs: Map<string, Paise>
  adj: Map<string, Paise>
  ledgerCount: Map<string, number>
  adjRows: Map<string, number[]>
}

function add(m: Map<string, Paise>, k: string, v: Paise) {
  m.set(k, (m.get(k) ?? 0) + v)
}

function sums(join: JoinResult, k: PeriodKey): Sums {
  const s: Sums = { byLocal: new Map(), byIndAs: new Map(), adj: new Map(), ledgerCount: new Map(), adjRows: new Map() }
  for (const l of join.ledgers) {
    add(s.byLocal, l.localCaption[k], l.used[k] ?? 0)
    const seen = new Set<string>()
    for (const a of l.alloc[k] ?? []) {
      add(s.byIndAs, a.captionId, a.amount)
      if (a.amount !== 0 && !seen.has(a.captionId)) {
        seen.add(a.captionId)
        s.ledgerCount.set(a.captionId, (s.ledgerCount.get(a.captionId) ?? 0) + 1)
      }
    }
  }
  for (const a of join.adjustments) {
    const v = a.row.jv[k] ?? 0
    add(s.adj, a.captionId, v)
    if (v !== 0) s.adjRows.set(a.captionId, [...(s.adjRows.get(a.captionId) ?? []), a.row.sheetRow])
  }
  // Profit for the year sits in other equity on the balance sheet (local GAAP and Ind AS-mapped).
  // The P layer is NOT added: the workbook carries its effect on equity in layer A.
  const plSum = (m: Map<string, Paise>) => [...m].filter(([c]) => isPL(CAPTION_BY_ID.get(c))).reduce((a, [, v]) => a + v, 0)
  add(s.byLocal, 'otherequity', plSum(s.byLocal))
  add(s.byIndAs, 'otherequity', plSum(s.byIndAs))
  return s
}

function lines(s: Sums, kind: 'BS' | 'PL', refs: Map<string, Paise>): BridgeLine[] {
  const out: BridgeLine[] = []
  for (const cap of CAPTIONS) {
    if (cap.statement !== kind) continue
    const local = s.byLocal.get(cap.id) ?? 0
    const mapped = s.byIndAs.get(cap.id) ?? 0
    const adj = s.adj.get(cap.id) ?? 0
    if (local === 0 && mapped === 0 && adj === 0 && !refs.has(cap.id)) continue
    const p = (v: Paise) => presentation(cap.section, v)
    out.push({
      caption: cap,
      local: p(local),
      reclass: p(mapped - local),
      adj: p(adj),
      indas: p(mapped + adj),
      reference: refs.get(cap.id) ?? null,
      ledgerCount: s.ledgerCount.get(cap.id) ?? 0,
      adjustmentRows: s.adjRows.get(cap.id) ?? [],
    })
  }
  return out
}

function totalsFrom(get: (id: string) => Paise): Omit<StatementTotals, 'oci' | 'tci'> {
  let assets = 0, eq = 0, liab = 0, unmappedBS = 0, income = 0, expenses = 0, tax = 0
  for (const cap of CAPTIONS) {
    const v = get(cap.id)
    if (!v) continue
    switch (cap.section) {
      case 'NCA': case 'CA': assets += v; break
      case 'EQ': eq += -v; break
      case 'NCL': case 'CL': liab += -v; break
      case 'INC': income += -v; break
      case 'EXP': expenses += v; break
      case 'TAX': tax += v; break
      case 'UNMAPPED':
        if (cap.statement === 'PL') expenses += v
        else unmappedBS += v
        break
    }
  }
  const pbt = income - expenses
  return { assets, equity: eq, liabilities: liab, unmappedBS, income, expenses, pbt, tax, pat: pbt - tax }
}

// ---------------------------------------------------------------- statement references

function stripPrefix(label: string): string {
  return label.replace(/^\s*(financial (assets|liabilities)\s*[—–-]\s*)/i, '')
}

export function statementCaptionRefs(figs: StatementFigure[], k: PeriodKey): Map<string, Paise> {
  const out = new Map<string, Paise>()
  for (const f of figs) {
    const t = normaliseText(f.label)
    if (/^(total|check|profit|loss|earnings|items that|other comprehensive)/.test(t)) continue
    const v = f.values[k]
    if (v === undefined) continue
    const cap = classifyIndAsCategory(stripPrefix(f.label))
    if (cap) out.set(cap, (out.get(cap) ?? 0) + v)
  }
  return out
}

export function findFigure(figs: StatementFigure[], test: RegExp, k: PeriodKey): Paise | null {
  for (const f of figs) if (test.test(normaliseText(f.label)) && f.values[k] !== undefined) return f.values[k]
  return null
}

// ---------------------------------------------------------------- waterfalls

function shortLine(text: string): string {
  return text.replace(/\s*\((Ind AS[^)]*|Note[^)]*)\)\s*/gi, ' ').replace(/\s+/g, ' ').trim()
}

function plSteps(join: JoinResult, k: PeriodKey, localPat: Paise, s: Sums): Step[] {
  const steps: Step[] = [{ id: 'local-pat', label: 'Local GAAP profit after tax', amount: localPat, kind: 'start', detail: 'From the signed books' }]
  // Reclassification between balance sheet and P&L, if any ledger changes statement.
  const plLocal = [...s.byLocal].filter(([c]) => isPL(CAPTION_BY_ID.get(c))).reduce((a, [, v]) => a + v, 0)
  const plMapped = [...s.byIndAs].filter(([c]) => isPL(CAPTION_BY_ID.get(c))).reduce((a, [, v]) => a + v, 0)
  if (Math.abs(plMapped - plLocal) > VARIANCE_TOLERANCE) {
    steps.push({ id: 'reclass', label: 'Ledgers moved between balance sheet and P&L', amount: -(plMapped - plLocal), kind: 'step', tone: 'warn' })
  }
  const groups = new Map<string, { label: string; amount: Paise; rows: number[]; narr: string[] }>()
  for (const a of join.adjustments) {
    if (!isPL(CAPTION_BY_ID.get(a.captionId))) continue
    const v = a.row.jv[k] ?? 0
    if (v === 0) continue
    const label = `${a.standard} · ${shortLine(a.row.lineItem || a.row.category)}`
    const g = groups.get(label) ?? { label, amount: 0, rows: [], narr: [] }
    g.amount += -v
    g.rows.push(a.row.sheetRow)
    g.narr.push(a.row.ledger)
    groups.set(label, g)
  }
  for (const g of groups.values()) {
    if (g.amount === 0) continue
    steps.push({ id: `adj:${g.label}`, label: g.label, amount: g.amount, kind: 'step', rows: g.rows, detail: g.narr.join('\n') })
  }
  return steps
}

function equityGroups(adjs: Adjustment[], k: PeriodKey): Step[] {
  const groups = new Map<string, { amount: Paise; rows: number[]; captions: Set<string> }>()
  for (const a of adjs) {
    const cap = CAPTION_BY_ID.get(a.captionId)
    if (!cap || cap.statement !== 'BS' || cap.section === 'EQ') continue
    const v = a.row.jv[k] ?? 0
    if (v === 0) continue
    const g = groups.get(a.standard) ?? { amount: 0, rows: [], captions: new Set<string>() }
    g.amount += v // a debit to an asset raises equity; a credit to a liability lowers it
    g.rows.push(a.row.sheetRow)
    g.captions.add(cap.label)
    groups.set(a.standard, g)
  }
  return [...groups]
    .filter(([, g]) => g.amount !== 0)
    .sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }))
    .map(([std, g]) => ({ id: `eq:${std}`, label: std, amount: g.amount, kind: 'step' as const, rows: g.rows, detail: [...g.captions].join(' · ') }))
}

// ---------------------------------------------------------------- main

export function buildBridge(join: JoinResult, copy: CopyModel, k: PeriodKey): PeriodBridge {
  const period = join.periods.find((p) => p.key === k)
  if (!period) throw new Error(`Unknown period ${k}`)
  const prior = join.periods.find((p) => p.key === k - 1) ?? null
  const s = sums(join, k)

  const bsRefs = statementCaptionRefs(copy.statements.bs, k)
  const plRefs = statementCaptionRefs(copy.statements.pl, k)
  const bsLines = lines(s, 'BS', bsRefs)
  const plLines = period.hasPL ? lines(s, 'PL', plRefs) : []

  const local = { ...totalsFrom((id) => s.byLocal.get(id) ?? 0), oci: null, tci: null } as StatementTotals
  const indas = { ...totalsFrom((id) => (s.byIndAs.get(id) ?? 0) + (s.adj.get(id) ?? 0)), oci: null, tci: null } as StatementTotals
  // Ind AS profit uses the P layer; the BS totals above already exclude it (P captions are PL only).

  // OCI and TCI from the movement in other equity (needs the prior period).
  if (prior && period.hasPL) {
    const sp = sums(join, prior.key)
    const oe = (m: Sums, ind: boolean) => -((ind ? (m.byIndAs.get('otherequity') ?? 0) + (m.adj.get('otherequity') ?? 0) : (m.byLocal.get('otherequity') ?? 0)))
    const otherMovements = oe(s, false) - oe(sp, false) - local.pat
    indas.tci = oe(s, true) - oe(sp, true) - otherMovements
    indas.oci = indas.tci - indas.pat
    local.tci = local.pat
    local.oci = 0
  }

  let plWaterfall: Step[] | null = null
  if (period.hasPL) {
    plWaterfall = plSteps(join, k, local.pat, s)
    plWaterfall.push({ id: 'indas-pat', label: 'Ind AS profit after tax', amount: indas.pat, kind: 'total' })
    if (indas.oci !== null && indas.tci !== null) {
      plWaterfall.push({ id: 'oci', label: 'Other comprehensive income', amount: indas.oci, kind: 'step', detail: 'Derived: movement in Ind AS other equity less profit and other equity movements' })
      plWaterfall.push({ id: 'tci', label: 'Total comprehensive income', amount: indas.tci, kind: 'total' })
    }
  }

  const equitySteps: Step[] = [{ id: 'local-eq', label: 'Local GAAP total equity', amount: local.equity, kind: 'start' }]
  const mappedEq = -CAPTIONS.filter((c) => c.section === 'EQ').reduce((a, c) => a + (s.byIndAs.get(c.id) ?? 0), 0)
  if (Math.abs(mappedEq - local.equity) > VARIANCE_TOLERANCE) equitySteps.push({ id: 'eq-reclass', label: 'Reclassified into / out of equity', amount: mappedEq - local.equity, kind: 'step', tone: 'warn' })
  const groups = equityGroups(join.adjustments, k)
  equitySteps.push(...groups)
  const explained = mappedEq + groups.reduce((a, g) => a + g.amount, 0)
  if (Math.abs(explained - indas.equity) > VARIANCE_TOLERANCE) {
    equitySteps.push({ id: 'eq-unexplained', label: 'Equity entries without an asset / liability side', amount: indas.equity - explained, kind: 'step', tone: 'warn' })
  }
  // Fold sub-rupee float residue into the last step so the waterfall lands exactly on the total.
  const running = equitySteps.reduce((a, st) => a + st.amount, 0)
  if (running !== indas.equity && equitySteps.length > 1) equitySteps[equitySteps.length - 1].amount += indas.equity - running
  equitySteps.push({ id: 'indas-eq', label: 'Ind AS total equity', amount: indas.equity, kind: 'total' })

  const ties = buildTies(join, copy, k, period, indas, local, s)
  return { period, prior, bsLines, plLines, totals: { local, indas }, plWaterfall, equityWaterfall: equitySteps, ties, lineItems: lineItemNodes(join, k) }
}

function tie(id: string, label: string, engine: Paise | null, reference: Paise | null, note?: string): Tie {
  const diff = engine !== null && reference !== null ? engine - reference : null
  return { id, label, engine, reference, diff, status: tieStatus(diff), note }
}

function buildTies(join: JoinResult, copy: CopyModel, k: PeriodKey, period: Period, indas: StatementTotals, local: StatementTotals, s: Sums): Tie[] {
  const out: Tie[] = []
  const signedTotal = join.ledgers.reduce((a, l) => a + (l.used[k] ?? 0), 0)
  out.push(tie('signed-nil', 'Signed trial balance nets to nil', signedTotal, 0))
  const aTotal = join.adjustments.filter((a) => a.row.layer === 'A').reduce((x, a) => x + (a.row.jv[k] ?? 0), 0)
  out.push(tie('a-nil', 'Ind AS balance-sheet adjustments (layer A) net to nil', aTotal, 0))
  out.push(tie('bs-balances', 'Ind AS balance sheet balances (assets = equity + liabilities)', indas.assets + indas.unmappedBS, indas.equity + indas.liabilities))
  out.push(tie('local-bs-balances', 'Local GAAP balance sheet balances', local.assets + local.unmappedBS, local.equity + local.liabilities))
  const bs = copy.statements.bs
  const pl = copy.statements.pl
  if (bs.length) {
    out.push(tie('fs-assets', 'Total assets vs Ind AS Balance Sheet', indas.assets, findFigure(bs, /^total assets$/, k)))
    out.push(tie('fs-equity', 'Total equity vs Ind AS Balance Sheet', indas.equity, findFigure(bs, /^total equity$/, k)))
    out.push(tie('fs-eql', 'Total equity and liabilities vs Ind AS Balance Sheet', indas.equity + indas.liabilities, findFigure(bs, /^total equity and liabilities$/, k)))
  }
  if (pl.length && period.hasPL) {
    out.push(tie('fs-pbt', 'Profit before tax vs Statement of P&L', indas.pbt, findFigure(pl, /^profit before tax/, k)))
    out.push(tie('fs-tax', 'Tax expense vs Statement of P&L', indas.tax, findFigure(pl, /^(total )?tax expense/, k)))
    out.push(tie('fs-pat', 'Profit for the year vs Statement of P&L', indas.pat, findFigure(pl, /^profit (for the (year|period)|after tax)/, k)))
    out.push(tie('fs-oci', 'Other comprehensive income vs Statement of P&L', indas.oci, findFigure(pl, /^(?!total).*(items that will not be reclassified|items that will be reclassified|other comprehensive income)/, k)))
    out.push(tie('fs-tci', 'Total comprehensive income vs Statement of P&L', indas.tci, findFigure(pl, /^total comprehensive income/, k)))
  }
  void s
  return out
}

function lineItemNodes(join: JoinResult, k: PeriodKey): LineItemNode[] {
  const map = new Map<string, LineItemNode>()
  const get = (captionId: string, category: string, lineItem: string) => {
    const id = `${captionId}|${normaliseText(lineItem || category)}`
    let n = map.get(id)
    if (!n) {
      n = { id, captionId, category, lineItem: lineItem || category, local: 0, jv: 0, ledgerKeys: [], adjustmentRows: [] }
      map.set(id, n)
    }
    return n
  }
  for (const l of join.ledgers as JoinedLedger[]) {
    for (const a of l.alloc[k] ?? []) {
      const n = get(a.captionId, a.category, a.lineItem)
      n.local += a.amount
      if (a.amount !== 0 || l.copyRows.length <= 1) n.ledgerKeys.push(l.key)
    }
  }
  for (const a of join.adjustments) {
    const v = a.row.jv[k] ?? 0
    if (v === 0) continue
    const n = get(a.captionId, a.row.category, a.row.lineItem)
    n.jv += v
    n.adjustmentRows.push(a.row.sheetRow)
  }
  return [...map.values()]
}
