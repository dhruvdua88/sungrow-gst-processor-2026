/**
 * Parser for the Ind AS workbook: the "TB" sheet (ledger -> Ind AS mapping plus the Ind AS JV layer),
 * the face statements (for tie-out only) and the "Ind AS Entries" voucher register (narrative only).
 * Headers are located by text.
 */
import { cellText, findCol, findRow, fyLabel, ledgerKey, norm, periodFromLabel } from './grid'
import { lakhsToPaise, rupeesToPaise, toNumber } from './money'
import type { CopyModel, CopyRow, Grid, Layer, Period, PeriodKey, StatementFigure, VoucherEntry } from './types'

export const COPY_TB_SHEETS = ['TB', 'Trial Balance']
export const COPY_BS_SHEETS = ['Balance Sheet']
export const COPY_PL_SHEETS = ['Statement of P&L', 'Statement of Profit and Loss', 'Profit and Loss']
export const COPY_ENTRY_SHEETS = ['Ind AS Entries', 'Ind AS Adjustments']

export interface CopyGrids {
  tb: Grid
  bs?: Grid
  pl?: Grid
  entries?: Grid
}

interface YearBlock {
  period: PeriodKey
  local: number
  jv: number
  indas: number
}

export function parseCopy(grids: CopyGrids): CopyModel {
  const warnings: string[] = []
  const { rows, periods } = parseCopyTB(grids.tb, warnings)
  const statements = {
    bs: grids.bs ? parseStatement(grids.bs) : [],
    pl: grids.pl ? parseStatement(grids.pl) : [],
  }
  if (!grids.bs) warnings.push('No "Balance Sheet" sheet found; face-statement tie-out is skipped.')
  if (!grids.pl) warnings.push('No "Statement of P&L" sheet found; face-statement tie-out is skipped.')
  const entries = grids.entries ? parseEntries(grids.entries) : []
  return { rows, periods, statements, entries, warnings }
}

function parseCopyTB(grid: Grid, warnings: string[]): { rows: CopyRow[]; periods: Period[] } {
  const h = findRow(grid, (c) => c.includes('layer') && c.some((x) => /ledger/.test(x)) && c.includes('local gaap'), 40)
  if (h < 0) throw new Error('Could not find the TB header (needs "Ledger", "Layer" and "Local GAAP" columns).')
  const cells = (grid[h] ?? []).map(norm)
  const refCol = findCol(cells, /^#$|^s\.? ?no|^ref/)
  const codeCol = findCol(cells, /account code|^code$/)
  const ledgerCol = findCol(cells, /ledger/)
  const catCol = findCol(cells, /statement category|^category$/)
  const subCol = findCol(cells, /sub-?heading/)
  const lineCol = findCol(cells, /line item/)
  const layerCol = findCol(cells, /^layer$/)

  // Year blocks: each "Local GAAP" header starts a block; the period label sits above it.
  const blocks: YearBlock[] = []
  const transitionHints = new Map<PeriodKey, boolean>()
  cells.forEach((c, col) => {
    if (c !== 'local gaap') return
    let period: PeriodKey | null = null
    let transition = false
    for (let r = h - 1; r >= Math.max(0, h - 4) && period === null; r--) {
      const above = grid[r] ?? []
      period = periodFromLabel(above[col])
      if (period !== null) {
        const below = grid[r + 1]?.[col]
        transition = /transition|1 april|1-apr|opening/i.test(cellText(below)) || /transition/i.test(cellText(above[col]))
      }
    }
    if (period === null) return
    const jv = findCol(cells, /^ind as jv$|^ind as adj/, col + 1)
    const indas = findCol(cells, /^ind as$/, col + 1)
    blocks.push({ period, local: col, jv, indas })
    transitionHints.set(period, transition)
  })
  if (blocks.length === 0) throw new Error('No year columns ("FY 2025-26" above "Local GAAP") found on the TB sheet.')
  blocks.sort((a, b) => b.period - a.period)

  const earliest = blocks[blocks.length - 1].period
  const periods: Period[] = blocks.map((b) => {
    const isTransition = transitionHints.get(b.period) || (blocks.length >= 3 && b.period === earliest)
    return isTransition
      ? { key: b.period, label: `1-Apr-${b.period} (transition)`, bsDate: `01-04-${b.period}`, hasPL: false }
      : { key: b.period, label: fyLabel(b.period), bsDate: `31-03-${b.period}`, hasPL: true }
  })

  const rows: CopyRow[] = []
  for (let r = h + 1; r < grid.length; r++) {
    const row = grid[r] ?? []
    const layer = cellText(row[layerCol]).toUpperCase()
    if (layer !== 'G' && layer !== 'A' && layer !== 'P') continue
    const code = cellText(row[codeCol])
    const ledger = cellText(row[ledgerCol])
    // Memo / check lines carry a layer letter but no account code: skip them.
    if (!code || !ledger) continue
    const local: Record<PeriodKey, number> = {}
    const jv: Record<PeriodKey, number> = {}
    const indas: Record<PeriodKey, number> = {}
    for (const b of blocks) {
      local[b.period] = rupeesToPaise(row[b.local])
      jv[b.period] = b.jv >= 0 ? rupeesToPaise(row[b.jv]) : 0
      indas[b.period] = b.indas >= 0 ? rupeesToPaise(row[b.indas]) : local[b.period] + jv[b.period]
    }
    rows.push({
      sheetRow: r + 1,
      ref: refCol >= 0 ? cellText(row[refCol]) : '',
      code,
      ledger,
      key: ledgerKey(ledger),
      category: cellText(row[catCol]),
      subHeading: subCol >= 0 ? cellText(row[subCol]) : '',
      lineItem: lineCol >= 0 ? cellText(row[lineCol]) : '',
      layer: layer as Layer,
      local,
      jv,
      indas,
    })
  }
  if (rows.length === 0) warnings.push('The TB sheet has no rows with layer G, A or P.')
  return { rows, periods }
}

function unitFactor(grid: Grid): (v: unknown) => number {
  for (let r = 0; r < Math.min(grid.length, 12); r++) {
    const t = (grid[r] ?? []).map(norm).join(' ')
    if (/in ₹? ?lakhs?|in lakhs?|rs\.? ?lakhs?/.test(t)) return lakhsToPaise
    if (/in ₹? ?crores?|in crores?/.test(t)) return (v) => Math.round((toNumber(v) ?? 0) * 1e9)
  }
  return rupeesToPaise
}

/** Face statement: label in the first text column, one numeric column per period header. */
export function parseStatement(grid: Grid): StatementFigure[] {
  const toPaise = unitFactor(grid)
  const h = findRow(grid, (c) => c.filter((x) => periodFromLabel(x) !== null && /20\d\d/.test(x)).length >= 1 && c.some((x) => /particulars/.test(x)), 20)
  if (h < 0) return []
  const header = grid[h] ?? []
  const cols: { col: number; period: PeriodKey }[] = []
  header.forEach((v, col) => {
    const p = periodFromLabel(v)
    if (p !== null && /20\d\d/.test(cellText(v))) cols.push({ col, period: p })
  })
  const out: StatementFigure[] = []
  for (let r = h + 1; r < grid.length; r++) {
    const row = grid[r] ?? []
    const label = cellText(row[0])
    if (!label) continue
    const values: Record<PeriodKey, number> = {}
    let any = false
    for (const { col, period } of cols) {
      if (toNumber(row[col]) !== null) {
        values[period] = toPaise(row[col])
        any = true
      }
    }
    if (any) out.push({ label, values })
  }
  return out
}

function parseEntries(grid: Grid): VoucherEntry[] {
  const h = findRow(grid, (c) => c.includes('ref') && c.includes('layer'), 30)
  if (h < 0) return []
  const raw = grid[h] ?? []
  const cells = raw.map(norm)
  const refCol = findCol(cells, /^ref$/)
  const tbCol = findCol(cells, /^tb row/)
  const vCol = findCol(cells, /^voucher/)
  const narCol = findCol(cells, /narration|account/)
  const lineCol = findCol(cells, /statement line/)
  const layerCol = findCol(cells, /^layer$/)
  const refStdCol = findCol(cells, /ind as reference|reference/)
  const whyCol = findCol(cells, /^why/)
  const toPaise = unitFactor(grid)
  const amountCols: { col: number; period: PeriodKey }[] = []
  raw.forEach((v, col) => {
    const p = periodFromLabel(v)
    if (p !== null && col > layerCol) amountCols.push({ col, period: p })
  })
  const out: VoucherEntry[] = []
  let section = ''
  for (let r = h + 1; r < grid.length; r++) {
    const row = grid[r] ?? []
    const layer = cellText(row[layerCol])
    const first = cellText(row[0])
    if (!layer) {
      const filled = row.filter((v) => cellText(v) !== '').length
      if (filled === 1 && first) section = first
      continue
    }
    if (!/^[A-Z]$/.test(layer)) continue
    const tb = toNumber(row[tbCol])
    const amounts: Record<PeriodKey, number> = {}
    for (const { col, period } of amountCols) amounts[period] = toPaise(row[col])
    out.push({
      ref: first || cellText(row[refCol]),
      tbRow: tb === null ? null : Math.round(tb),
      voucher: cellText(row[vCol]),
      narration: cellText(row[narCol]),
      statementLine: cellText(row[lineCol]),
      layer,
      amounts,
      reference: refStdCol >= 0 ? cellText(row[refStdCol]) : '',
      why: whyCol >= 0 ? cellText(row[whyCol]) : '',
      section,
    })
  }
  return out
}
