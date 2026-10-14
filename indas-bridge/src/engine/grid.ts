/** Small helpers for locating headers by text in a 2-D sheet grid. */
import type { Grid, PeriodKey } from './types'

export function cellText(v: unknown): string {
  if (v === null || v === undefined) return ''
  if (typeof v === 'string') return v.replace(/\s+/g, ' ').trim()
  if (typeof v === 'number' || typeof v === 'boolean') return String(v)
  if (v instanceof Date) return v.toISOString()
  return String(v)
}

/** Lower-cased, single-spaced, hyphen-normalised text for matching. */
export function norm(v: unknown): string {
  return cellText(v).toLowerCase().replace(/[‐-―]/g, '-').replace(/_/g, ' ').replace(/\s+/g, ' ').trim()
}

/** Join key for a ledger: whitespace collapsed, upper case, spaces around "/" removed. */
export function ledgerKey(s: unknown): string {
  return cellText(s).toUpperCase().replace(/\s*\/\s*/g, '/').replace(/\s+/g, ' ').trim()
}

/** Leading account code of a "code/name" ledger string (digits or alphanumerics before the first "/"). */
export function ledgerCode(s: unknown): string {
  const m = /^\s*([A-Za-z0-9-]+)\s*\//.exec(cellText(s))
  return m ? m[1].toUpperCase() : ''
}

/** First row index (within maxRows) for which `predicate(normalisedCells)` is true. */
export function findRow(grid: Grid, predicate: (cells: string[]) => boolean, maxRows = 60): number {
  const limit = Math.min(grid.length, maxRows)
  for (let r = 0; r < limit; r++) {
    const row = grid[r] ?? []
    if (predicate(row.map(norm))) return r
  }
  return -1
}

export function findCol(cells: string[], test: RegExp, from = 0): number {
  for (let c = from; c < cells.length; c++) if (test.test(cells[c])) return c
  return -1
}

const MONTHS: Record<string, number> = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
}

/** Excel date serial (1900 system) to {y,m,d}. */
export function serialToYMD(serial: number): { y: number; m: number; d: number } {
  const ms = Math.round((serial - 25569) * 86400 * 1000)
  const dt = new Date(ms)
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() }
}

/** Parse a header cell that holds a date (text like "31-March-2026", a Date, or an Excel serial). */
export function headerDate(v: unknown): { y: number; m: number; d: number } | null {
  if (v instanceof Date) return { y: v.getUTCFullYear(), m: v.getUTCMonth() + 1, d: v.getUTCDate() }
  if (typeof v === 'number' && v > 20000 && v < 80000 && Number.isInteger(v)) return serialToYMD(v)
  if (typeof v === 'string') {
    const t = v.trim().toLowerCase()
    const m = /^(\d{1,2})[\s\-./]*([a-z]{3})[a-z]*[\s\-./,]*(\d{4})$/.exec(t)
    if (m && MONTHS[m[2]]) return { y: Number(m[3]), m: MONTHS[m[2]], d: Number(m[1]) }
    const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(t)
    if (iso) return { y: Number(iso[1]), m: Number(iso[2]), d: Number(iso[3]) }
    const dmy = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/.exec(t)
    if (dmy) return { y: Number(dmy[3]), m: Number(dmy[2]), d: Number(dmy[1]) }
  }
  return null
}

/** Period key for a balance-sheet date: the year in which that financial year (April-March) ends. */
export function periodForDate(d: { y: number; m: number; d: number }): PeriodKey {
  // 1 April of year Y is the opening of FY ending Y+1, but as a balance sheet date it equals the
  // closing of FY ending Y (31 March Y). We key opening balances to the earlier year.
  if (d.m === 4 && d.d === 1) return d.y
  return d.m <= 3 ? d.y : d.y + 1
}

/** Period key from text like "FY 2025-26", "2025-26", "31 Mar 2026" or "1 Apr 2024". */
export function periodFromLabel(v: unknown): PeriodKey | null {
  if (v instanceof Date || typeof v === 'number') {
    const d0 = headerDate(v)
    return d0 ? periodForDate(d0) : null
  }
  const t = cellText(v)
  const d = headerDate(t.replace(/\s+/g, '-'))
  if (d) return periodForDate(d)
  const fy = /(?:FY\s*)?\b(\d{4})\s*[-–]\s*(\d{2}|\d{4})\b/i.exec(t)
  if (fy) return Number(fy[1]) + 1
  return null
}

export function fyLabel(key: PeriodKey): string {
  return `FY ${key - 1}-${String(key % 100).padStart(2, '0')}`
}
