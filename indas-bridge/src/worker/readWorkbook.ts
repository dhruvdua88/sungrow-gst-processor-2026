/**
 * SheetJS boundary: turn a workbook file (xlsb / xlsx) into plain grids for the pure engine.
 * Runs inside the parse worker (and in Node for local checks). Only the sheets needed are parsed.
 */
import * as XLSX from 'xlsx'
import type { Grid } from '../engine/types'

export function sheetNames(buf: ArrayBuffer): string[] {
  const wb = XLSX.read(buf, { type: 'array', bookSheets: true })
  return wb.SheetNames
}

export function pickSheet(names: string[], candidates: string[]): string | undefined {
  const lower = names.map((n) => n.trim().toLowerCase())
  for (const c of candidates) {
    const i = lower.indexOf(c.toLowerCase())
    if (i >= 0) return names[i]
  }
  return undefined
}

export function readGrids(buf: ArrayBuffer, sheets: string[]): Record<string, Grid> {
  const wb = XLSX.read(buf, { type: 'array', sheets, dense: true, cellDates: false, cellFormula: false, cellHTML: false, cellText: false })
  const out: Record<string, Grid> = {}
  for (const name of sheets) {
    const ws = wb.Sheets[name]
    if (!ws) continue
    out[name] = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true, defval: null, blankrows: true })
  }
  return out
}
