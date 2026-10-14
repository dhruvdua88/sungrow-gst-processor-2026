/**
 * Parser for the signed local GAAP trial balance (sheet "T_B" of the Schedule 3 Automation Tool
 * workbook). Columns are located by header text, never by fixed position.
 */
import { cellText, findCol, findRow, headerDate, ledgerCode, ledgerKey, norm, periodForDate } from './grid'
import { rupeesToPaise } from './money'
import type { Grid, PeriodKey, SignedLedger, SignedTB } from './types'

export const SIGNED_SHEET_CANDIDATES = ['T_B', 'TB', 'Trial Balance']

export function parseSignedTB(grid: Grid): SignedTB {
  const warnings: string[] = []
  const h = findRow(
    grid,
    (cells) => cells.some((c) => c === 'name of ledger' || c === 'ledger name') && cells.some((c) => /^system primary group|^face grouping.*current year/.test(c)),
    40,
  )
  if (h < 0) throw new Error('Could not find the "Name of Ledger" header in the signed trial balance sheet.')
  const raw = grid[h] ?? []
  const cells = raw.map(norm)

  const nameCol = findCol(cells, /^(name of ledger|ledger name)$/)
  const companyCol = findCol(cells, /^name of company/)
  const primaryCol = findCol(cells, /system primary group/)
  const parentCol = findCol(cells, /system parent group/)
  const finalCol = findCol(cells, /^final codes?$/)

  // Amount columns: header cells holding a date (text or Excel serial).
  const amountCols: { col: number; period: PeriodKey }[] = []
  raw.forEach((v, col) => {
    if (col <= nameCol) return
    const d = headerDate(typeof v === 'string' ? v.trim() : v)
    if (d && d.m === 3 && d.d === 31) amountCols.push({ col, period: periodForDate(d) })
  })
  if (amountCols.length === 0) throw new Error('No dated amount columns (e.g. "31-March-2026") found next to "Name of Ledger".')
  amountCols.sort((a, b) => b.period - a.period)
  const seen = new Set<number>()
  const amountByPeriod = amountCols.filter((a) => (seen.has(a.period) ? false : (seen.add(a.period), true)))

  const faceCY = findCol(cells, /^face grouping.*current year/)
  const facePY = findCol(cells, /^face grouping.*(?<!preceding )previous year/)
  const facePPY = findCol(cells, /^face grouping.*preceding/)
  const noteCY = findCol(cells, /^note grouping.*current year/)
  const notePY = findCol(cells, /^note grouping.*previous year/)
  const subNoteCY = findCol(cells, /^sub-?note grouping.*current year/)
  if (faceCY < 0) warnings.push('Face grouping (current year) column not found; local captions will be unmapped.')

  const faceCols = [faceCY, facePY, facePPY]
  const noteCols = [noteCY, notePY, -1]

  const byKey = new Map<string, SignedLedger>()
  let rowCount = 0
  for (let r = h + 1; r < grid.length; r++) {
    const row = grid[r] ?? []
    const nameRaw = row[nameCol]
    if (typeof nameRaw !== 'string') continue
    const name = cellText(nameRaw)
    if (!name) continue
    rowCount++
    const key = ledgerKey(name)
    let led = byKey.get(key)
    if (!led) {
      led = {
        key,
        code: ledgerCode(name),
        name,
        company: companyCol >= 0 ? cellText(row[companyCol]) : '',
        primaryGroup: primaryCol >= 0 ? cellText(row[primaryCol]) : '',
        parentGroup: parentCol >= 0 ? cellText(row[parentCol]) : '',
        finalCode: finalCol >= 0 ? cellText(row[finalCol]) : '',
        face: {},
        note: {},
        subNote: subNoteCY >= 0 ? cellText(row[subNoteCY]) : '',
        amounts: {},
        rowCount: 0,
      }
      byKey.set(key, led)
    }
    led.rowCount++
    amountByPeriod.forEach(({ col, period }, i) => {
      led!.amounts[period] = (led!.amounts[period] ?? 0) + rupeesToPaise(row[col])
      const fc = faceCols[i] ?? -1
      const nc = noteCols[i] ?? -1
      const face = fc >= 0 ? cellText(row[fc]) : ''
      const note = nc >= 0 ? cellText(row[nc]) : ''
      if (face && !led!.face[period]) led!.face[period] = face
      if (note && !led!.note[period]) led!.note[period] = note
    })
  }
  // Fall back to the current-year face / note when a period's grouping is blank.
  const cy = amountByPeriod[0].period
  for (const led of byKey.values()) {
    for (const { period } of amountByPeriod) {
      if (!led.face[period]) led.face[period] = led.face[cy] ?? ''
      if (!led.note[period]) led.note[period] = led.note[cy] ?? ''
    }
  }
  return { ledgers: [...byKey.values()], periods: amountByPeriod.map((a) => a.period), rowCount, warnings }
}
