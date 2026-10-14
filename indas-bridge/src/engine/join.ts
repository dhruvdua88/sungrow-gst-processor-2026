/**
 * Join the signed local GAAP ledgers (xlsb) to the Ind AS mapping rows (Copy TB, layer G).
 *
 * Rules (see docs in README "Engine rules"):
 *  1. Join on the full normalised "code/name" string. Code-only fallback is used only when the code is
 *     unique on both sides and the full string found no partner.
 *  2. xlsb wins: for every period the xlsb carries, the xlsb amount is the start of the bridge.
 *     Copy supplies only the mapping (category / line item) and the Ind AS JV layer.
 *  3. When Copy maps one ledger to several rows, the xlsb amount is allocated:
 *     a. Copy rows add up to the xlsb amount  -> keep Copy's split;
 *     b. one Copy row equals the xlsb amount   -> all of it goes to that row (a regroup);
 *     c. otherwise                             -> other rows keep Copy's amount, the largest row takes
 *        the balance (a regroup).
 *  4. xlsb-only ledgers take an assumed Ind AS caption from their local face grouping.
 */
import { classifyIndAsCategory, classifyLocalFace, CAPTION_BY_ID, looksLikePL, unmappedCaption } from './captions'
import type { Paise } from './money'
import type { CopyModel, CopyRow, Period, PeriodKey, SignedLedger, SignedTB } from './types'

/** Differences below this many paise (₹1) are treated as rounding, not variances. */
export const VARIANCE_TOLERANCE: Paise = 100

export type MatchStatus = 'matched' | 'xlsb-only' | 'copy-only'

export interface Allocation {
  /** Index into JoinedLedger.copyRows, or -1 for an xlsb-only ledger with an assumed mapping. */
  row: number
  amount: Paise
  captionId: string
  category: string
  lineItem: string
}

export interface JoinedLedger {
  key: string
  code: string
  name: string
  status: MatchStatus
  matchedBy: 'key' | 'code' | null
  signed: SignedLedger | null
  copyRows: CopyRow[]
  primaryGroup: string
  parentGroup: string
  face: Record<PeriodKey, string>
  note: Record<PeriodKey, string>
  /** Amount from the signed books (undefined when the xlsb has no column for the period). */
  xlsb: Record<PeriodKey, Paise | undefined>
  copyLocal: Record<PeriodKey, Paise>
  /** The amount the bridge starts from: xlsb when available, else Copy local GAAP. */
  used: Record<PeriodKey, Paise>
  source: Record<PeriodKey, 'xlsb' | 'copy'>
  /** xlsb − Copy local GAAP. */
  diff: Record<PeriodKey, Paise>
  alloc: Record<PeriodKey, Allocation[]>
  /** True when the allocation departs from Copy's own split (ledger regrouped). */
  regroup: Record<PeriodKey, boolean>
  /** Local GAAP caption from the signed face grouping. */
  localCaption: Record<PeriodKey, string>
  /** True when the Ind AS mapping was assumed (xlsb-only ledger). */
  assumedMapping: boolean
}

export interface Adjustment {
  row: CopyRow
  captionId: string
  standard: string
}

export interface JoinResult {
  periods: Period[]
  ledgers: JoinedLedger[]
  adjustments: Adjustment[]
  stats: { matched: number; matchedByCode: number; xlsbOnly: number; copyOnly: number; signedLedgers: number; copyLedgers: number }
  warnings: string[]
}

export function captionForCategory(category: string, fallbackText = ''): string {
  return classifyIndAsCategory(category) ?? unmappedCaption(looksLikePL(category || fallbackText) ? 'PL' : 'BS')
}

export function localCaptionForFace(face: string, fallbackCaption: string | null): string {
  const c = classifyLocalFace(face)
  if (c) return c
  if (fallbackCaption) return fallbackCaption
  return unmappedCaption(looksLikePL(face) ? 'PL' : 'BS')
}

function emptyRecord<T>(periods: Period[], value: T): Record<PeriodKey, T> {
  const r: Record<PeriodKey, T> = {}
  for (const p of periods) r[p.key] = value
  return r
}

export function allocate(rows: CopyRow[], period: PeriodKey, amount: Paise): { alloc: { row: number; amount: Paise }[]; regroup: boolean } {
  if (rows.length === 1) return { alloc: [{ row: 0, amount }], regroup: false }
  const locals = rows.map((r) => r.local[period] ?? 0)
  const total = locals.reduce((a, b) => a + b, 0)
  const largest = locals.reduce((best, v, i) => (Math.abs(v) > Math.abs(locals[best]) ? i : best), 0)
  if (Math.abs(total - amount) <= VARIANCE_TOLERANCE) {
    const alloc = locals.map((v, i) => ({ row: i, amount: v }))
    alloc[largest].amount += amount - total
    return { alloc, regroup: false }
  }
  const exact = locals.findIndex((v) => Math.abs(v - amount) <= VARIANCE_TOLERANCE)
  if (exact >= 0) {
    return { alloc: locals.map((_, i) => ({ row: i, amount: i === exact ? amount : 0 })), regroup: true }
  }
  const others = total - locals[largest]
  return { alloc: locals.map((v, i) => ({ row: i, amount: i === largest ? amount - others : v })), regroup: true }
}

export function joinBooks(signed: SignedTB, copy: CopyModel): JoinResult {
  const warnings: string[] = []
  const periods = mergePeriods(signed, copy)
  const gRows = copy.rows.filter((r) => r.layer === 'G')
  const copyByKey = new Map<string, CopyRow[]>()
  for (const r of gRows) {
    const list = copyByKey.get(r.key)
    if (list) list.push(r)
    else copyByKey.set(r.key, [r])
  }
  const signedByKey = new Map(signed.ledgers.map((l) => [l.key, l]))

  // Code-only fallback: codes unique on both sides.
  const codeCount = (codes: string[]) => {
    const m = new Map<string, number>()
    for (const c of codes) if (c) m.set(c, (m.get(c) ?? 0) + 1)
    return m
  }
  const signedCodes = codeCount(signed.ledgers.map((l) => l.code))
  const copyKeyCodes = codeCount([...copyByKey.values()].map((rows) => codeOfRows(rows)))
  const copyKeyByCode = new Map<string, string>()
  for (const [key, rows] of copyByKey) {
    const c = codeOfRows(rows)
    if (c && copyKeyCodes.get(c) === 1) copyKeyByCode.set(c, key)
  }

  const usedCopyKeys = new Set<string>()
  const pairs: { signed: SignedLedger | null; copyKey: string | null; by: 'key' | 'code' | null }[] = []
  const unmatchedSigned: SignedLedger[] = []
  for (const l of signed.ledgers) {
    if (copyByKey.has(l.key)) {
      pairs.push({ signed: l, copyKey: l.key, by: 'key' })
      usedCopyKeys.add(l.key)
    } else unmatchedSigned.push(l)
  }
  for (const l of unmatchedSigned) {
    const ck = l.code && signedCodes.get(l.code) === 1 ? copyKeyByCode.get(l.code) : undefined
    if (ck && !usedCopyKeys.has(ck) && !signedByKey.has(ck)) {
      pairs.push({ signed: l, copyKey: ck, by: 'code' })
      usedCopyKeys.add(ck)
    } else pairs.push({ signed: l, copyKey: null, by: null })
  }
  for (const key of copyByKey.keys()) if (!usedCopyKeys.has(key)) pairs.push({ signed: null, copyKey: key, by: null })

  const ledgers: JoinedLedger[] = pairs.map(({ signed: s, copyKey, by }) => buildLedger(s, copyKey ? copyByKey.get(copyKey)! : [], by, periods, signed.periods))

  const adjustments: Adjustment[] = copy.rows
    .filter((r) => r.layer !== 'G')
    .map((row) => ({ row, captionId: captionForCategory(row.category, row.lineItem), standard: standardOf(row) }))

  for (const a of adjustments) {
    const cap = CAPTION_BY_ID.get(a.captionId)
    if (cap && ((a.row.layer === 'P' && cap.statement !== 'PL') || (a.row.layer === 'A' && cap.statement !== 'BS'))) {
      warnings.push(`Adjustment on TB row ${a.row.sheetRow} is layer ${a.row.layer} but its category "${a.row.category}" reads as ${cap.statement === 'BS' ? 'balance sheet' : 'profit and loss'}.`)
    }
  }
  const unmappedCats = new Set(gRows.filter((r) => !classifyIndAsCategory(r.category)).map((r) => r.category || '(blank)'))
  for (const c of unmappedCats) warnings.push(`Statement category "${c}" is not recognised; its ledgers sit in the Unmapped bucket.`)

  const stats = {
    matched: ledgers.filter((l) => l.status === 'matched').length,
    matchedByCode: ledgers.filter((l) => l.matchedBy === 'code').length,
    xlsbOnly: ledgers.filter((l) => l.status === 'xlsb-only').length,
    copyOnly: ledgers.filter((l) => l.status === 'copy-only').length,
    signedLedgers: signed.ledgers.length,
    copyLedgers: copyByKey.size,
  }
  return { periods, ledgers, adjustments, stats, warnings }
}

function codeOfRows(rows: CopyRow[]): string {
  const fromLedger = /^\s*([A-Za-z0-9-]+)\s*\//.exec(rows[0].ledger)
  return fromLedger ? fromLedger[1].toUpperCase() : ''
}

function standardOf(row: CopyRow): string {
  const m = /Ind\s*AS\s*(\d{1,3})/i.exec(`${row.ledger} ${row.lineItem}`)
  return m ? `Ind AS ${m[1]}` : 'Other'
}

function mergePeriods(signed: SignedTB, copy: CopyModel): Period[] {
  if (copy.periods.length) return copy.periods
  return signed.periods.map((k) => ({ key: k, label: `FY ${k - 1}-${String(k % 100).padStart(2, '0')}`, bsDate: `31-03-${k}`, hasPL: true }))
}

function buildLedger(
  s: SignedLedger | null,
  rows: CopyRow[],
  by: 'key' | 'code' | null,
  periods: Period[],
  signedPeriods: PeriodKey[],
): JoinedLedger {
  const status: MatchStatus = s && rows.length ? 'matched' : s ? 'xlsb-only' : 'copy-only'
  const first = rows[0]
  const led: JoinedLedger = {
    key: s?.key ?? first.key,
    code: s?.code ?? codeOfRows(rows),
    name: s?.name ?? first.ledger,
    status,
    matchedBy: by,
    signed: s,
    copyRows: rows,
    primaryGroup: s?.primaryGroup ?? '',
    parentGroup: s?.parentGroup ?? '',
    face: s ? { ...s.face } : {},
    note: s ? { ...s.note } : {},
    xlsb: {},
    copyLocal: emptyRecord(periods, 0),
    used: emptyRecord(periods, 0),
    source: emptyRecord(periods, 'copy' as const),
    diff: emptyRecord(periods, 0),
    alloc: {},
    regroup: emptyRecord(periods, false),
    localCaption: {},
    assumedMapping: rows.length === 0,
  }
  for (const p of periods) {
    const k = p.key
    const copyLocal = rows.reduce((a, r) => a + (r.local[k] ?? 0), 0)
    const hasX = signedPeriods.includes(k)
    const x = hasX ? (s?.amounts[k] ?? 0) : undefined
    led.copyLocal[k] = copyLocal
    led.xlsb[k] = x
    led.used[k] = x ?? copyLocal
    led.source[k] = x === undefined ? 'copy' : 'xlsb'
    led.diff[k] = x === undefined ? 0 : x - copyLocal

    // Ind AS mapping of this ledger's amount.
    if (rows.length) {
      const { alloc, regroup } = allocate(rows, k, led.used[k])
      led.regroup[k] = regroup
      led.alloc[k] = alloc.map((a) => ({
        row: a.row,
        amount: a.amount,
        captionId: captionForCategory(rows[a.row].category, rows[a.row].lineItem),
        category: rows[a.row].category,
        lineItem: rows[a.row].lineItem,
      }))
    } else {
      const face = s?.face[k] ?? ''
      const cap = classifyLocalFace(face)
      const captionId = cap && !cap.startsWith('div1:') ? cap : unmappedCaption(looksLikePL(face) ? 'PL' : 'BS')
      led.alloc[k] = [{ row: -1, amount: led.used[k], captionId, category: face || '(no mapping)', lineItem: 'Not in Ind AS workbook' }]
    }
    // Local caption: signed face grouping; Copy's mapping when the xlsb is silent.
    const main = led.alloc[k].reduce((best, a) => (Math.abs(a.amount) > Math.abs(best.amount) ? a : best), led.alloc[k][0])
    led.localCaption[k] = s ? localCaptionForFace(s.face[k] ?? '', main.captionId) : main.captionId
    if (!led.face[k]) led.face[k] = s ? '' : `(Ind AS: ${first.category})`
    if (!led.note[k]) led.note[k] = ''
  }
  return led
}
