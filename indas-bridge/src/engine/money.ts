/**
 * Money helpers. Every amount inside the engine is an integer number of paise.
 * Workbook values (rupees, possibly fractional floats) are converted once, at the parser boundary.
 * Sign convention: debit positive, credit negative.
 */

export type Paise = number

export type Unit = 'rupees' | 'lakhs' | 'crores'

/** Convert a rupee value read from a workbook cell into integer paise. Non-numbers become 0. */
export function rupeesToPaise(value: unknown): Paise {
  const n = toNumber(value)
  if (n === null) return 0
  return Math.round(n * 100)
}

/** Convert a value expressed in ₹ lakhs (as on the face statements) into paise. */
export function lakhsToPaise(value: unknown): Paise {
  const n = toNumber(value)
  if (n === null) return 0
  return Math.round(n * 1e7)
}

export function toNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (typeof value === 'string') {
    const cleaned = value.replace(/[₹,\s]/g, '').replace(/^\((.*)\)$/, '-$1')
    if (cleaned === '' || cleaned === '-') return null
    const n = Number(cleaned)
    return Number.isFinite(n) ? n : null
  }
  return null
}

/** Group an integer string the Indian way: 1855455 -> 18,55,455 */
function groupIndian(intDigits: string): string {
  if (intDigits.length <= 3) return intDigits
  const last3 = intDigits.slice(-3)
  let rest = intDigits.slice(0, -3)
  const parts: string[] = []
  while (rest.length > 2) {
    parts.unshift(rest.slice(-2))
    rest = rest.slice(0, -2)
  }
  if (rest) parts.unshift(rest)
  return `${parts.join(',')},${last3}`
}

const UNIT_PAISE: Record<Unit, number> = { rupees: 100, lakhs: 1e7, crores: 1e9 }
const UNIT_SUFFIX: Record<Unit, string> = { rupees: '', lakhs: ' L', crores: ' Cr' }

/**
 * Format paise for display. Rupees show 2 decimals (₹18,55,455.91); lakhs and crores show 2 decimals
 * of the unit. Negative values use a leading minus. Integer arithmetic only for the rupee case.
 */
export function formatMoney(paise: Paise, unit: Unit = 'rupees', opts: { symbol?: boolean; signed?: boolean } = {}): string {
  const symbol = opts.symbol ?? true
  const neg = paise < 0
  const abs = Math.abs(Math.round(paise))
  let body: string
  if (unit === 'rupees') {
    const rupees = Math.floor(abs / 100)
    const p = abs % 100
    body = `${groupIndian(String(rupees))}.${String(p).padStart(2, '0')}`
  } else {
    // Round half away from zero to 2 decimals of the unit using integer maths.
    const per = UNIT_PAISE[unit] / 100 // paise per 0.01 unit
    const hundredths = Math.round(abs / per)
    const whole = Math.floor(hundredths / 100)
    const frac = hundredths % 100
    body = `${groupIndian(String(whole))}.${String(frac).padStart(2, '0')}`
  }
  const sign = neg ? '−' : opts.signed && abs > 0 ? '+' : ''
  return `${sign}${symbol ? '₹' : ''}${body}${UNIT_SUFFIX[unit]}`
}

/** Plain numeric string in the chosen unit for CSV export (no grouping, dot decimal). */
export function toUnitString(paise: Paise, unit: Unit = 'rupees'): string {
  const value = paise / UNIT_PAISE[unit]
  return value.toFixed(2)
}

export function sum(values: Iterable<Paise>): Paise {
  let total = 0
  for (const v of values) total += v
  return total
}
