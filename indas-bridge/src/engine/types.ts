import type { Paise } from './money'

/** A period is keyed by the calendar year in which the financial year ends (FY 2025-26 -> 2026). */
export type PeriodKey = number

export interface Period {
  key: PeriodKey
  /** e.g. "FY 2025-26" or "1-Apr-2024 (transition)" */
  label: string
  /** Balance sheet date, DD-MM-YYYY */
  bsDate: string
  /** False for the Ind AS transition date: there is no Ind AS profit and loss for that column. */
  hasPL: boolean
}

/** A sheet as a dense 2-D array of cell values (string | number | boolean | null). */
export type Grid = unknown[][]

// ---------------------------------------------------------------- signed local GAAP (xlsb T_B)

export interface SignedLedger {
  /** Normalised join key (full "code/name"). */
  key: string
  code: string
  name: string
  company: string
  primaryGroup: string
  parentGroup: string
  finalCode: string
  /** Face grouping per period (falls back to current-year face when blank). */
  face: Record<PeriodKey, string>
  note: Record<PeriodKey, string>
  subNote: string
  amounts: Record<PeriodKey, Paise>
  /** Number of T_B rows aggregated into this ledger. */
  rowCount: number
}

export interface SignedTB {
  ledgers: SignedLedger[]
  /** Periods for which the T_B carries an amount column. */
  periods: PeriodKey[]
  rowCount: number
  warnings: string[]
}

// ---------------------------------------------------------------- Ind AS workbook (Copy TB)

export type Layer = 'G' | 'A' | 'P'

export interface CopyRow {
  /** 1-based row number on the TB sheet (used by the voucher register's "TB row"). */
  sheetRow: number
  ref: string
  code: string
  ledger: string
  key: string
  category: string
  subHeading: string
  lineItem: string
  layer: Layer
  local: Record<PeriodKey, Paise>
  jv: Record<PeriodKey, Paise>
  indas: Record<PeriodKey, Paise>
}

export interface StatementFigure {
  label: string
  values: Record<PeriodKey, Paise>
}

export interface VoucherEntry {
  ref: string
  tbRow: number | null
  voucher: string
  narration: string
  statementLine: string
  layer: string
  amounts: Record<PeriodKey, Paise>
  reference: string
  why: string
  section: string
}

export interface CopyModel {
  rows: CopyRow[]
  periods: Period[]
  statements: { bs: StatementFigure[]; pl: StatementFigure[] }
  entries: VoucherEntry[]
  warnings: string[]
}

// ---------------------------------------------------------------- classification

export type StatementKind = 'BS' | 'PL'
export type SectionId = 'NCA' | 'CA' | 'EQ' | 'NCL' | 'CL' | 'INC' | 'EXP' | 'TAX' | 'UNMAPPED'

export interface Caption {
  id: string
  statement: StatementKind
  section: SectionId
  label: string
  order: number
}
