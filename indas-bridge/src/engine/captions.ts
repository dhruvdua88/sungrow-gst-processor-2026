/**
 * Classification data: Schedule III (Division II, Ind AS) face captions and the text rules that place
 * a statement category (from the Ind AS workbook) or a local GAAP face grouping (Division I, from the
 * signed workbook) on a caption. Rules are tried in order; the first match wins, so specific rules
 * (e.g. "changes in inventories") come before general ones (e.g. "inventories").
 *
 * Only generic Schedule III wording lives here — no client names or ledger codes.
 */
import type { Caption, SectionId, StatementKind } from './types'

export const SECTION_LABEL: Record<SectionId, string> = {
  NCA: 'Non-current assets',
  CA: 'Current assets',
  EQ: 'Equity',
  NCL: 'Non-current liabilities',
  CL: 'Current liabilities',
  INC: 'Income',
  EXP: 'Expenses',
  TAX: 'Tax expense',
  UNMAPPED: 'Unmapped',
}

export const SECTION_ORDER: SectionId[] = ['NCA', 'CA', 'EQ', 'NCL', 'CL', 'INC', 'EXP', 'TAX', 'UNMAPPED']

/** Credit-natured sections are shown with the sign flipped on the face statements. */
export const CREDIT_SECTIONS: ReadonlySet<SectionId> = new Set(['EQ', 'NCL', 'CL', 'INC'])

export function statementOf(section: SectionId): StatementKind {
  return section === 'INC' || section === 'EXP' || section === 'TAX' ? 'PL' : 'BS'
}

type Rule = { test: RegExp; caption: string }

const C = (id: string, section: SectionId, label: string, order: number): Caption => ({
  id,
  section,
  label,
  order,
  statement: statementOf(section),
})

export const CAPTIONS: Caption[] = [
  C('ppe', 'NCA', 'Property, plant and equipment', 10),
  C('cwip', 'NCA', 'Capital work-in-progress', 11),
  C('invprop', 'NCA', 'Investment property', 12),
  C('goodwill', 'NCA', 'Goodwill', 13),
  C('rou', 'NCA', 'Right-of-use assets', 14),
  C('intangibles', 'NCA', 'Other intangible assets', 15),
  C('ncinvest', 'NCA', 'Investments (non-current)', 16),
  C('ncloans', 'NCA', 'Loans (non-current)', 17),
  C('ncofa', 'NCA', 'Other financial assets (non-current)', 18),
  C('dta', 'NCA', 'Deferred tax assets (net)', 19),
  C('nctax', 'NCA', 'Non-current tax assets (net)', 20),
  C('ncoa', 'NCA', 'Other non-current assets', 21),
  C('inventories', 'CA', 'Inventories', 30),
  C('cinvest', 'CA', 'Investments (current)', 31),
  C('receivables', 'CA', 'Trade receivables', 32),
  C('cash', 'CA', 'Cash and cash equivalents', 33),
  C('otherbank', 'CA', 'Bank balances other than cash and cash equivalents', 34),
  C('cloans', 'CA', 'Loans (current)', 35),
  C('cofa', 'CA', 'Other financial assets (current)', 36),
  C('ctax', 'CA', 'Current tax assets (net)', 37),
  C('coa', 'CA', 'Other current assets', 38),
  C('sharecap', 'EQ', 'Equity share capital', 50),
  C('otherequity', 'EQ', 'Other equity', 51),
  C('ncborrow', 'NCL', 'Borrowings (non-current)', 60),
  C('nclease', 'NCL', 'Lease liabilities (non-current)', 61),
  C('ncofl', 'NCL', 'Other financial liabilities (non-current)', 62),
  C('ncprov', 'NCL', 'Provisions (non-current)', 63),
  C('dtl', 'NCL', 'Deferred tax liabilities (net)', 64),
  C('ncol', 'NCL', 'Other non-current liabilities', 65),
  C('cborrow', 'CL', 'Borrowings (current)', 70),
  C('clease', 'CL', 'Lease liabilities (current)', 71),
  C('payables', 'CL', 'Trade payables', 72),
  C('cofl', 'CL', 'Other financial liabilities (current)', 73),
  C('col', 'CL', 'Other current liabilities', 74),
  C('cprov', 'CL', 'Provisions (current)', 75),
  C('ctaxl', 'CL', 'Current tax liabilities (net)', 76),
  C('revenue', 'INC', 'Revenue from operations', 80),
  C('otherincome', 'INC', 'Other income', 81),
  C('materials', 'EXP', 'Cost of materials consumed', 90),
  C('purchases', 'EXP', 'Purchases of stock-in-trade', 91),
  C('invchange', 'EXP', 'Changes in inventories of finished goods and WIP', 92),
  C('employee', 'EXP', 'Employee benefits expense', 93),
  C('finance', 'EXP', 'Finance costs', 94),
  C('depreciation', 'EXP', 'Depreciation and amortisation expense', 95),
  C('otherexp', 'EXP', 'Other expenses', 96),
  C('exceptional', 'EXP', 'Exceptional items', 97),
  C('tax', 'TAX', 'Tax expense', 100),
  // Local GAAP (Division I) captions with no direct Division II equivalent. They appear in the bridge
  // with a local GAAP balance that is reclassified away in full.
  C('div1:ltla', 'NCA', 'Long-term loans and advances (Division I)', 22),
  C('div1:stla', 'CA', 'Short-term loans and advances (Division I)', 39),
  C('unmapped:BS', 'UNMAPPED', 'Unmapped balance sheet items', 200),
  { ...C('unmapped:PL', 'UNMAPPED', 'Unmapped profit and loss items', 201), statement: 'PL' },
]

export const CAPTION_BY_ID: Map<string, Caption> = new Map(CAPTIONS.map((c) => [c.id, c]))

const NC = String.raw`non[\s-]*current|long[\s-]*term`

/** Rules for Ind AS statement categories (Division II wording). */
const IND_AS_RULES: Rule[] = [
  { test: /changes? in inventor/, caption: 'invchange' },
  { test: /cost of materials?/, caption: 'materials' },
  { test: /purchases? of stock/, caption: 'purchases' },
  { test: /employee benefit/, caption: 'employee' },
  { test: /finance cost/, caption: 'finance' },
  { test: /depreciation|amorti[sz]ation/, caption: 'depreciation' },
  { test: /exceptional/, caption: 'exceptional' },
  { test: /other expenses?/, caption: 'otherexp' },
  { test: /revenue from operations?|^revenue$|^sales/, caption: 'revenue' },
  { test: /other income/, caption: 'otherincome' },
  { test: /tax expenses?|^current tax$|income tax expense/, caption: 'tax' },
  { test: /right[\s-]*of[\s-]*use/, caption: 'rou' },
  { test: /capital work/, caption: 'cwip' },
  { test: /investment property/, caption: 'invprop' },
  { test: /goodwill/, caption: 'goodwill' },
  { test: /intangible/, caption: 'intangibles' },
  { test: /property,? plant|fixed assets/, caption: 'ppe' },
  { test: /deferred tax assets?/, caption: 'dta' },
  { test: /deferred tax/, caption: 'dtl' },
  { test: new RegExp(`(${NC}) tax assets?|tax assets?.*(${NC})`), caption: 'nctax' },
  { test: /income tax assets?|current tax assets?|advance tax/, caption: 'ctax' },
  { test: /current tax liabilit|income tax liabilit/, caption: 'ctaxl' },
  { test: new RegExp(`lease liabilit.*(${NC})|(${NC}).*lease liabilit`), caption: 'nclease' },
  { test: /lease liabilit/, caption: 'clease' },
  { test: new RegExp(`borrowings?.*(${NC})|(${NC}).*borrowings?`), caption: 'ncborrow' },
  { test: /borrowings?/, caption: 'cborrow' },
  { test: new RegExp(`other financial liabilit.*(${NC})`), caption: 'ncofl' },
  { test: /other financial liabilit/, caption: 'cofl' },
  { test: new RegExp(`other financial assets?.*(${NC})`), caption: 'ncofa' },
  { test: /other financial assets?/, caption: 'cofa' },
  { test: new RegExp(`provisions?.*(${NC})|(${NC}).*provisions?`), caption: 'ncprov' },
  { test: /provisions?/, caption: 'cprov' },
  { test: new RegExp(`investments?.*(${NC})`), caption: 'ncinvest' },
  { test: /investments?/, caption: 'cinvest' },
  { test: new RegExp(`loans?.*(${NC})`), caption: 'ncloans' },
  { test: /^loans?\b/, caption: 'cloans' },
  { test: /inventor/, caption: 'inventories' },
  { test: /trade receivable|sundry debtor/, caption: 'receivables' },
  { test: /other bank|bank balances? other/, caption: 'otherbank' },
  { test: /cash and cash/, caption: 'cash' },
  { test: /trade payable|sundry creditor/, caption: 'payables' },
  { test: /share capital/, caption: 'sharecap' },
  { test: /other equity|reserves? and surplus|retained earnings/, caption: 'otherequity' },
  { test: /other non[\s-]*current assets?/, caption: 'ncoa' },
  { test: /other current assets?/, caption: 'coa' },
  { test: /other non[\s-]*current liabilit/, caption: 'ncol' },
  { test: /other current liabilit/, caption: 'col' },
]

/** Extra rules for local GAAP face groupings (Division I wording), tried before the Ind AS rules. */
const DIV1_RULES: Rule[] = [
  { test: /long[\s-]*term loans? and advances/, caption: 'div1:ltla' },
  { test: /short[\s-]*term loans? and advances/, caption: 'div1:stla' },
  { test: /long[\s-]*term provisions?/, caption: 'ncprov' },
  { test: /short[\s-]*term provisions?/, caption: 'cprov' },
  { test: /long[\s-]*term borrowings?/, caption: 'ncborrow' },
  { test: /short[\s-]*term borrowings?/, caption: 'cborrow' },
  { test: /change in inventor/, caption: 'invchange' },
  { test: /cost of material/, caption: 'materials' },
]

export function normaliseText(s: string): string {
  return s.toLowerCase().replace(/[‐-―]/g, '-').replace(/\s+/g, ' ').trim()
}

function apply(rules: Rule[], text: string): string | null {
  const t = normaliseText(text)
  if (!t) return null
  for (const r of rules) if (r.test.test(t)) return r.caption
  return null
}

/** Caption for an Ind AS statement category; null when no rule matches. */
export function classifyIndAsCategory(category: string): string | null {
  return apply(IND_AS_RULES, category)
}

/** Caption for a local GAAP (Division I) face grouping; null when no rule matches. */
export function classifyLocalFace(face: string): string | null {
  return apply(DIV1_RULES, face) ?? apply(IND_AS_RULES, face)
}

/** Unmapped bucket for a statement kind. */
export function unmappedCaption(kind: StatementKind): string {
  return kind === 'BS' ? 'unmapped:BS' : 'unmapped:PL'
}

/** Guess whether a local face grouping text belongs to the P&L (used only for unmapped items). */
export function looksLikePL(text: string): boolean {
  return /expense|income|revenue|sales|purchase|cost|depreciation|tax expense|profit|loss/i.test(text)
}

/** Ind AS standard tag (e.g. "Ind AS 116") found in a narration, or "Other". */
export function standardTag(text: string): string {
  const m = /Ind\s*AS\s*(\d{1,3})/i.exec(text)
  return m ? `Ind AS ${m[1]}` : 'Other'
}
