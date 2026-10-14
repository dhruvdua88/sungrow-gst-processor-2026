import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'
import {
  allocate,
  analyse,
  formatMoney,
  joinBooks,
  lakhsToPaise,
  parseCopy,
  parseSignedTB,
  rupeesToPaise,
  type Analysis,
  type CopyRow,
  type Step,
} from '../src/engine'
import { classifyIndAsCategory, classifyLocalFace } from '../src/engine/captions'
import { headerDate, ledgerKey, periodFromLabel } from '../src/engine/grid'
import { demoBalanceSheetGrid, demoCopyTBGrid, demoEntriesGrid, demoProfitLossGrid, demoSignedGrid } from '../src/demo/demo'
import { loadCopy, loadSigned } from '../src/worker/load'

function demo(): Analysis {
  const signed = parseSignedTB(demoSignedGrid())
  const copy = parseCopy({ tb: demoCopyTBGrid(), bs: demoBalanceSheetGrid(), pl: demoProfitLossGrid(), entries: demoEntriesGrid() })
  return analyse(signed, copy)
}

const stepsTotal = (steps: Step[], uptoId: string) => {
  let total = 0
  for (const s of steps) {
    if (s.id === uptoId) return total
    if (s.kind !== 'total') total += s.amount
  }
  throw new Error(`no step ${uptoId}`)
}

describe('money', () => {
  it('formats paise in Indian grouping', () => {
    expect(formatMoney(185545591)).toBe('₹18,55,455.91')
    expect(formatMoney(-185545591)).toBe('−₹18,55,455.91')
    expect(formatMoney(5)).toBe('₹0.05')
    expect(formatMoney(12345678900, 'lakhs')).toBe('₹1,234.57 L')
    expect(formatMoney(12345678900, 'crores')).toBe('₹12.35 Cr')
  })
  it('converts workbook values to integer paise', () => {
    expect(rupeesToPaise(12345678.7732548)).toBe(1234567877)
    expect(rupeesToPaise('1,23,456.78')).toBe(12345678)
    expect(rupeesToPaise(null)).toBe(0)
    expect(lakhsToPaise(1.5)).toBe(15000000)
  })
})

describe('header helpers', () => {
  it('reads period labels and dates', () => {
    expect(periodFromLabel('FY 2025-26')).toBe(2026)
    expect(periodFromLabel('31 Mar 2025')).toBe(2025)
    expect(periodFromLabel('1 Apr 2024')).toBe(2024)
    expect(periodFromLabel(45382)).toBe(2024)
    expect(headerDate('31-March-2026')).toEqual({ y: 2026, m: 3, d: 31 })
  })
  it('normalises ledger keys', () => {
    expect(ledgerKey(' D9001 / Demo  Ledger ')).toBe('D9001/DEMO LEDGER')
  })
  it('classifies Division I and Division II captions', () => {
    expect(classifyIndAsCategory('Changes in inventories of finished goods, Stock in Trade and work in progress')).toBe('invchange')
    expect(classifyIndAsCategory('Inventories')).toBe('inventories')
    expect(classifyIndAsCategory('Other financial assets - non current')).toBe('ncofa')
    expect(classifyIndAsCategory('Other financial assets -  current')).toBe('cofa')
    expect(classifyIndAsCategory('Bank balances other than cash and cash equivalents')).toBe('otherbank')
    expect(classifyIndAsCategory('Deferred tax liabilities, net')).toBe('dtl')
    expect(classifyIndAsCategory('Income Tax Assets')).toBe('ctax')
    expect(classifyIndAsCategory('Non-current tax assets (net)')).toBe('nctax')
    expect(classifyLocalFace('Short term loans and advances')).toBe('div1:stla')
    expect(classifyLocalFace('Long term provisions')).toBe('ncprov')
    expect(classifyLocalFace('Trade payables due to others')).toBe('payables')
  })
})

describe('parsers', () => {
  it('parses the signed T_B by header text and aggregates duplicate ledgers', () => {
    const grid = demoSignedGrid()
    grid.push([...grid[grid.length - 3]]) // duplicate a ledger row
    const tb = parseSignedTB(grid)
    expect(tb.periods).toEqual([2026, 2025, 2024])
    expect(tb.rowCount).toBe(grid.length - 3)
    expect(tb.ledgers.length).toBe(grid.length - 4)
    const dup = tb.ledgers.find((l) => l.rowCount === 2)!
    expect(dup).toBeDefined()
  })
  it('parses the Ind AS TB, skipping memo lines without an account code', () => {
    const copy = parseCopy({ tb: demoCopyTBGrid() })
    expect(copy.periods.map((p) => p.key)).toEqual([2026, 2025, 2024])
    expect(copy.periods[2].hasPL).toBe(false)
    expect(copy.rows.some((r) => r.ledger.startsWith('Memo'))).toBe(false)
    expect(copy.rows.filter((r) => r.layer === 'A').length).toBe(11)
    expect(copy.rows.filter((r) => r.layer === 'P').length).toBe(8)
  })
  it('parses face statements in lakhs and the voucher register', () => {
    const copy = parseCopy({ tb: demoCopyTBGrid(), bs: demoBalanceSheetGrid(), pl: demoProfitLossGrid(), entries: demoEntriesGrid() })
    expect(copy.statements.bs.find((f) => f.label === 'Total assets')?.values[2024]).toBeDefined()
    expect(copy.statements.pl.find((f) => /^Profit for the year/.test(f.label))?.values[2026]).toBeTypeOf('number')
    expect(copy.entries.length).toBe(19)
    expect(copy.entries[0].tbRow).toBeGreaterThan(0)
  })
})

describe('join and conflicts (xlsb wins)', () => {
  it('matches ledgers and reports one-sided ledgers', () => {
    const a = demo()
    expect(a.join.stats.matched).toBe(a.join.stats.signedLedgers - 1)
    expect(a.join.stats.xlsbOnly).toBe(1)
    expect(a.join.stats.copyOnly).toBe(1)
  })
  it('flags variances where Copy local GAAP differs from the signed books, and uses the xlsb amount', () => {
    const a = demo()
    const v3 = a.join.ledgers.find((l) => l.name.includes('DEMO VENDOR 3'))!
    expect(v3.diff[2026]).toBe(1000000) // xlsb −4,00,000 vs Copy −4,10,000 → +₹10,000
    expect(v3.used[2026]).toBe(-40000000)
    expect(v3.diff[2025]).toBe(0)
  })
  it('allocates a ledger split over two Copy rows to the row that equals the xlsb amount (regroup)', () => {
    const a = demo()
    const split = a.join.ledgers.find((l) => l.name.includes('PROVISION FOR GRATUITY'))!
    expect(split.regroup[2026]).toBe(true)
    expect(split.regroup[2025]).toBe(false)
    const alloc = split.alloc[2026]
    expect(alloc.find((x) => x.captionId === 'cprov')!.amount).toBe(-5000000)
    expect(alloc.find((x) => x.captionId === 'ncprov')!.amount).toBe(0)
  })
  it('allocate(): keeps the split when totals agree, else largest row takes the balance', () => {
    const row = (local: number): CopyRow => ({ sheetRow: 1, ref: '', code: '', ledger: '', key: '', category: '', subHeading: '', lineItem: '', layer: 'G', local: { 1: local }, jv: {}, indas: {} })
    expect(allocate([row(-300), row(-50)], 1, -350)).toEqual({ alloc: [{ row: 0, amount: -300 }, { row: 1, amount: -50 }], regroup: false })
    expect(allocate([row(-30000), row(-5000)], 1, -5000).regroup).toBe(true)
    expect(allocate([row(-30000), row(-5000)], 1, -9000)).toEqual({ alloc: [{ row: 0, amount: -4000 }, { row: 1, amount: -5000 }], regroup: true })
  })
  it('uses the code-only fallback only when the code is unique on both sides', () => {
    const signed = parseSignedTB(demoSignedGrid())
    const copy = parseCopy({ tb: demoCopyTBGrid() })
    // Rename one Copy ledger keeping its unique code.
    const target = copy.rows.find((r) => r.code === 'D1404')!
    target.ledger = 'D1404/DEMO PREPAIDS RENAMED'
    target.key = ledgerKey(target.ledger)
    const j = joinBooks(signed, copy)
    const led = j.ledgers.find((l) => l.code === 'D1404')!
    expect(led.status).toBe('matched')
    expect(led.matchedBy).toBe('code')
  })
})

describe('bridges reconcile', () => {
  const a = demo()
  for (const p of a.join.periods) {
    const b = a.bridges[p.key]
    it(`${p.label}: local + reclass + adjustment = Ind AS on every line`, () => {
      for (const l of [...b.bsLines, ...b.plLines]) expect(l.local + l.reclass + l.adj).toBe(l.indas)
    })
    it(`${p.label}: line totals equal statement totals`, () => {
      const sum = (lines: typeof b.bsLines, sections: string[], key: 'local' | 'indas') =>
        lines.filter((l) => sections.includes(l.caption.section)).reduce((x, l) => x + l[key], 0)
      expect(sum(b.bsLines, ['NCA', 'CA'], 'indas')).toBe(b.totals.indas.assets)
      expect(sum(b.bsLines, ['EQ'], 'indas')).toBe(b.totals.indas.equity)
      expect(sum(b.bsLines, ['NCA', 'CA'], 'local')).toBe(b.totals.local.assets)
    })
    it(`${p.label}: every tie-out check ties`, () => {
      for (const t of b.ties) expect([t.id, t.status]).toEqual([t.id, 'tie'])
    })
    it(`${p.label}: equity waterfall ends at Ind AS equity`, () => {
      const steps = b.equityWaterfall
      expect(stepsTotal(steps, 'indas-eq')).toBe(b.totals.indas.equity)
      expect(steps[0].amount).toBe(b.totals.local.equity)
    })
    if (p.hasPL) {
      it(`${p.label}: profit waterfall ends at Ind AS PAT and TCI`, () => {
        const steps = b.plWaterfall!
        expect(steps[0].amount).toBe(b.totals.local.pat)
        expect(stepsTotal(steps, 'indas-pat')).toBe(b.totals.indas.pat)
        expect(stepsTotal(steps, 'tci')).toBe(b.totals.indas.tci)
        const pSum = a.join.adjustments.filter((x) => x.row.layer === 'P').reduce((s, x) => s + (x.row.jv[p.key] ?? 0), 0)
        expect(b.totals.indas.pat).toBe(b.totals.local.pat - pSum)
      })
    } else {
      it(`${p.label}: no profit bridge at the transition date`, () => expect(b.plWaterfall).toBeNull())
    }
  }
  it('local GAAP BS reclassifies Division I captions into Ind AS captions', () => {
    const b = a.bridges[2026]
    const ob = b.bsLines.find((l) => l.caption.id === 'otherbank')!
    expect(ob.local).toBe(0)
    expect(ob.reclass).toBe(200000000)
    const stla = b.bsLines.find((l) => l.caption.id === 'div1:stla')!
    expect(stla.indas).toBe(0)
  })
})

describe('SheetJS round trip (xlsb and xlsx)', () => {
  const toBook = (sheets: Record<string, unknown[][]>, bookType: 'xlsb' | 'xlsx') => {
    const wb = XLSX.utils.book_new()
    for (const [name, grid] of Object.entries(sheets)) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(grid), name)
    const out = XLSX.write(wb, { bookType, type: 'array' }) as ArrayBuffer
    return out
  }
  it('gives the same analysis as the grid path', () => {
    const signedBuf = toBook({ Readme: [['synthetic']], T_B: demoSignedGrid() }, 'xlsb')
    const copyBuf = toBook({ 'Balance Sheet': demoBalanceSheetGrid(), 'Statement of P&L': demoProfitLossGrid(), TB: demoCopyTBGrid(), 'Ind AS Entries': demoEntriesGrid() }, 'xlsx')
    const viaFiles = analyse(loadSigned(signedBuf), loadCopy(copyBuf))
    const viaGrids = demo()
    expect(viaFiles.join.stats).toEqual(viaGrids.join.stats)
    for (const p of viaGrids.join.periods) expect(viaFiles.bridges[p.key].totals).toEqual(viaGrids.bridges[p.key].totals)
  })
})

describe('flow graph', () => {
  it('links only move left to right and the focus mode lists ledgers', async () => {
    const { buildFlow } = await import('../src/engine')
    const a = demo()
    const g = buildFlow(a.join, 2026, { statement: 'ALL' })
    const col = new Map(g.nodes.map((n) => [n.id, n.column]))
    for (const l of g.links) expect(col.get(l.source)!).toBeLessThan(col.get(l.target)!)
    expect(g.nodes.some((n) => n.kind === 'jv')).toBe(true)
    const f = buildFlow(a.join, 2026, { statement: 'ALL', focus: 'c:payables', topLedgers: 3 })
    expect(f.nodes.filter((n) => n.column === 0).length).toBe(4) // 3 ledgers + "others"
  })
})
