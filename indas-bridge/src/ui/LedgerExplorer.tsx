import { useVirtualizer } from '@tanstack/react-virtual'
import { useDeferredValue, useMemo, useRef, useState } from 'react'
import { CAPTION_BY_ID, type JoinedLedger, type PeriodKey } from '../engine'
import { useStore } from '../store'
import { csvAmount, downloadCsv, ExportButton, UNIT_LABEL, useFmt } from './common'

export interface LedgerRow {
  l: JoinedLedger
  name: string
  primary: string
  face: string
  note: string
  caption: string
  captionId: string
  line: string
  xlsb: number
  copy: number
  diff: number
  status: string
}

export function ledgerRows(ledgers: JoinedLedger[], k: PeriodKey): LedgerRow[] {
  return ledgers.map((l) => {
    const alloc = l.alloc[k] ?? []
    const main = alloc.reduce((b, a) => (Math.abs(a.amount) > Math.abs(b.amount) ? a : b), alloc[0])
    const status = l.status !== 'matched' ? (l.status === 'xlsb-only' ? 'Signed only' : 'Ind AS only') : Math.abs(l.diff[k]) > 100 ? 'Variance' : l.regroup[k] ? 'Regrouped' : 'Matched'
    return {
      l,
      name: l.name,
      primary: l.primaryGroup || '—',
      face: l.face[k] || '—',
      note: l.note[k] || '—',
      captionId: main?.captionId ?? '',
      caption: CAPTION_BY_ID.get(main?.captionId ?? '')?.label ?? '—',
      line: alloc.length > 1 ? alloc.filter((a) => a.amount !== 0).map((a) => a.lineItem).join(' + ') || main.lineItem : main?.lineItem ?? '—',
      xlsb: l.used[k] ?? 0,
      copy: l.copyLocal[k] ?? 0,
      diff: l.diff[k] ?? 0,
      status,
    }
  })
}

type GroupBy = 'none' | 'primary' | 'face' | 'note' | 'caption' | 'line'
type SortKey = 'name' | 'xlsb' | 'copy' | 'diff'
const GROUP_LABEL: Record<GroupBy, string> = { none: 'No grouping', primary: 'Primary group (signed)', face: 'Face grouping (signed)', note: 'Note grouping (signed)', caption: 'Ind AS caption', line: 'Ind AS line item' }

type VRow = { kind: 'group'; key: string; label: string; count: number; xlsb: number; copy: number; diff: number } | { kind: 'row'; r: LedgerRow }

export default function LedgerExplorer() {
  const { analysis, period, unit, set, captionFilter } = useStore()
  const fmt = useFmt()
  const [q, setQ] = useState('')
  const dq = useDeferredValue(q)
  const [groupBy, setGroupBy] = useState<GroupBy>('caption')
  const [statusF, setStatusF] = useState('all')
  const [sort, setSort] = useState<{ key: SortKey; dir: 1 | -1 }>({ key: 'xlsb', dir: -1 })
  const [hideZero, setHideZero] = useState(true)
  const parentRef = useRef<HTMLDivElement>(null)

  const rows = useMemo(() => (analysis && period !== null ? ledgerRows(analysis.join.ledgers, period) : []), [analysis, period])
  const filtered = useMemo(() => {
    const needle = dq.trim().toLowerCase()
    let out = rows.filter((r) =>
      (!needle || `${r.name} ${r.face} ${r.note} ${r.caption} ${r.line} ${r.primary}`.toLowerCase().includes(needle)) &&
      (statusF === 'all' || (statusF === 'issues' ? r.status !== 'Matched' : r.status === statusF)) &&
      (!captionFilter || r.l.alloc[period!]?.some((a) => a.captionId === captionFilter) || r.l.localCaption[period!] === captionFilter) &&
      (!hideZero || r.xlsb !== 0 || r.copy !== 0))
    const { key, dir } = sort
    out = [...out].sort((a, b) => (key === 'name' ? a.name.localeCompare(b.name) : Math.abs(a[key]) - Math.abs(b[key])) * dir)
    return out
  }, [rows, dq, statusF, captionFilter, period, sort, hideZero])

  const vrows: VRow[] = useMemo(() => {
    if (groupBy === 'none') return filtered.map((r) => ({ kind: 'row' as const, r }))
    const groups = new Map<string, LedgerRow[]>()
    for (const r of filtered) {
      const k = r[groupBy]
      const list = groups.get(k)
      if (list) list.push(r)
      else groups.set(k, [r])
    }
    const out: VRow[] = []
    const sorted = [...groups].sort((a, b) => Math.abs(b[1].reduce((s, r) => s + r.xlsb, 0)) - Math.abs(a[1].reduce((s, r) => s + r.xlsb, 0)))
    for (const [label, list] of sorted) {
      out.push({ kind: 'group', key: `g:${label}`, label, count: list.length, xlsb: list.reduce((s, r) => s + r.xlsb, 0), copy: list.reduce((s, r) => s + r.copy, 0), diff: list.reduce((s, r) => s + r.diff, 0) })
      for (const r of list) out.push({ kind: 'row', r })
    }
    return out
  }, [filtered, groupBy])

  // eslint-disable-next-line react-hooks/incompatible-library -- TanStack Virtual is used as documented
  const virt = useVirtualizer({ count: vrows.length, getScrollElement: () => parentRef.current, estimateSize: () => 36, overscan: 12 })
  if (!analysis || period === null) return null

  const totals = filtered.reduce((a, r) => ({ xlsb: a.xlsb + r.xlsb, copy: a.copy + r.copy, diff: a.diff + r.diff }), { xlsb: 0, copy: 0, diff: 0 })
  const exportCsv = () =>
    downloadCsv(`ledgers-${period}.csv`, ['Ledger', 'Primary group', 'Face grouping', 'Note grouping', 'Ind AS caption', 'Ind AS line item', `Signed books (${UNIT_LABEL[unit]})`, 'Ind AS workbook local GAAP', 'Difference', 'Status'],
      filtered.map((r) => [r.name, r.primary, r.face, r.note, r.caption, r.line, csvAmount(r.xlsb, unit), csvAmount(r.copy, unit), csvAmount(r.diff, unit), r.status]))
  const cols = 'minmax(220px,2.4fr) minmax(120px,1fr) minmax(160px,1.3fr) minmax(130px,1.1fr) minmax(130px,1.1fr) minmax(110px,.8fr)'
  const sortBtn = (key: SortKey, label: string, right = false) => (
    <button className={right ? 'r' : ''} onClick={() => setSort((s) => ({ key, dir: s.key === key ? (s.dir === 1 ? -1 : 1) : -1 }))} aria-sort={sort.key === key ? (sort.dir === 1 ? 'ascending' : 'descending') : 'none'}>
      {label}{sort.key === key ? (sort.dir === 1 ? ' ↑' : ' ↓') : ''}
    </button>
  )
  return (
    <div className="panel">
      <div className="panel-head">
        <div><h2>Ledger explorer · {analysis.bridges[period].period.label}</h2><p>{filtered.length.toLocaleString('en-IN')} of {rows.length.toLocaleString('en-IN')} ledgers · signed amounts (Dr +, Cr −) in {UNIT_LABEL[unit]} · click a row for its journey</p></div>
        <div className="spacer" />
        <ExportButton onClick={exportCsv} />
      </div>
      <div className="toolbar">
        <input className="search" type="search" placeholder="Search ledger, group or Ind AS line…" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Search ledgers" />
        <select className="sel" value={groupBy} onChange={(e) => setGroupBy(e.target.value as GroupBy)} aria-label="Group by">
          {Object.entries(GROUP_LABEL).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
        </select>
        <select className="sel" value={statusF} onChange={(e) => setStatusF(e.target.value)} aria-label="Status filter">
          <option value="all">All statuses</option><option value="issues">Any issue</option><option>Variance</option><option>Regrouped</option><option>Signed only</option><option>Ind AS only</option><option>Matched</option>
        </select>
        <label style={{ fontSize: 13, display: 'flex', gap: 6, alignItems: 'center' }}><input type="checkbox" checked={hideZero} onChange={(e) => setHideZero(e.target.checked)} />Hide nil balances</label>
        {captionFilter && <button className="chip accent" style={{ cursor: 'pointer' }} onClick={() => set({ captionFilter: null })}>Caption: {CAPTION_BY_ID.get(captionFilter)?.label} ✕</button>}
      </div>
      <div className="vt" ref={parentRef} style={{ ['--cols' as string]: cols }}>
        <div className="vt-head" style={{ minWidth: 980 }}>
          {sortBtn('name', 'Ledger')}<button>Local group</button><button>Ind AS line</button>{sortBtn('xlsb', 'Signed books', true)}{sortBtn('copy', 'Ind AS file', true)}{sortBtn('diff', 'Δ', true)}
        </div>
        <div style={{ height: virt.getTotalSize(), position: 'relative', minWidth: 980 }}>
          {virt.getVirtualItems().map((vi) => {
            const v = vrows[vi.index]
            const style = { transform: `translateY(${vi.start}px)`, height: vi.size }
            if (v.kind === 'group') {
              return (
                <div key={v.key} className="vt-row group" style={style}>
                  <div>{v.label} <span className="muted num" style={{ fontWeight: 400 }}>· {v.count}</span></div><div /><div />
                  <div className="r num">{fmt(v.xlsb)}</div><div className="r num">{fmt(v.copy)}</div>
                  <div className={`r num ${Math.abs(v.diff) > 100 ? 'delta-neg' : 'zero'}`}>{fmt(v.diff, { signed: true })}</div>
                </div>
              )
            }
            const r = v.r
            return (
              <div key={r.l.key} className="vt-row" style={style} tabIndex={0} role="button" aria-label={`Open ${r.name}`}
                onClick={() => set({ selected: r.l.key })} onKeyDown={(e) => { if (e.key === 'Enter') set({ selected: r.l.key }) }}>
                <div title={r.name}>
                  {r.status !== 'Matched' && <span className={`chip ${r.status === 'Variance' || r.status.endsWith('only') ? 'bad' : 'warn'}`} style={{ marginRight: 6, padding: '0 6px', fontSize: 11 }}>{r.status}</span>}
                  {r.name}
                </div>
                <div title={`${r.primary} › ${r.face} › ${r.note}`} className="muted">{r.face}</div>
                <div title={`${r.caption} › ${r.line}`}>{r.line}</div>
                <div className="r num">{fmt(r.xlsb)}</div>
                <div className="r num muted">{fmt(r.copy)}</div>
                <div className={`r num ${Math.abs(r.diff) > 100 ? 'delta-neg' : 'zero'}`}>{fmt(r.diff, { signed: true })}</div>
              </div>
            )
          })}
        </div>
      </div>
      <div className="toolbar" style={{ borderTop: '1px solid var(--rule)', fontSize: 13 }}>
        <span className="muted">Filtered totals</span><div className="spacer" />
        <span>Signed <b className="num">{fmt(totals.xlsb, { symbol: true })}</b></span>
        <span>Ind AS file <b className="num">{fmt(totals.copy, { symbol: true })}</b></span>
        <span>Δ <b className="num">{fmt(totals.diff, { symbol: true, signed: true })}</b></span>
      </div>
    </div>
  )
}
