import { sankey, sankeyLinkHorizontal, type SankeyGraph, type SankeyLink, type SankeyNode } from 'd3-sankey'
import { useLayoutEffect, useMemo, useRef, useState } from 'react'
import { buildFlow, CAPTIONS, SECTION_LABEL, SECTION_ORDER, type FlowLink, type FlowNode, type JoinedLedger } from '../engine'
import { useStore } from '../store'
import { useFmt, useTooltip } from './common'

type SNode = SankeyNode<FlowNode, FlowLink>
type SLink = SankeyLink<FlowNode, FlowLink & { raw: number }>
const rawOf = (n: SNode) => {
  const out = (n.sourceLinks ?? []) as SLink[]
  const inn = (n.targetLinks ?? []) as SLink[]
  return Math.max(out.reduce((a, l) => a + l.raw, 0), inn.reduce((a, l) => a + l.raw, 0))
}

const COLUMN_TITLES = ['Signed books · primary group', 'Local face grouping + Ind AS JV', 'Ind AS caption', 'Statement section']
const COLUMN_TITLES_FOCUS = ['Ledgers', 'Local face grouping + Ind AS JV', 'Ind AS caption', 'Statement section']

export default function FlowView() {
  const { flowMode, set } = useStore()
  return (
    <div className="panel">
      <div className="panel-head">
        <div><h2>Flow · signed books to Ind AS</h2><p>{flowMode === 'sankey' ? 'Band width = gross amount (compressed scale by default). Hover for exact figures; click a node or band to drill into its ledgers.' : 'Local groupings on the left, Ind AS statements on the right. Hover to trace a path; expand a group to see its ledgers.'}</p></div>
        <div className="spacer" />
        <div className="seg" role="group" aria-label="Flow view">
          <button aria-pressed={flowMode === 'sankey'} onClick={() => set({ flowMode: 'sankey' })}>Sankey</button>
          <button aria-pressed={flowMode === 'tree'} onClick={() => set({ flowMode: 'tree' })}>Linked tree</button>
        </div>
      </div>
      <div className="panel-body" key={flowMode} style={{ animation: 'rise .3s' }}>
        {flowMode === 'sankey' ? <SankeyView /> : <LinkedTree />}
      </div>
    </div>
  )
}

// ====================================================================== Sankey

function SankeyView() {
  const { analysis, period, set } = useStore()
  const fmt = useFmt()
  const { show, hide, node: tip } = useTooltip()
  const [statement, setStatement] = useState<'ALL' | 'BS' | 'PL'>('BS')
  const [focus, setFocus] = useState<{ id: string; label: string } | null>(null)
  const [drill, setDrill] = useState<{ label: string; keys: string[] } | null>(null)
  const [hover, setHover] = useState<string | null>(null)
  const [compressed, setCompressed] = useState(true)
  const W = 1180
  const graph = useMemo(() => {
    if (!analysis || period === null) return null
    const g = buildFlow(analysis.join, period, { statement, focus: focus?.id ?? null })
    if (!g.links.length) return null
    const H = Math.max(520, Math.min(1100, g.nodes.filter((n) => n.column === 1).length * 26 + 80))
    const layout = sankey<FlowNode, FlowLink & { raw: number }>()
      .nodeId((n) => n.id)
      .nodeAlign((n) => (n as SNode).column)
      .nodeWidth(12)
      .nodePadding(focus ? 10 : 7)
      .nodeSort((a, b) => (b.value ?? 0) - (a.value ?? 0))
      .extent([[8, 28], [W - 8, H - 8]])
    const data: SankeyGraph<FlowNode, FlowLink & { raw: number }> = layout({ nodes: g.nodes.map((n) => ({ ...n })), links: g.links.map((l) => ({ ...l, raw: l.value, value: compressed ? Math.sqrt(l.value) : l.value })) })
    return { ...data, H }
  }, [analysis, period, statement, focus, compressed])
  if (!analysis || period === null) return null
  const byKey = new Map(analysis.join.ledgers.map((l) => [l.key, l]))
  const path = sankeyLinkHorizontal<FlowNode, FlowLink & { raw: number }>()
  const lit = (l: SLink) => !hover || (l.source as SNode).id === hover || (l.target as SNode).id === hover || `${(l.source as SNode).id}→${(l.target as SNode).id}` === hover
  const titles = focus ? COLUMN_TITLES_FOCUS : COLUMN_TITLES
  const columnsX = graph ? [0, 1, 2, 3].map((c) => graph.nodes.find((n) => n.column === c)?.x0 ?? null) : []
  const onNode = (n: SNode) => {
    if (n.kind === 'ledger') { set({ selected: n.id.slice(2) }); return }
    if (n.kind === 'jv') { set({ tab: 'entries' }); return }
    setDrill({ label: n.label, keys: n.ledgerKeys })
    if (n.kind !== 'others' && n.kind !== 'section') setFocus({ id: n.id, label: n.label })
  }
  return (
    <div>
      <div className="toolbar" style={{ padding: '0 0 12px' }}>
        <div className="seg" role="group" aria-label="Statement">
          {(['BS', 'PL', 'ALL'] as const).map((s) => (
            <button key={s} aria-pressed={statement === s} onClick={() => { setStatement(s); setFocus(null); setDrill(null) }}>{s === 'BS' ? 'Balance sheet' : s === 'PL' ? 'Profit & loss' : 'Both'}</button>
          ))}
        </div>
        <div className="seg" role="group" aria-label="Band scale">
          <button aria-pressed={compressed} onClick={() => setCompressed(true)} title="Band width proportional to the square root of the amount, so small flows stay visible">Compressed</button>
          <button aria-pressed={!compressed} onClick={() => setCompressed(false)}>Linear</button>
        </div>
        {focus && <button className="chip accent" style={{ cursor: 'pointer' }} onClick={() => { setFocus(null); setDrill(null) }}>Focused on: {focus.label} ✕</button>}
        <div className="spacer" />
        <div className="legend"><span><i style={{ background: 'var(--debit)' }} />Debit balances</span><span><i style={{ background: 'var(--credit)' }} />Credit balances</span><span><i style={{ background: 'var(--warn)' }} />Ind AS JV</span></div>
      </div>
      {!graph ? <p className="muted">Nothing to show for this selection.</p> : (
        <div className="chart-wrap">
          <svg viewBox={`0 0 ${W} ${graph.H}`} width="100%" style={{ minWidth: 820 }} role="img" aria-label="Sankey diagram of ledger flows from signed books to Ind AS statements">
            {titles.map((t, i) => columnsX[i] !== null && (
              <text key={t} x={(columnsX[i] ?? 0) + (i === 3 ? 12 : 0)} y={14} fontSize="11" textAnchor={i === 3 ? 'end' : i === 2 ? 'middle' : 'start'} style={{ fill: 'var(--ink-3)', letterSpacing: '.08em', textTransform: 'uppercase', fontFamily: 'var(--font-mono)' }}>{t}</text>
            ))}
            <g fill="none">
              {graph.links.map((l, i) => {
                const s = l.source as SNode, t = l.target as SNode
                const color = s.kind === 'jv' ? 'var(--warn)' : l.dr ? 'var(--debit)' : 'var(--credit)'
                return (
                  <path key={i} d={path(l) ?? ''} stroke={color} strokeWidth={Math.max(1, l.width ?? 1)} strokeOpacity={lit(l) ? (hover ? 0.6 : 0.32) : 0.06}
                    style={{ cursor: 'pointer', transition: 'stroke-opacity .15s' }}
                    onMouseMove={(e) => { setHover(`${s.id}→${t.id}`); show(e, <><b>{s.label} → {t.label}</b><span className="num">{fmt(l.raw, { symbol: true })}</span> <span className="muted">({l.dr ? 'net debit' : 'net credit'}{l.ledgerKeys.length ? ` · ${l.ledgerKeys.length} ledgers` : ''})</span></>) }}
                    onMouseLeave={() => { setHover(null); hide() }}
                    onClick={() => l.ledgerKeys.length ? setDrill({ label: `${s.label} → ${t.label}`, keys: l.ledgerKeys }) : set({ tab: 'entries' })} />
                )
              })}
            </g>
            {graph.nodes.map((n) => {
              const h = Math.max(1, (n.y1 ?? 0) - (n.y0 ?? 0))
              const right = n.column >= 2
              const fill = n.kind === 'jv' ? 'var(--warn)' : n.kind === 'caption' || n.kind === 'section' ? 'var(--accent)' : 'var(--ink-2)'
              const showLabel = h > 7 || n.kind === 'section' || n.kind === 'jv' || hover === n.id
              return (
                <g key={n.id} tabIndex={0} role="button" aria-label={`${n.label}: ${fmt(rawOf(n), { symbol: true })}`} style={{ cursor: 'pointer', outline: 'none' }}
                  onMouseMove={(e) => { setHover(n.id); show(e, <><b>{n.label}</b><span className="num">{fmt(rawOf(n), { symbol: true })}</span>{n.ledgerKeys.length ? <span className="muted"> · {n.ledgerKeys.length} ledgers</span> : null}<div className="muted" style={{ marginTop: 2 }}>Click to {n.kind === 'ledger' ? 'open its journey' : n.kind === 'jv' ? 'see the entries' : 'focus and list ledgers'}</div></>) }}
                  onMouseLeave={() => { setHover(null); hide() }}
                  onFocus={() => setHover(n.id)} onBlur={() => setHover(null)}
                  onClick={() => onNode(n)} onKeyDown={(e) => { if (e.key === 'Enter') onNode(n) }}>
                  <rect x={n.x0} y={n.y0} width={(n.x1 ?? 0) - (n.x0 ?? 0)} height={h} rx={2} fill={fill} opacity={hover && hover !== n.id ? 0.6 : 1} />
                  {showLabel && (
                    <text x={right ? (n.x0 ?? 0) - 6 : (n.x1 ?? 0) + 6} y={((n.y0 ?? 0) + (n.y1 ?? 0)) / 2 + 4} fontSize="11.5" textAnchor={right ? 'end' : 'start'} style={{ fill: 'var(--ink)', paintOrder: 'stroke', stroke: 'var(--panel)', strokeWidth: 3 }}>
                      {n.label.length > 38 ? `${n.label.slice(0, 37)}…` : n.label}
                    </text>
                  )}
                </g>
              )
            })}
          </svg>
        </div>
      )}
      {drill && <DrillList label={drill.label} ledgers={drill.keys.map((k) => byKey.get(k)!).filter(Boolean)} onClose={() => setDrill(null)} />}
      {tip}
    </div>
  )
}

function DrillList({ label, ledgers, onClose }: { label: string; ledgers: JoinedLedger[]; onClose: () => void }) {
  const { period, set } = useStore()
  const fmt = useFmt()
  const sorted = [...ledgers].sort((a, b) => Math.abs(b.used[period!]) - Math.abs(a.used[period!]))
  return (
    <div className="panel" style={{ marginTop: 16 }}>
      <div className="panel-head"><div><h2 style={{ fontSize: 17 }}>{label}</h2><p>{ledgers.length} ledgers · click one for its journey</p></div><div className="spacer" /><button className="icon-btn" aria-label="Close list" onClick={onClose}>✕</button></div>
      <div style={{ maxHeight: 340, overflow: 'auto' }}>
        <table className="ledger">
          <tbody>
            {sorted.slice(0, 300).map((l) => (
              <tr key={l.key} className="click" tabIndex={0} onClick={() => set({ selected: l.key })} onKeyDown={(e) => { if (e.key === 'Enter') set({ selected: l.key }) }}>
                <td>{l.name}{Math.abs(l.diff[period!]) > 100 && <span className="chip bad" style={{ marginLeft: 6, fontSize: 11 }}>Δ {fmt(l.diff[period!], { signed: true })}</span>}</td>
                <td className="muted">{l.face[period!]}</td>
                <td className="r num">{fmt(l.used[period!])}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {sorted.length > 300 && <p className="muted" style={{ padding: '0 12px' }}>Showing the 300 largest. Use the Ledger explorer for the full list.</p>}
      </div>
    </div>
  )
}

// ====================================================================== Linked tree

interface LeftGroup { face: string; total: number; notes: Map<string, JoinedLedger[]>; targets: Map<string, number> }

function LinkedTree() {
  const { analysis, period, set } = useStore()
  const fmt = useFmt()
  const boxRef = useRef<HTMLDivElement>(null)
  const [open, setOpen] = useState<Set<string>>(new Set())
  const [hover, setHover] = useState<{ faces: Set<string>; captions: Set<string> } | null>(null)
  const [geo, setGeo] = useState<{ w: number; h: number; left: Record<string, number>; right: Record<string, number>; lx: number; rx: number }>({ w: 0, h: 0, left: {}, right: {}, lx: 0, rx: 0 })

  const model = useMemo(() => {
    if (!analysis || period === null) return null
    const k = period
    const groups = new Map<string, LeftGroup>()
    const capTotals = new Map<string, number>()
    for (const l of analysis.join.ledgers) {
      const used = l.used[k] ?? 0
      if (used === 0) continue
      const face = l.face[k] || '(no face grouping)'
      const g = groups.get(face) ?? { face, total: 0, notes: new Map(), targets: new Map() }
      g.total += used
      const note = l.note[k] || '(no note)'
      g.notes.set(note, [...(g.notes.get(note) ?? []), l])
      for (const a of l.alloc[k] ?? []) {
        if (a.amount === 0) continue
        g.targets.set(a.captionId, (g.targets.get(a.captionId) ?? 0) + Math.abs(a.amount))
        capTotals.set(a.captionId, (capTotals.get(a.captionId) ?? 0) + a.amount)
      }
      groups.set(face, g)
    }
    const jv = new Map<string, number>()
    for (const a of analysis.join.adjustments) {
      const v = a.row.jv[k] ?? 0
      if (v) jv.set(a.captionId, (jv.get(a.captionId) ?? 0) + v)
    }
    const left = [...groups.values()].sort((a, b) => Math.abs(b.total) - Math.abs(a.total))
    const caps = CAPTIONS.filter((c) => capTotals.has(c.id) || jv.has(c.id))
    const max = Math.max(1, ...left.flatMap((g) => [...g.targets.values()]))
    return { left, caps, capTotals, jv, max }
  }, [analysis, period])

  useLayoutEffect(() => {
    const box = boxRef.current
    if (!box) return
    const measure = () => {
      const r = box.getBoundingClientRect()
      const left: Record<string, number> = {}
      const right: Record<string, number> = {}
      let lx = 0, rx = r.width
      box.querySelectorAll<HTMLElement>('[data-face]').forEach((el) => {
        const e = el.getBoundingClientRect()
        left[el.dataset.face!] = e.top - r.top + Math.min(e.height, 34) / 2
        lx = Math.max(lx, e.right - r.left)
      })
      box.querySelectorAll<HTMLElement>('[data-cap]').forEach((el) => {
        const e = el.getBoundingClientRect()
        right[el.dataset.cap!] = e.top - r.top + e.height / 2
        rx = Math.min(rx, e.left - r.left)
      })
      setGeo({ w: r.width, h: r.height, left, right, lx, rx })
    }
    measure()
    const ro = new ResizeObserver(measure)
    ro.observe(box)
    return () => ro.disconnect()
  }, [model, open])

  if (!analysis || period === null || !model) return null
  const toggle = (f: string) => setOpen((s) => { const n = new Set(s); if (n.has(f)) n.delete(f); else n.add(f); return n })
  const hoverLedger = (l: JoinedLedger) => setHover({ faces: new Set([l.face[period] || '(no face grouping)']), captions: new Set((l.alloc[period] ?? []).filter((a) => a.amount !== 0).map((a) => a.captionId)) })
  const sections = SECTION_ORDER.filter((s) => model.caps.some((c) => c.section === s))
  return (
    <div ref={boxRef} style={{ position: 'relative', display: 'grid', gridTemplateColumns: 'minmax(260px, 1fr) minmax(80px, 0.55fr) minmax(260px, 1fr)', gap: 0 }} onMouseLeave={() => setHover(null)}>
      <svg width={geo.w} height={geo.h} style={{ position: 'absolute', inset: 0, pointerEvents: 'none' }} aria-hidden="true">
        {model.left.flatMap((g) => [...g.targets].map(([cid, v]) => {
          const y1 = geo.left[g.face], y2 = geo.right[cid]
          if (y1 === undefined || y2 === undefined) return null
          const active = !hover || (hover.faces.has(g.face) && hover.captions.has(cid))
          const x1 = geo.lx + 4, x2 = geo.rx - 4, mx = (x1 + x2) / 2
          return <path key={`${g.face}|${cid}`} d={`M${x1} ${y1} C ${mx} ${y1}, ${mx} ${y2}, ${x2} ${y2}`} fill="none" stroke={active && hover ? 'var(--accent)' : 'var(--ink-3)'} strokeOpacity={active ? (hover ? 0.9 : 0.35) : 0.05} strokeWidth={1 + 7 * Math.sqrt(v / model.max)} style={{ transition: 'stroke-opacity .15s' }} />
        }))}
      </svg>
      <div>
        <div className="eyebrow" style={{ marginBottom: 8 }}>Signed books · face › note › ledger</div>
        {model.left.map((g) => (
          <div key={g.face} style={{ marginBottom: 4 }}>
            <button data-face={g.face} className="btn ghost" aria-expanded={open.has(g.face)} style={{ width: '100%', justifyContent: 'space-between', padding: '6px 10px', fontSize: 13, background: hover?.faces.has(g.face) ? 'var(--accent-soft)' : undefined }}
              onClick={() => toggle(g.face)}
              onMouseEnter={() => setHover({ faces: new Set([g.face]), captions: new Set(g.targets.keys()) })}
              onFocus={() => setHover({ faces: new Set([g.face]), captions: new Set(g.targets.keys()) })}>
              <span style={{ textAlign: 'left' }}>{open.has(g.face) ? '▾' : '▸'} {g.face}</span><span className="num muted">{fmt(g.total)}</span>
            </button>
            {open.has(g.face) && (
              <div style={{ padding: '4px 0 6px 18px', fontSize: 12.5 }}>
                {[...g.notes].map(([note, ls]) => (
                  <div key={note} style={{ marginBottom: 4 }}>
                    <div className="muted" style={{ fontFamily: 'var(--font-mono)', fontSize: 11, textTransform: 'uppercase', letterSpacing: '.06em' }}>{note} · {ls.length}</div>
                    {ls.sort((a, b) => Math.abs(b.used[period]) - Math.abs(a.used[period])).slice(0, 25).map((l) => (
                      <div key={l.key} role="button" tabIndex={0} style={{ display: 'flex', justifyContent: 'space-between', gap: 8, padding: '2px 6px', cursor: 'pointer', borderRadius: 4 }}
                        onMouseEnter={() => hoverLedger(l)} onFocus={() => hoverLedger(l)} onClick={() => set({ selected: l.key })} onKeyDown={(e) => { if (e.key === 'Enter') set({ selected: l.key }) }}>
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{l.name}</span><span className="num">{fmt(l.used[period])}</span>
                      </div>
                    ))}
                    {ls.length > 25 && <div className="muted" style={{ padding: '0 6px' }}>+ {ls.length - 25} more in the Ledger explorer</div>}
                  </div>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>
      <div />
      <div>
        <div className="eyebrow" style={{ marginBottom: 8 }}>Ind AS · section › caption</div>
        {sections.map((s) => (
          <div key={s} style={{ marginBottom: 10 }}>
            <div style={{ fontFamily: 'var(--font-display)', fontWeight: 600, fontSize: 14, margin: '6px 0 4px' }}>{SECTION_LABEL[s]}</div>
            {model.caps.filter((c) => c.section === s).map((c) => {
              const g = model.capTotals.get(c.id) ?? 0
              const j = model.jv.get(c.id) ?? 0
              const active = hover?.captions.has(c.id)
              return (
                <div key={c.id} data-cap={c.id} role="button" tabIndex={0} onClick={() => set({ tab: 'ledgers', captionFilter: c.id })} onKeyDown={(e) => { if (e.key === 'Enter') set({ tab: 'ledgers', captionFilter: c.id }) }}
                  onMouseEnter={() => setHover({ faces: new Set(model.left.filter((x) => x.targets.has(c.id)).map((x) => x.face)), captions: new Set([c.id]) })}
                  style={{ display: 'flex', justifyContent: 'space-between', gap: 8, padding: '5px 10px', marginBottom: 3, borderRadius: 6, border: `1px solid ${active ? 'var(--accent)' : 'var(--rule)'}`, background: active ? 'var(--accent-soft)' : 'var(--panel-2)', fontSize: 13, cursor: 'pointer' }}>
                  <span>{c.label}{j !== 0 && <span className="chip warn" style={{ marginLeft: 6, fontSize: 10.5, padding: '0 6px' }} title="Includes Ind AS JV">+JV</span>}</span>
                  <span className="num">{fmt(g + j)}</span>
                </div>
              )
            })}
          </div>
        ))}
        <p className="muted" style={{ fontSize: 12 }}>Amounts signed (Dr +, Cr −), including Ind AS JV. Other equity here excludes the profit for the year.</p>
      </div>
    </div>
  )
}

