import { useMemo, useState } from 'react'
import type { Step } from '../engine'
import { useFmt, useTooltip } from './common'

/**
 * Horizontal waterfall. Start/total bars are anchored to the left edge of a broken axis so that small
 * adjustments stay visible next to large profit or equity figures; the axis note says so.
 */
export function Waterfall({ steps, title }: { steps: Step[]; title: string }) {
  const fmt = useFmt()
  const { show, hide, node } = useTooltip()
  const [focus, setFocus] = useState<string | null>(null)
  const rows = useMemo(() => layoutSteps(steps), [steps])
  const values = rows.flatMap((r) => [r.from, r.to])
  let lo = Math.min(...values)
  let hi = Math.max(...values)
  const span = Math.max(hi - lo, Math.abs(hi) * 0.002, 1)
  lo -= span * 0.35
  hi += span * 0.12
  const W = 900, labelW = 300, valueW = 140, rowH = 38, top = 10
  const plotW = W - labelW - valueW
  const x = (v: number) => labelW + ((v - lo) / (hi - lo)) * plotW
  const H = top + rows.length * rowH + 26
  const focused = rows.find((r) => r.s.id === focus)?.s
  return (
    <div>
      <div className="chart-wrap">
        <svg viewBox={`0 0 ${W} ${H}`} width="100%" role="img" aria-label={title} style={{ minWidth: 640 }}>
          <line x1={labelW} x2={labelW} y1={top - 4} y2={H - 22} stroke="var(--rule-strong)" />
          <path d={`M${labelW - 5} ${H - 30} l5 -5 l5 5`} fill="none" stroke="var(--ink-3)" />
          {rows.map((r, i) => {
            const y = top + i * rowH
            const isAnchor = r.s.kind !== 'step'
            const x0 = isAnchor ? labelW : x(Math.min(r.from, r.to))
            const x1 = x(isAnchor ? r.to : Math.max(r.from, r.to))
            const up = r.s.amount >= 0
            const fill = isAnchor ? 'var(--accent)' : r.s.tone === 'warn' ? 'var(--warn)' : up ? 'var(--debit)' : 'var(--credit)'
            const prev = rows[i - 1]
            return (
              <g key={r.s.id} tabIndex={0} role="button" aria-label={`${r.s.label}: ${fmt(r.s.amount, { symbol: true, signed: !isAnchor })}`}
                onMouseMove={(e) => show(e, <><b>{r.s.label}</b><span className="num">{fmt(r.s.amount, { symbol: true, signed: !isAnchor })}</span>{r.s.detail && <div className="muted" style={{ whiteSpace: 'pre-line', marginTop: 4 }}>{r.s.detail}</div>}</>)}
                onMouseLeave={hide} onClick={() => setFocus(r.s.id === focus ? null : r.s.id)} onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setFocus(r.s.id === focus ? null : r.s.id) } }}
                style={{ cursor: 'pointer', outline: 'none' }}>
                {i % 2 === 0 && <rect x={0} y={y} width={W} height={rowH} fill="var(--grid-line)" />}
                {prev && !isAnchor && <line x1={x(r.from)} x2={x(r.from)} y1={y - rowH + 8} y2={y + 8} stroke="var(--ink-3)" strokeDasharray="2 3" />}
                <text x={labelW - 12} y={y + rowH / 2 + 4} textAnchor="end" fontSize="13" style={{ fill: isAnchor ? 'var(--ink)' : 'var(--ink-2)', fontWeight: isAnchor ? 600 : 400 }}>
                  {r.s.label.length > 44 ? `${r.s.label.slice(0, 43)}…` : r.s.label}
                </text>
                <rect x={x0} y={y + 8} width={Math.max(2, x1 - x0)} height={rowH - 16} rx={3} fill={fill} opacity={focus && focus !== r.s.id ? 0.35 : 0.92} />
                <text x={W - 6} y={y + rowH / 2 + 4} textAnchor="end" fontSize="13" fontFamily="var(--font-mono)" style={{ fill: isAnchor ? 'var(--ink)' : up ? 'var(--debit)' : 'var(--credit)' }}>
                  {fmt(r.s.amount, { signed: !isAnchor })}
                </text>
              </g>
            )
          })}
          <text x={labelW + 4} y={H - 6} fontSize="11" style={{ fill: 'var(--ink-3)' }}>Axis starts near the lowest running total (broken axis) so small adjustments stay visible.</text>
        </svg>
      </div>
      <div className="legend" style={{ marginTop: 8 }}>
        <span><i style={{ background: 'var(--accent)' }} />Totals</span>
        <span><i style={{ background: 'var(--debit)' }} />Increase</span>
        <span><i style={{ background: 'var(--credit)' }} />Decrease</span>
        <span><i style={{ background: 'var(--warn)' }} />Reclassification / unexplained</span>
      </div>
      {focused?.detail && (
        <div className="panel" style={{ marginTop: 12, padding: 12, fontSize: 13 }}>
          <b>{focused.label}</b>
          {focused.rows && <div className="muted">Ind AS workbook TB rows: {focused.rows.join(', ')}</div>}
          <div style={{ whiteSpace: 'pre-line', marginTop: 6 }}>{focused.detail}</div>
        </div>
      )}
      {node}
    </div>
  )
}

function layoutSteps(steps: Step[]) {
  const out: { s: Step; from: number; to: number }[] = []
  let run = 0
  for (const s of steps) {
    if (s.kind === 'start') run = s.amount
    if (s.kind !== 'step') { out.push({ s, from: s.amount, to: s.amount }); continue }
    out.push({ s, from: run, to: run + s.amount })
    run += s.amount
  }
  return out
}
