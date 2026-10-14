import { Fragment, useEffect, useRef } from 'react'
import { CAPTION_BY_ID, SECTION_LABEL } from '../engine'
import { useStore } from '../store'
import { useFmt } from './common'

export function LedgerDrawer() {
  const { analysis, selected, period, set } = useStore()
  const fmt = useFmt()
  const closeRef = useRef<HTMLButtonElement>(null)
  const led = analysis?.join.ledgers.find((l) => l.key === selected)
  useEffect(() => {
    if (!led) return
    const prev = document.activeElement as HTMLElement | null
    closeRef.current?.focus()
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') set({ selected: null }) }
    window.addEventListener('keydown', onKey)
    return () => { window.removeEventListener('keydown', onKey); prev?.focus?.() }
  }, [led, set])
  if (!analysis || !led || period === null) return null
  const k = period
  const local = CAPTION_BY_ID.get(led.localCaption[k])
  const diff = led.diff[k]
  const flagged = Math.abs(diff) > 100
  const alloc = (led.alloc[k] ?? []).filter((a, _, arr) => a.amount !== 0 || arr.length === 1)
  return (
    <>
      <div className="scrim" onClick={() => set({ selected: null })} />
      <aside className="drawer" role="dialog" aria-modal="true" aria-label={`Ledger journey: ${led.name}`}>
        <header>
          <div style={{ flex: 1 }}>
            <div className="eyebrow">Ledger journey · {analysis.bridges[k].period.label}</div>
            <h2>{led.name}</h2>
            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginTop: 8 }}>
              <span className={`chip ${led.status === 'matched' ? 'ok' : 'bad'}`}>{led.status === 'matched' ? `Matched${led.matchedBy === 'code' ? ' by code' : ''}` : led.status === 'xlsb-only' ? 'Only in signed books' : 'Only in Ind AS workbook'}</span>
              {flagged && <span className="chip bad">Variance {fmt(diff, { symbol: true, signed: true })}</span>}
              {led.regroup[k] && <span className="chip warn">Regrouped across Ind AS rows</span>}
              {led.assumedMapping && <span className="chip warn">Ind AS mapping assumed from local face</span>}
              {led.source[k] === 'copy' && <span className="chip warn">No signed column for this period: Ind AS file amount used</span>}
            </div>
          </div>
          <button ref={closeRef} className="icon-btn" aria-label="Close" onClick={() => set({ selected: null })}>✕</button>
        </header>
        <div className="body">
          <section className="journey" aria-label="Path from signed books to Ind AS">
            <Step k="Signed books" v={led.name} amt={fmt(led.used[k], { symbol: true })} n="1" />
            <Step k="Primary › parent group" v={`${led.primaryGroup || '—'} › ${led.parentGroup || '—'}`} n="2" />
            <Step k="Local face grouping (Div I)" v={`${led.face[k] || '—'}${local ? `  →  ${local.label}` : ''}`} n="3" />
            <Step k="Note grouping" v={led.note[k] || '—'} n="4" />
            <Step k="Ind AS workbook local GAAP" v={led.copyRows.length ? `${led.copyRows.length} mapping row${led.copyRows.length > 1 ? 's' : ''}` : 'Not present'} amt={fmt(led.copyLocal[k], { symbol: true })} n="5" flag={flagged || led.status !== 'matched'} />
            {alloc.map((a, i) => {
              const cap = CAPTION_BY_ID.get(a.captionId)
              return (
                <Step key={i} ind k={`Ind AS ${cap ? SECTION_LABEL[cap.section].toLowerCase() : ''}`} v={`${cap?.label ?? a.captionId} › ${a.lineItem}`} sub={a.category} amt={fmt(a.amount, { symbol: true })} n={String(6 + i)} />
              )
            })}
          </section>
          <section>
            <div className="eyebrow" style={{ marginBottom: 8 }}>All periods (Dr +, Cr −)</div>
            <table className="ledger">
              <thead><tr><th>Period</th><th className="r">Signed</th><th className="r">Ind AS file</th><th className="r">Δ</th></tr></thead>
              <tbody>
                {analysis.join.periods.map((p) => (
                  <tr key={p.key}><td>{p.label}</td><td className="r num">{fmt(led.used[p.key])}</td><td className="r num muted">{fmt(led.copyLocal[p.key])}</td>
                    <td className={`r num ${Math.abs(led.diff[p.key]) > 100 ? 'delta-neg' : 'zero'}`}>{fmt(led.diff[p.key], { signed: true })}</td></tr>
                ))}
              </tbody>
            </table>
          </section>
          {led.copyRows.length > 1 && (
            <section>
              <div className="eyebrow" style={{ marginBottom: 8 }}>How the signed amount was allocated</div>
              <p className="muted" style={{ fontSize: 13, margin: '0 0 8px' }}>The Ind AS workbook maps this ledger to several rows. If those rows add up to the signed amount their split is kept; otherwise the whole amount goes to the row that equals it, or the largest row takes the balance.</p>
              <dl className="facts">
                {led.copyRows.map((r, i) => (
                  <Fragment key={i}><dt>TB row {r.sheetRow} · {r.category}</dt><dd className="num">{fmt(r.local[k])} → {fmt(led.alloc[k]?.[i]?.amount ?? 0)}</dd></Fragment>
                ))}
              </dl>
            </section>
          )}
        </div>
      </aside>
    </>
  )
}

function Step({ k, v, amt, n, ind, flag, sub }: { k: string; v: string; amt?: string; n: string; ind?: boolean; flag?: boolean; sub?: string }) {
  return (
    <div className={`jstep${ind ? ' ind' : ''}${flag ? ' flag' : ''}`}>
      <div className="jdot">{n}</div>
      <div><div className="k">{k}</div><div className="v">{v}</div>{sub && <div className="muted" style={{ fontSize: 12 }}>{sub}</div>}</div>
      <div className="num" style={{ paddingTop: 14 }}>{amt}</div>
    </div>
  )
}
