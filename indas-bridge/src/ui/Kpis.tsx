import { useStore } from '../store'
import { TieChip, useFmt } from './common'

export function Kpis() {
  const { analysis, period, set } = useStore()
  const fmt = useFmt()
  if (!analysis || period === null) return null
  const b = analysis.bridges[period]
  const { stats } = analysis.join
  const varLedgers = analysis.join.ledgers.filter((l) => Math.abs(l.diff[period] ?? 0) > 100)
  const varNet = varLedgers.reduce((a, l) => a + l.diff[period], 0)
  const varGross = varLedgers.reduce((a, l) => a + Math.abs(l.diff[period]), 0)
  const unmatched = stats.xlsbOnly + stats.copyOnly
  const t = b.totals
  const ties = b.ties.filter((x) => x.status !== 'na')
  const tieCount = ties.filter((x) => x.status === 'tie').length
  return (
    <>
      <div className="kpis" aria-label="Key figures">
        <div className="kpi">
          <div className="label">Ledgers matched</div>
          <div className="value">{stats.matched.toLocaleString('en-IN')}</div>
          <div className="sub">of {stats.signedLedgers.toLocaleString('en-IN')} in signed books{stats.matchedByCode ? ` · ${stats.matchedByCode} by code` : ''}</div>
        </div>
        <div className="kpi">
          <div className="label">Unmatched</div>
          <div className="value" style={{ color: unmatched ? 'var(--warn)' : 'var(--tie)' }}>{unmatched}</div>
          <div className="sub">{stats.xlsbOnly} signed only · {stats.copyOnly} Ind AS only</div>
        </div>
        <div className="kpi">
          <div className="label">Variances · {b.period.label.replace(' (transition)', '')}</div>
          <div className="value" style={{ color: varLedgers.length ? 'var(--bad)' : 'var(--tie)' }}>{varLedgers.length}</div>
          <div className="sub"><span className="num">{fmt(varGross, { symbol: true })}</span> gross · net <span className="num">{fmt(varNet, { symbol: true })}</span></div>
        </div>
        <div className="kpi">
          <div className="label">Profit after tax</div>
          {b.period.hasPL ? (
            <>
              <div className="value num">{fmt(t.indas.pat, { symbol: true })}</div>
              <div className="sub">Local GAAP <span className="num">{fmt(t.local.pat, { symbol: true })}</span></div>
            </>
          ) : (
            <><div className="value num">—</div><div className="sub">No P&amp;L at the transition date</div></>
          )}
        </div>
        <div className="kpi">
          <div className="label">Equity bridge</div>
          <div className="value num">{fmt(t.indas.equity - t.local.equity, { symbol: true, signed: true })}</div>
          <div className="sub"><span className="num">{fmt(t.local.equity, { symbol: true })}</span> → <span className="num">{fmt(t.indas.equity, { symbol: true })}</span></div>
        </div>
        <div className="kpi">
          <div className="label">Tie-outs</div>
          <div className="value" style={{ color: tieCount === ties.length ? 'var(--tie)' : 'var(--warn)' }}>{tieCount}/{ties.length} ✓</div>
          <div className="sub">checks against books and face statements</div>
        </div>
      </div>
      <div className="ties-row" aria-label="Tie-out checks">
        {ties.map((x) => (
          <button key={x.id} className="chip" style={{ cursor: 'pointer', background: 'none', padding: 0, border: 0 }} onClick={() => set({ tab: x.id.startsWith('fs-p') || x.id.startsWith('fs-t') || x.id.startsWith('fs-o') ? 'pl' : 'bs' })}
            title={`${x.label}: engine ${fmt(x.engine, { symbol: true })} vs ${fmt(x.reference, { symbol: true })}`}>
            <TieChip status={x.status} label={shortTie(x.id)} />
          </button>
        ))}
      </div>
    </>
  )
}

function shortTie(id: string): string {
  return ({
    'signed-nil': 'Signed TB nil', 'a-nil': 'Layer A nil', 'bs-balances': 'Ind AS BS balances', 'local-bs-balances': 'Local BS balances',
    'fs-assets': 'Total assets', 'fs-equity': 'Equity', 'fs-eql': 'Equity + liabilities', 'fs-pbt': 'PBT', 'fs-tax': 'Tax',
    'fs-pat': 'PAT', 'fs-oci': 'OCI', 'fs-tci': 'TCI',
  } as Record<string, string>)[id] ?? id
}
