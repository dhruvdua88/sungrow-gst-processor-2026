import { useStore } from '../store'
import { BridgeTable } from './BridgeTable'
import { TieChip, useFmt } from './common'
import { Waterfall } from './Waterfall'

export default function PLBridge() {
  const { analysis, period } = useStore()
  const fmt = useFmt()
  if (!analysis || period === null) return null
  const b = analysis.bridges[period]
  if (!b.period.hasPL || !b.plWaterfall) {
    return <div className="panel panel-body">There is no Ind AS statement of profit and loss for {b.period.label}: it is the opening balance sheet date (Ind AS 101). Pick FY 2025-26 or FY 2024-25 above.</div>
  }
  const t = b.totals
  const plTies = b.ties.filter((x) => ['fs-pbt', 'fs-tax', 'fs-pat', 'fs-oci', 'fs-tci'].includes(x.id))
  return (
    <div className="stack">
      <div className="panel">
        <div className="panel-head">
          <div><h2>Profit bridge · {b.period.label}</h2><p>Local GAAP profit after tax → Ind AS adjustments → Ind AS profit → OCI → total comprehensive income. Click a bar for its TB rows.</p></div>
        </div>
        <div className="panel-body"><Waterfall steps={b.plWaterfall} title={`Profit bridge ${b.period.label}`} /></div>
      </div>
      <div className="panel">
        <div className="panel-head"><h2>Ties to the Statement of Profit and Loss</h2></div>
        <div className="panel-body" style={{ overflowX: 'auto' }}>
          <table className="ledger">
            <thead><tr><th>Check</th><th className="r">Bridge</th><th className="r">Ind AS workbook</th><th className="r">Difference</th><th>Status</th></tr></thead>
            <tbody>
              {plTies.map((x) => (
                <tr key={x.id}><td>{x.label}</td><td className="r num">{fmt(x.engine)}</td><td className="r num">{fmt(x.reference)}</td><td className="r num">{fmt(x.diff, { signed: true })}</td><td><TieChip status={x.status} /></td></tr>
              ))}
            </tbody>
          </table>
          {t.indas.oci !== null && <p className="muted" style={{ fontSize: 12.5 }}>OCI is derived from the movement in Ind AS other equity less Ind AS profit (and less any other equity movements seen in the local books), so it ties only if the equity entries are complete.</p>}
        </div>
      </div>
      <BridgeTable lines={b.plLines} title="Statement of profit and loss — line by line" subtitle="Local GAAP captions use the signed face grouping; Ind AS captions use the Ind AS workbook mapping"
        filename={`pl-bridge-${b.period.key}.csv`}
        subtotals={[
          { at: 'EXP', label: 'Profit before tax', local: t.local.pbt, indas: t.indas.pbt },
          { at: 'TAX', label: 'Profit for the year', local: t.local.pat, indas: t.indas.pat, reference: b.ties.find((x) => x.id === 'fs-pat')?.reference },
        ]} />
    </div>
  )
}
