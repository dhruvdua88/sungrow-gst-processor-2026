import { useStore } from '../store'
import { BridgeTable } from './BridgeTable'
import { TieChip, useFmt } from './common'
import { Waterfall } from './Waterfall'

export default function BSBridge() {
  const { analysis, period } = useStore()
  const fmt = useFmt()
  if (!analysis || period === null) return null
  const b = analysis.bridges[period]
  const t = b.totals
  const ref = (id: string) => b.ties.find((x) => x.id === id)?.reference
  const bsTies = b.ties.filter((x) => !['fs-pbt', 'fs-tax', 'fs-pat', 'fs-oci', 'fs-tci'].includes(x.id))
  return (
    <div className="stack">
      <div className="stack">
        <div className="panel">
          <div className="panel-head"><div><h2>Equity bridge · {b.period.label}</h2><p>Local GAAP equity → effect of each Ind AS standard on net assets → Ind AS equity</p></div></div>
          <div className="panel-body"><Waterfall steps={b.equityWaterfall} title={`Equity bridge ${b.period.label}`} /></div>
        </div>
        <div className="panel">
          <div className="panel-head"><h2>Tie-out checks</h2></div>
          <div className="panel-body" style={{ overflowX: 'auto' }}>
            <table className="ledger">
              <thead><tr><th>Check</th><th className="r">Bridge</th><th className="r">Reference</th><th>Status</th></tr></thead>
              <tbody>
                {bsTies.map((x) => (
                  <tr key={x.id} title={x.diff ? `Difference ${fmt(x.diff, { symbol: true, signed: true })}` : undefined}>
                    <td>{x.label}</td><td className="r num">{fmt(x.engine)}</td><td className="r num">{fmt(x.reference)}</td><td><TieChip status={x.status} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
      <BridgeTable lines={b.bsLines} title={`Balance sheet — line by line · ${b.period.bsDate}`} subtitle="Other equity includes the profit for the year; Division I captions with no Ind AS equivalent are reclassified in full"
        filename={`bs-bridge-${b.period.key}.csv`}
        subtotals={[
          { at: 'CA', label: 'Total assets', local: t.local.assets, indas: t.indas.assets, reference: ref('fs-assets') },
          { at: 'CL', label: 'Total equity and liabilities', local: t.local.equity + t.local.liabilities, indas: t.indas.equity + t.indas.liabilities, reference: ref('fs-eql') },
        ]} />
    </div>
  )
}
