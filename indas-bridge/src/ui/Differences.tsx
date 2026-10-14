import { CAPTION_BY_ID } from '../engine'
import { useStore } from '../store'
import { csvAmount, downloadCsv, ExportButton, UNIT_LABEL, useFmt } from './common'

export default function Differences() {
  const { analysis, unit, set } = useStore()
  const fmt = useFmt()
  if (!analysis) return null
  const periods = analysis.join.periods
  const rows = analysis.join.ledgers
    .filter((l) => l.status !== 'matched' || periods.some((p) => Math.abs(l.diff[p.key]) > 100 || l.regroup[p.key]))
    .sort((a, b) => Math.max(...periods.map((p) => Math.abs(b.diff[p.key]))) - Math.max(...periods.map((p) => Math.abs(a.diff[p.key]))))
  const kind = (l: (typeof rows)[number]) =>
    l.status === 'xlsb-only' ? ['bad', 'Only in signed books'] : l.status === 'copy-only' ? ['bad', 'Only in Ind AS workbook'] : periods.some((p) => Math.abs(l.diff[p.key]) > 100) ? ['bad', 'Amount differs'] : ['warn', 'Regrouped']
  const mapping = (l: (typeof rows)[number]) => {
    const k = periods[0].key
    return (l.alloc[k] ?? []).map((a) => `${CAPTION_BY_ID.get(a.captionId)?.label ?? a.captionId}: ${fmt(a.amount)}`).join(' · ')
  }
  const exportCsv = () =>
    downloadCsv('differences.csv', ['Ledger', 'Issue', ...periods.flatMap((p) => [`${p.label} signed (${UNIT_LABEL[unit]})`, `${p.label} Ind AS file`, `${p.label} difference`]), 'Ind AS mapping used (latest period)'],
      rows.map((l) => [l.name, kind(l)[1], ...periods.flatMap((p) => [csvAmount(l.used[p.key], unit), csvAmount(l.copyLocal[p.key], unit), csvAmount(l.diff[p.key], unit)]), mapping(l)]))
  return (
    <div className="panel">
      <div className="panel-head">
        <div><h2>Differences between the two files</h2><p>Signed books win: the bridge uses the signed amount. Fix these in the Ind AS workbook, or record why they differ. {rows.length} ledger{rows.length === 1 ? '' : 's'}.</p></div>
        <div className="spacer" />
        <ExportButton onClick={exportCsv} />
      </div>
      {rows.length === 0 ? (
        <div className="panel-body"><span className="chip ok">✓ Every ledger matches in every period</span></div>
      ) : (
        <div style={{ overflowX: 'auto' }}>
          <table className="ledger">
            <thead>
              <tr><th>Ledger</th><th>Issue</th>{periods.map((p) => <th key={p.key} className="r">{p.label.replace(' (transition)', '')}<br /><span style={{ textTransform: 'none' }}>signed / Ind AS file / Δ</span></th>)}<th>Mapping used ({periods[0].label})</th></tr>
            </thead>
            <tbody>
              {rows.map((l) => {
                const [cls, text] = kind(l)
                return (
                  <tr key={l.key} className="click" tabIndex={0} onClick={() => set({ selected: l.key })} onKeyDown={(e) => { if (e.key === 'Enter') set({ selected: l.key }) }}>
                    <td style={{ maxWidth: 320 }}>{l.name}</td>
                    <td><span className={`chip ${cls}`}>{text}</span></td>
                    {periods.map((p) => (
                      <td key={p.key} className="r num" style={{ lineHeight: 1.35 }}>
                        {fmt(l.used[p.key])}<br /><span className="muted">{fmt(l.copyLocal[p.key])}</span><br />
                        <span className={Math.abs(l.diff[p.key]) > 100 ? 'delta-neg' : 'zero'}>{fmt(l.diff[p.key], { signed: true })}</span>
                        {l.regroup[p.key] && <><br /><span className="chip warn" style={{ fontSize: 11 }}>regrouped</span></>}
                      </td>
                    ))}
                    <td className="muted" style={{ fontSize: 12.5 }}>{mapping(l)}</td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
