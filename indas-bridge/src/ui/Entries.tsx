import { Fragment } from 'react'
import { CAPTION_BY_ID, type PeriodKey } from '../engine'
import { useStore } from '../store'
import { csvAmount, downloadCsv, ExportButton, UNIT_LABEL, useFmt } from './common'

export default function Entries() {
  const { analysis, unit } = useStore()
  const fmt = useFmt()
  if (!analysis) return null
  const periods = analysis.join.periods
  const byRow = new Map(analysis.join.adjustments.map((a) => [a.row.sheetRow, a]))
  // Prefer the workbook's voucher register (narrative); amounts always come from the TB rows.
  const items = analysis.copy.entries.length
    ? analysis.copy.entries.map((e) => {
        const adj = e.tbRow !== null ? byRow.get(e.tbRow) : undefined
        return { key: `${e.ref}-${e.tbRow}`, ref: e.ref, section: e.section || 'Vouchers', layer: e.layer, narration: e.narration, line: e.statementLine, reference: e.reference, why: e.why, tbRow: e.tbRow, caption: adj ? CAPTION_BY_ID.get(adj.captionId)?.label : undefined, amounts: (k: PeriodKey) => (adj ? adj.row.jv[k] ?? 0 : e.amounts[k] ?? 0), fromTB: !!adj }
      })
    : analysis.join.adjustments.map((a) => ({ key: String(a.row.sheetRow), ref: a.row.ref, section: a.standard, layer: a.row.layer, narration: a.row.ledger, line: a.row.lineItem, reference: a.standard, why: '', tbRow: a.row.sheetRow, caption: CAPTION_BY_ID.get(a.captionId)?.label, amounts: (k: PeriodKey) => a.row.jv[k] ?? 0, fromTB: true }))
  const sections = [...new Set(items.map((i) => i.section))]
  const exportCsv = () =>
    downloadCsv('ind-as-entries.csv', ['Section', 'Ref', 'TB row', 'Layer', 'Narration', 'Statement line', ...periods.map((p) => `${p.label} (${UNIT_LABEL[unit]}, Dr+ / Cr−)`), 'Ind AS reference', 'Why'],
      items.map((i) => [i.section, i.ref, i.tbRow ?? '', i.layer, i.narration, i.line, ...periods.map((p) => csvAmount(i.amounts(p.key), unit)), i.reference, i.why]))
  return (
    <div className="panel">
      <div className="panel-head">
        <div><h2>Ind AS adjustment entries</h2><p>{items.length} entries · layer A = balance sheet only, layer P = profit and loss · amounts from the TB (Dr +, Cr −) in {UNIT_LABEL[unit]}</p></div>
        <div className="spacer" />
        <ExportButton onClick={exportCsv} />
      </div>
      <div style={{ overflowX: 'auto' }}>
        <table className="ledger">
          <thead><tr><th>Ref</th><th>Entry</th><th>Layer</th>{periods.map((p) => <th key={p.key} className="r">{p.label.replace(' (transition)', '')}</th>)}<th>Standard</th></tr></thead>
          <tbody>
            {sections.map((s) => (
              <Fragment key={s}>
                <tr className="section"><td colSpan={4 + periods.length}>{s}</td></tr>
                {items.filter((i) => i.section === s).map((i) => (
                  <tr key={i.key}>
                    <td className="num muted">{i.ref}{i.tbRow ? <div style={{ fontSize: 11 }}>TB {i.tbRow}</div> : null}</td>
                    <td style={{ maxWidth: 520 }}>
                      <div>{i.narration}</div>
                      <div className="muted" style={{ fontSize: 12.5 }}>{i.caption ? `${i.caption} › ` : ''}{i.line}</div>
                      {i.why && <details style={{ fontSize: 12.5, marginTop: 2 }}><summary className="muted" style={{ cursor: 'pointer' }}>Why</summary>{i.why}</details>}
                    </td>
                    <td><span className={`chip ${i.layer === 'P' ? 'accent' : ''}`}>{i.layer}</span></td>
                    {periods.map((p) => <td key={p.key} className={`r num ${i.amounts(p.key) ? '' : 'zero'}`}>{fmt(i.amounts(p.key))}</td>)}
                    <td className="muted" style={{ fontSize: 12.5 }}>{i.reference}</td>
                  </tr>
                ))}
              </Fragment>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}
