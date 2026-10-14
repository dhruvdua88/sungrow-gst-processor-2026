import { Fragment } from 'react'
import { SECTION_LABEL, type BridgeLine, type SectionId } from '../engine'
import { tieStatus } from '../engine/bridge'
import { useStore } from '../store'
import { csvAmount, downloadCsv, ExportButton, TieChip, UNIT_LABEL, useFmt } from './common'

export interface Subtotal { at: SectionId; label: string; local: number; indas: number; reference?: number | null }

export function BridgeTable({ lines, title, subtitle, subtotals = [], filename }: { lines: BridgeLine[]; title: string; subtitle: string; subtotals?: Subtotal[]; filename: string }) {
  const fmt = useFmt()
  const { unit, set } = useStore()
  const sections = [...new Set(lines.map((l) => l.caption.section))]
  const sum = (sec: SectionId, k: 'local' | 'reclass' | 'adj' | 'indas') => lines.filter((l) => l.caption.section === sec).reduce((a, l) => a + l[k], 0)
  const exportCsv = () =>
    downloadCsv(filename, ['Section', 'Caption', `Local GAAP (${UNIT_LABEL[unit]})`, 'Reclass', 'Ind AS adjustment', 'Ind AS', 'Ind AS workbook face statement', 'Difference'],
      lines.map((l) => [SECTION_LABEL[l.caption.section], l.caption.label, csvAmount(l.local, unit), csvAmount(l.reclass, unit), csvAmount(l.adj, unit), csvAmount(l.indas, unit), csvAmount(l.reference, unit), csvAmount(l.reference === null ? null : l.indas - l.reference, unit)]))
  return (
    <div className="panel">
      <div className="panel-head">
        <div><h2>{title}</h2><p>{subtitle} · amounts in {UNIT_LABEL[unit]} · presentation sign</p></div>
        <div className="spacer" />
        <ExportButton onClick={exportCsv} />
      </div>
      <div style={{ overflowX: 'auto' }}>
        <table className="ledger">
          <thead>
            <tr><th>Caption</th><th className="r">Local GAAP</th><th className="r">Reclass</th><th className="r">Ind AS adj.</th><th className="r">Ind AS</th><th className="r">Face statement</th><th>Tie</th></tr>
          </thead>
          <tbody>
            {sections.map((sec) => (
              <Fragment key={sec}>
                <tr className="section"><td colSpan={7}>{SECTION_LABEL[sec]}</td></tr>
                {lines.filter((l) => l.caption.section === sec).map((l) => {
                  const diff = l.reference === null ? null : l.indas - l.reference
                  return (
                    <tr key={l.caption.id} className="click" tabIndex={0} title="Open these ledgers in the explorer"
                      onClick={() => set({ tab: 'ledgers', captionFilter: l.caption.id })}
                      onKeyDown={(e) => { if (e.key === 'Enter') set({ tab: 'ledgers', captionFilter: l.caption.id }) }}>
                      <td>{l.caption.label}{l.ledgerCount > 0 && <span className="muted" style={{ fontSize: 12 }}> · {l.ledgerCount} ledgers</span>}{l.adjustmentRows.length > 0 && <span className="muted" style={{ fontSize: 12 }}> · {l.adjustmentRows.length} adj.</span>}</td>
                      <td className="r num">{fmt(l.local)}</td>
                      <td className={`r num ${l.reclass ? '' : 'zero'}`}>{fmt(l.reclass, { signed: true })}</td>
                      <td className={`r num ${l.adj ? '' : 'zero'}`} style={l.adj ? { color: 'var(--accent)' } : undefined}>{fmt(l.adj, { signed: true })}</td>
                      <td className="r num" style={{ fontWeight: 600 }}>{fmt(l.indas)}</td>
                      <td className="r num muted">{fmt(l.reference)}</td>
                      <td>{diff === null ? <span className="muted">—</span> : <TieChip status={tieStatus(diff)} label={tieStatus(diff) === 'tie' ? undefined : fmt(diff, { signed: true })} />}</td>
                    </tr>
                  )
                })}
                <tr className="total">
                  <td>Total {SECTION_LABEL[sec].toLowerCase()}</td>
                  <td className="r num">{fmt(sum(sec, 'local'))}</td>
                  <td className="r num">{fmt(sum(sec, 'reclass'), { signed: true })}</td>
                  <td className="r num">{fmt(sum(sec, 'adj'), { signed: true })}</td>
                  <td className="r num">{fmt(sum(sec, 'indas'))}</td>
                  <td /><td />
                </tr>
                {subtotals.filter((s) => s.at === sec).map((s) => (
                  <tr key={s.label} className="total">
                    <td style={{ fontFamily: 'var(--font-display)' }}>{s.label}</td>
                    <td className="r num">{fmt(s.local)}</td>
                    <td className="r num" colSpan={2} />
                    <td className="r num">{fmt(s.indas)}</td>
                    <td className="r num muted">{fmt(s.reference ?? null)}</td>
                    <td>{s.reference === undefined || s.reference === null ? null : <TieChip status={tieStatus(s.indas - s.reference)} />}</td>
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
