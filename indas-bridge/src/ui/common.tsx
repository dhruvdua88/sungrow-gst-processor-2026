import { useCallback, useState, type ReactNode } from 'react'
import { formatMoney, toUnitString, type Paise, type TieStatus, type Unit } from '../engine'
import { useStore } from '../store'

export const UNIT_LABEL: Record<Unit, string> = { rupees: '₹', lakhs: '₹ lakhs', crores: '₹ crores' }

/** Formatter bound to the selected display unit. */
export function useFmt() {
  const unit = useStore((s) => s.unit)
  return useCallback((v: Paise | null | undefined, opts: { symbol?: boolean; signed?: boolean } = {}) =>
    v === null || v === undefined ? '—' : formatMoney(v, unit, { symbol: opts.symbol ?? false, signed: opts.signed }), [unit])
}

export function Amount({ v, symbol, signed, dim }: { v: Paise | null | undefined; symbol?: boolean; signed?: boolean; dim?: boolean }) {
  const fmt = useFmt()
  const cls = v === 0 || v === null || v === undefined ? 'num zero' : dim && v < 0 ? 'num neg' : 'num'
  return <span className={cls}>{fmt(v, { symbol, signed })}</span>
}

export function TieChip({ status, label }: { status: TieStatus; label?: string }) {
  const map = { tie: ['ok', '✓ Ties'], rounding: ['warn', '≈ Rounding'], diff: ['bad', '✕ Differs'], na: ['', 'n/a'] } as const
  const [cls, text] = map[status]
  return <span className={`chip ${cls}`} title={label}>{label ? `${text} · ${label}` : text}</span>
}

export function VarianceChip({ diff }: { diff: Paise }) {
  const fmt = useFmt()
  if (Math.abs(diff) <= 100) return null
  return <span className="chip bad" title="Signed books minus Ind AS workbook local GAAP">Δ {fmt(diff, { signed: true })}</span>
}

/** Floating tooltip that follows the pointer. */
export function useTooltip() {
  const [tip, setTip] = useState<{ x: number; y: number; content: ReactNode } | null>(null)
  const show = useCallback((e: { clientX: number; clientY: number }, content: ReactNode) => setTip({ x: e.clientX, y: e.clientY, content }), [])
  const hide = useCallback(() => setTip(null), [])
  const node = tip ? (
    <div className="tooltip" role="status" style={{ left: Math.min(tip.x + 14, window.innerWidth - 370), top: tip.y + 14 }}>{tip.content}</div>
  ) : null
  return { show, hide, node }
}

/** Client-side CSV download (no upload). */
export function downloadCsv(filename: string, header: string[], rows: (string | number)[][]) {
  const esc = (v: string | number) => {
    const s = String(v)
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
  }
  const text = [header, ...rows].map((r) => r.map(esc).join(',')).join('\r\n')
  const blob = new Blob(['﻿' + text], { type: 'text/csv;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

export function csvAmount(v: Paise | null | undefined, unit: Unit): string {
  return v === null || v === undefined ? '' : toUnitString(v, unit)
}

export function ExportButton({ onClick }: { onClick: () => void }) {
  return <button className="btn ghost" onClick={onClick} title="Download the current table as CSV (created in your browser)">⤓ CSV</button>
}
