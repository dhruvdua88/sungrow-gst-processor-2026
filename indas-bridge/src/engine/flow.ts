/**
 * Flow graph for the Sankey view: signed ledgers → local groups → (+ Ind AS JV) → Ind AS captions →
 * statement sections. Link values are absolute amounts (paise); `dr` tells whether the net flow is a
 * debit. Pure: no layout, no rendering.
 */
import { CAPTION_BY_ID, SECTION_LABEL } from './captions'
import type { JoinResult } from './join'
import type { Paise } from './money'
import type { PeriodKey, StatementKind } from './types'

export type FlowColumn = 0 | 1 | 2 | 3
export interface FlowNode { id: string; label: string; column: FlowColumn; kind: 'ledger' | 'others' | 'primary' | 'face' | 'jv' | 'caption' | 'section'; ledgerKeys: string[]; net: Paise }
export interface FlowLink { source: string; target: string; value: Paise; dr: boolean; ledgerKeys: string[] }
export interface FlowGraph { nodes: FlowNode[]; links: FlowLink[] }

export interface FlowOptions {
  statement: 'ALL' | StatementKind
  /** When set, only ledgers flowing through this node are shown, with a ledger column on the left. */
  focus?: string | null
  topLedgers?: number
}

export function buildFlow(join: JoinResult, k: PeriodKey, opts: FlowOptions): FlowGraph {
  const nodes = new Map<string, FlowNode>()
  const links = new Map<string, FlowLink & { net: Paise }>()
  const node = (id: string, label: string, column: FlowColumn, kind: FlowNode['kind']) => {
    let n = nodes.get(id)
    if (!n) { n = { id, label, column, kind, ledgerKeys: [], net: 0 }; nodes.set(id, n) }
    return n
  }
  const link = (source: string, target: string, amount: Paise, ledgerKey?: string) => {
    if (amount === 0) return
    const id = `${source}→${target}`
    let l = links.get(id)
    if (!l) { l = { source, target, value: 0, dr: true, ledgerKeys: [], net: 0 }; links.set(id, l) }
    l.value += Math.abs(amount)
    l.net += amount
    if (ledgerKey && !l.ledgerKeys.includes(ledgerKey)) l.ledgerKeys.push(ledgerKey)
  }
  const wanted = (captionId: string) => {
    const cap = CAPTION_BY_ID.get(captionId)
    return opts.statement === 'ALL' || cap?.statement === opts.statement
  }

  // Ledger paths (one path per allocation).
  type Path = { key: string; name: string; amount: Paise; ids: string[]; labels: string[] }
  const paths: Path[] = []
  for (const l of join.ledgers) {
    for (const a of l.alloc[k] ?? []) {
      if (a.amount === 0 || !wanted(a.captionId)) continue
      const cap = CAPTION_BY_ID.get(a.captionId)!
      const primary = l.primaryGroup || '(no primary group)'
      const face = l.face[k] || '(no face grouping)'
      paths.push({
        key: l.key,
        name: l.name,
        amount: a.amount,
        ids: [`p:${primary}`, `f:${face}`, `c:${a.captionId}`, `s:${cap.section}:${cap.statement}`],
        labels: [primary, face, cap.label, `${SECTION_LABEL[cap.section]}${cap.section === 'UNMAPPED' ? ` (${cap.statement})` : ''}`],
      })
    }
  }
  const focus = opts.focus ?? null
  const inFocus = focus ? paths.filter((p) => p.ids.includes(focus) || `l:${p.key}` === focus) : paths

  // Ledger column in focus mode: top N by size, rest grouped.
  const ledgerCol = new Map<string, string>()
  if (focus) {
    const byLedger = new Map<string, Paise>()
    for (const p of inFocus) byLedger.set(p.key, (byLedger.get(p.key) ?? 0) + Math.abs(p.amount))
    const top = [...byLedger].sort((a, b) => b[1] - a[1]).slice(0, opts.topLedgers ?? 14).map(([key]) => key)
    for (const p of inFocus) ledgerCol.set(p.key, top.includes(p.key) ? `l:${p.key}` : 'l:__others')
    const others = byLedger.size - top.length
    for (const p of inFocus) {
      const id = ledgerCol.get(p.key)!
      node(id, id === 'l:__others' ? `${others} other ledgers` : p.name, 0, id === 'l:__others' ? 'others' : 'ledger')
    }
  }
  const colOf = (i: number): FlowColumn => i as FlowColumn
  for (const p of inFocus) {
    const chain = focus ? [ledgerCol.get(p.key)!, ...p.ids.slice(1)] : p.ids
    const labels = focus ? ['', ...p.labels.slice(1)] : p.labels
    chain.forEach((id, i) => {
      const n = nodes.get(id) ?? node(id, labels[i], colOf(i), (['primary', 'face', 'caption', 'section'] as const)[i] ?? 'primary')
      if (!n.ledgerKeys.includes(p.key)) n.ledgerKeys.push(p.key)
      n.net += i === chain.length - 1 ? p.amount : 0
    })
    for (let i = 0; i < chain.length - 1; i++) link(chain[i], chain[i + 1], p.amount, p.key)
  }

  // Ind AS JV layer feeds the captions directly.
  const captionsShown = new Set([...nodes.values()].filter((n) => n.kind === 'caption').map((n) => n.id))
  for (const a of join.adjustments) {
    const v = a.row.jv[k] ?? 0
    if (v === 0 || !wanted(a.captionId)) continue
    const cid = `c:${a.captionId}`
    if (focus && !captionsShown.has(cid)) continue
    const cap = CAPTION_BY_ID.get(a.captionId)!
    node('jv', 'Ind AS JV adjustments', 1, 'jv')
    node(cid, cap.label, 2, 'caption')
    const sid = `s:${cap.section}:${cap.statement}`
    node(sid, SECTION_LABEL[cap.section], 3, 'section')
    link('jv', cid, v)
    link(cid, sid, v)
  }
  const out: FlowLink[] = [...links.values()].map(({ net, ...l }) => ({ ...l, dr: net >= 0 }))
  const used = new Set(out.flatMap((l) => [l.source, l.target]))
  return { nodes: [...nodes.values()].filter((n) => used.has(n.id)), links: out }
}
