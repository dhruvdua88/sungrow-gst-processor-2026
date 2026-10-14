/** Public entry point of the pure engine (no React, no DOM, no file I/O). */
import { buildBridge, type PeriodBridge } from './bridge'
import { joinBooks, type JoinResult } from './join'
import type { CopyModel, PeriodKey, SignedTB } from './types'

export * from './types'
export * from './money'
export * from './captions'
export * from './join'
export * from './bridge'
export { parseSignedTB } from './parseSigned'
export { parseCopy, parseStatement } from './parseCopy'

export interface Analysis {
  signed: SignedTB
  copy: CopyModel
  join: JoinResult
  bridges: Record<PeriodKey, PeriodBridge>
  warnings: string[]
}

export function analyse(signed: SignedTB, copy: CopyModel): Analysis {
  const join = joinBooks(signed, copy)
  const bridges: Record<PeriodKey, PeriodBridge> = {}
  for (const p of join.periods) bridges[p.key] = buildBridge(join, copy, p.key)
  return { signed, copy, join, bridges, warnings: [...signed.warnings, ...copy.warnings, ...join.warnings] }
}
export { buildFlow, type FlowGraph, type FlowNode, type FlowLink } from './flow'
