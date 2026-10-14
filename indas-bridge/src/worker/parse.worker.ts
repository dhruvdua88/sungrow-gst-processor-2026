/// <reference lib="webworker" />
/**
 * Parses the two workbooks off the main thread. Files arrive as ArrayBuffers from the page and
 * never leave the browser; nothing here performs network I/O.
 */
import type { CopyModel, SignedTB } from '../engine/types'
import { loadCopy, loadSigned } from './load'

export type ParseRequest = { signed: ArrayBuffer; copy: ArrayBuffer }
export type ParseResponse = { ok: true; signed: SignedTB; copy: CopyModel } | { ok: false; error: string; which: 'signed' | 'copy' }

self.onmessage = (e: MessageEvent<ParseRequest>) => {
  let which: 'signed' | 'copy' = 'signed'
  try {
    const signed = loadSigned(e.data.signed)
    which = 'copy'
    const copy = loadCopy(e.data.copy)
    ;(self as unknown as Worker).postMessage({ ok: true, signed, copy } satisfies ParseResponse)
  } catch (err) {
    ;(self as unknown as Worker).postMessage({ ok: false, which, error: err instanceof Error ? err.message : String(err) } satisfies ParseResponse)
  }
}
