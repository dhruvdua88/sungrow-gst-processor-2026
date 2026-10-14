import { create } from 'zustand'
import { analyse, parseCopy, parseSignedTB, type Analysis, type PeriodKey, type Unit } from './engine'
import type { ParseResponse } from './worker/parse.worker'

export type Tab = 'flow' | 'pl' | 'bs' | 'ledgers' | 'diffs' | 'entries'
export type Theme = 'dark' | 'light'

interface State {
  status: 'idle' | 'parsing' | 'ready' | 'error'
  error: string | null
  files: { signed: File | null; copy: File | null }
  analysis: Analysis | null
  isDemo: boolean
  period: PeriodKey | null
  unit: Unit
  tab: Tab
  flowMode: 'sankey' | 'tree'
  selected: string | null
  theme: Theme
  /** Ledger explorer pre-filter set by clicking a caption elsewhere. */
  captionFilter: string | null
  setFile: (which: 'signed' | 'copy', f: File | null) => void
  run: () => Promise<void>
  loadDemo: () => Promise<void>
  reset: () => void
  set: (p: Partial<Pick<State, 'period' | 'unit' | 'tab' | 'flowMode' | 'selected' | 'theme' | 'captionFilter'>>) => void
}

function initialTheme(): Theme {
  try {
    const saved = localStorage.getItem('indas-bridge-theme')
    if (saved === 'light' || saved === 'dark') return saved
  } catch { /* storage unavailable */ }
  return typeof matchMedia === 'function' && matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark'
}

function parseInWorker(signed: ArrayBuffer, copy: ArrayBuffer): Promise<ParseResponse> {
  return new Promise((resolve) => {
    const w = new Worker(new URL('./worker/parse.worker.ts', import.meta.url), { type: 'module' })
    w.onmessage = (e: MessageEvent<ParseResponse>) => { resolve(e.data); w.terminate() }
    w.onerror = (e) => { resolve({ ok: false, which: 'signed', error: e.message || 'The parser stopped unexpectedly.' }); w.terminate() }
    w.postMessage({ signed, copy }, [signed, copy])
  })
}

export const useStore = create<State>((set, get) => ({
  status: 'idle',
  error: null,
  files: { signed: null, copy: null },
  analysis: null,
  isDemo: false,
  period: null,
  unit: 'lakhs',
  tab: 'flow',
  flowMode: 'sankey',
  selected: null,
  theme: initialTheme(),
  captionFilter: null,
  setFile: (which, f) => set((s) => ({ files: { ...s.files, [which]: f }, error: null })),
  run: async () => {
    const { signed, copy } = get().files
    if (!signed || !copy) return
    set({ status: 'parsing', error: null })
    const res = await parseInWorker(await signed.arrayBuffer(), await copy.arrayBuffer())
    if (!res.ok) {
      set({ status: 'error', error: `${res.which === 'signed' ? 'Signed books (xlsb)' : 'Ind AS workbook (xlsx)'}: ${res.error}` })
      return
    }
    try {
      const analysis = analyse(res.signed, res.copy)
      set({ status: 'ready', analysis, isDemo: false, period: analysis.join.periods[0]?.key ?? null, tab: 'flow', selected: null })
    } catch (err) {
      set({ status: 'error', error: err instanceof Error ? err.message : String(err) })
    }
  },
  loadDemo: async () => {
    set({ status: 'parsing', error: null })
    const d = await import('./demo/demo')
    const signed = parseSignedTB(d.demoSignedGrid())
    const copy = parseCopy({ tb: d.demoCopyTBGrid(), bs: d.demoBalanceSheetGrid(), pl: d.demoProfitLossGrid(), entries: d.demoEntriesGrid() })
    const analysis = analyse(signed, copy)
    set({ status: 'ready', analysis, isDemo: true, period: analysis.join.periods[0].key, tab: 'flow', selected: null })
  },
  reset: () => set({ status: 'idle', analysis: null, files: { signed: null, copy: null }, isDemo: false, error: null, selected: null }),
  set: (p) => {
    if (p.theme) {
      try { localStorage.setItem('indas-bridge-theme', p.theme) } catch { /* ignore */ }
    }
    set(p)
  },
}))
