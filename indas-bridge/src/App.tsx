import { lazy, Suspense, useEffect, useState, type DragEvent } from 'react'
import { useStore, type Tab } from './store'
import { Kpis } from './ui/Kpis'
import { LedgerDrawer } from './ui/LedgerDrawer'
import type { Unit } from './engine'

const FlowView = lazy(() => import('./ui/FlowView'))
const PLBridge = lazy(() => import('./ui/PLBridge'))
const BSBridge = lazy(() => import('./ui/BSBridge'))
const LedgerExplorer = lazy(() => import('./ui/LedgerExplorer'))
const Differences = lazy(() => import('./ui/Differences'))
const Entries = lazy(() => import('./ui/Entries'))

export default function App() {
  const theme = useStore((s) => s.theme)
  const status = useStore((s) => s.status)
  useEffect(() => {
    document.documentElement.dataset.theme = theme
  }, [theme])
  return (
    <div className="app">
      <TopBar />
      {status === 'ready' ? <Dashboard /> : status === 'parsing' ? <Loading /> : <Landing />}
      <footer className="foot">
        <span>Ind AS Bridge · Dhruv Dua &amp; Co · Schedule III, Division I → Division II, Companies Act, 2013</span>
        <span>Your files never leave this browser. No uploads, no analytics, no server.</span>
      </footer>
    </div>
  )
}

function TopBar() {
  const { status, unit, theme, period, analysis, isDemo, set, reset } = useStore()
  const units: [Unit, string][] = [['rupees', '₹'], ['lakhs', 'Lakhs'], ['crores', 'Crores']]
  return (
    <header className="topbar">
      <div className="brand">
        <svg className="brand-mark" viewBox="0 0 32 32" aria-hidden="true">
          <rect width="32" height="32" rx="7" fill="var(--panel-2)" />
          <path d="M5 22c6 0 6-12 11-12s5 12 11 12" fill="none" stroke="var(--accent)" strokeWidth="3" strokeLinecap="round" />
          <path d="M5 25h22" stroke="var(--ink-2)" strokeWidth="1.5" />
        </svg>
        Ind AS Bridge <small>{isDemo ? 'demo data' : 'beta'}</small>
      </div>
      <span className="privacy-pill">Your files never leave this browser</span>
      <div className="spacer" />
      {status === 'ready' && analysis && (
        <>
          <div className="seg" role="group" aria-label="Period">
            {analysis.join.periods.map((p) => (
              <button key={p.key} aria-pressed={period === p.key} onClick={() => set({ period: p.key })}>{p.label.replace(' (transition)', '')}</button>
            ))}
          </div>
          <div className="seg" role="group" aria-label="Display unit">
            {units.map(([u, l]) => (
              <button key={u} aria-pressed={unit === u} onClick={() => set({ unit: u })}>{l}</button>
            ))}
          </div>
          <button className="btn ghost" onClick={reset}>New files</button>
        </>
      )}
      <button className="icon-btn" aria-label={`Switch to ${theme === 'dark' ? 'light' : 'dark'} theme`} onClick={() => set({ theme: theme === 'dark' ? 'light' : 'dark' })}>
        {theme === 'dark' ? '☀' : '☾'}
      </button>
    </header>
  )
}

function Loading() {
  return (
    <div className="loading" role="status" aria-live="polite">
      <div className="spinner" />
      <div>Reading the workbooks in your browser…</div>
      <div className="muted" style={{ fontSize: 13 }}>Large .xlsb files take a few seconds. Nothing is uploaded.</div>
    </div>
  )
}

function DropZone({ which, title, tag, hint, accept }: { which: 'signed' | 'copy'; title: string; tag: string; hint: string; accept: string }) {
  const file = useStore((s) => s.files[which])
  const setFile = useStore((s) => s.setFile)
  const [over, setOver] = useState(false)
  const onDrop = (e: DragEvent) => {
    e.preventDefault()
    setOver(false)
    const f = e.dataTransfer.files?.[0]
    if (f) setFile(which, f)
  }
  return (
    <label
      className={`drop${over ? ' over' : ''}${file ? ' ready' : ''}`}
      onDragOver={(e) => { e.preventDefault(); setOver(true) }}
      onDragLeave={() => setOver(false)}
      onDrop={onDrop}
    >
      <input type="file" accept={accept} aria-label={title} onChange={(e) => setFile(which, e.target.files?.[0] ?? null)} />
      <span className="tag">{tag}</span>
      <h3>{title}</h3>
      <p>{hint}</p>
      <span style={{ flex: 1 }} />
      {file ? (
        <span className="chip ok"><span className="file">✓ {file.name}</span> · {(file.size / 1024 / 1024).toFixed(1)} MB</span>
      ) : (
        <span className="chip">Drop file or click to choose</span>
      )}
    </label>
  )
}

function Landing() {
  const { files, run, loadDemo, error } = useStore()
  const ready = !!files.signed && !!files.copy
  return (
    <main>
      <section className="hero">
        <div>
          <div className="eyebrow">Local GAAP → Ind AS · first-time adoption</div>
          <h1>Trace every ledger from the <em>signed books</em> to the Ind AS statements.</h1>
          <p className="lede">
            Drop the signed Schedule III workbook and the Ind AS working file. Ind AS Bridge joins them ledger by ledger,
            shows every reclassification and adjustment, and checks that it all ties back to the face statements.
          </p>
          <div className="drops">
            <DropZone which="signed" tag="1 · Signed books" title="Local GAAP workbook" hint="The .xlsb from the Schedule 3 Automation Tool (sheet T_B)." accept=".xlsb,.xlsx,.xlsm" />
            <DropZone which="copy" tag="2 · Ind AS working" title="Ind AS workbook" hint="The .xlsx with the TB sheet: Local GAAP | Ind AS JV | Ind AS." accept=".xlsx,.xlsm,.xlsb" />
          </div>
          <div className="actions">
            <button className="btn primary" disabled={!ready} onClick={run}>Build the bridge →</button>
            <button className="btn" onClick={loadDemo}>Load demo data</button>
            <span className="muted" style={{ fontSize: 13 }}>Demo uses invented ledgers and amounts.</span>
          </div>
          {error && <div className="banner error" role="alert">{error}</div>}
        </div>
        <HeroArt />
      </section>
      <section className="how" aria-label="How it works">
        <div><b>Private by design</b><p>Files are read by your browser with SheetJS inside a Web Worker. Nothing is sent anywhere; close the tab and it is gone.</p></div>
        <div><b>Signed books win</b><p>Every bridge starts from the signed amounts. Where the Ind AS file carries a different local figure, you get a red variance chip.</p></div>
        <div><b>Ties, not trust</b><p>Totals are checked against the Ind AS Balance Sheet and Statement of Profit and Loss in the same file.</p></div>
      </section>
    </main>
  )
}

function HeroArt() {
  const left = [40, 92, 144, 196, 248, 300]
  const right = [70, 150, 230, 290]
  const links = [[0, 0], [1, 0], [1, 1], [2, 1], [3, 2], [4, 2], [4, 3], [5, 3]]
  return (
    <svg className="hero-art" viewBox="0 0 520 360" role="img" aria-label="Illustration: ledgers flowing into Ind AS statement lines">
      <defs>
        <linearGradient id="flow" x1="0" x2="1">
          <stop offset="0" stopColor="var(--ink-3)" stopOpacity=".35" />
          <stop offset="1" stopColor="var(--accent)" stopOpacity=".7" />
        </linearGradient>
      </defs>
      {links.map(([a, b], i) => (
        <path key={i} d={`M120 ${left[a]} C 260 ${left[a]}, 260 ${right[b]}, 400 ${right[b]}`} fill="none" stroke="url(#flow)" strokeWidth={6 + (i % 3) * 5} strokeLinecap="round" opacity=".85" />
      ))}
      {left.map((y, i) => (
        <g key={i}>
          <rect x="16" y={y - 13} width="104" height="26" rx="5" fill="var(--panel)" stroke="var(--rule-strong)" />
          <rect x="26" y={y - 3} width={40 + (i * 13) % 50} height="6" rx="3" fill="var(--ink-3)" opacity=".6" />
        </g>
      ))}
      <rect x="250" y="160" width="20" height="40" rx="4" fill="var(--warn)" opacity=".85" />
      <text x="260" y="222" textAnchor="middle" fontSize="11">JV</text>
      {right.map((y, i) => (
        <g key={i}>
          <rect x="400" y={y - 17} width="104" height="34" rx="6" fill="var(--panel-2)" stroke="var(--accent)" />
          <rect x="412" y={y - 3} width={50 + (i * 17) % 30} height="6" rx="3" fill="var(--accent)" opacity=".8" />
        </g>
      ))}
      <text x="68" y="340" textAnchor="middle" fontSize="12" fontFamily="var(--font-mono)">SIGNED BOOKS</text>
      <text x="452" y="340" textAnchor="middle" fontSize="12" fontFamily="var(--font-mono)">IND AS</text>
    </svg>
  )
}

const TABS: [Tab, string][] = [
  ['flow', 'Flow'],
  ['pl', 'P&L bridge'],
  ['bs', 'BS bridge'],
  ['ledgers', 'Ledger explorer'],
  ['diffs', 'Differences'],
  ['entries', 'Ind AS entries'],
]

function Dashboard() {
  const { analysis, tab, set, isDemo, period } = useStore()
  if (!analysis || period === null) return null
  const diffCount = analysis.join.ledgers.filter((l) => l.status !== 'matched' || analysis.join.periods.some((p) => Math.abs(l.diff[p.key]) > 100 || l.regroup[p.key])).length
  const counts: Partial<Record<Tab, number>> = { ledgers: analysis.join.ledgers.length, diffs: diffCount, entries: analysis.copy.entries.length || analysis.join.adjustments.length }
  return (
    <main>
      {isDemo && <div className="banner">Demo mode: every company, ledger and amount on this page is invented. Load your own files with “New files”.</div>}
      {analysis.warnings.length > 0 && (
        <details className="banner" style={{ borderStyle: 'solid' }}>
          <summary>{analysis.warnings.length} note{analysis.warnings.length > 1 ? 's' : ''} from reading the files</summary>
          <ul>{analysis.warnings.map((w, i) => <li key={i}>{w}</li>)}</ul>
        </details>
      )}
      <Kpis />
      <div className="tabs" role="tablist" aria-label="Views">
        {TABS.map(([id, label]) => (
          <button key={id} role="tab" id={`tab-${id}`} aria-selected={tab === id} aria-controls="view" onClick={() => set({ tab: id })}>
            {label}{counts[id] !== undefined && <span className="count">{counts[id]}</span>}
          </button>
        ))}
      </div>
      <section className="view" id="view" role="tabpanel" aria-labelledby={`tab-${tab}`} key={`${tab}-${period}`}>
        <Suspense fallback={<div className="loading"><div className="spinner" /></div>}>
          {tab === 'flow' && <FlowView />}
          {tab === 'pl' && <PLBridge />}
          {tab === 'bs' && <BSBridge />}
          {tab === 'ledgers' && <LedgerExplorer />}
          {tab === 'diffs' && <Differences />}
          {tab === 'entries' && <Entries />}
        </Suspense>
      </section>
      <LedgerDrawer />
    </main>
  )
}
