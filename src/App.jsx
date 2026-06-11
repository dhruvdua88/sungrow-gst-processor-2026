import { useMemo, useRef, useState } from "react";
import { buildWorkbook } from "./engine/index.js";
import DaybookAnalysis from "./daybook/DaybookAnalysis.jsx";
import Gstr1Processor from "./gstr1/Gstr1Processor.jsx";
import SalaryOA from "./salaryoa/SalaryOA.jsx";
import "./App.css";

const fmt2 = (n) =>
  Number(n || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function fmtCell(v) {
  if (typeof v === "number") return fmt2(v);
  return v === null || v === undefined ? "" : String(v);
}

// Click a header to sort; numeric-aware; second click flips direction.
function SortableTable({ title, headers, rows, empty, warnRows }) {
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState(1);

  const sorted = useMemo(() => {
    if (sortCol === null) return rows;
    return [...rows].sort((a, b) => {
      const x = a[sortCol];
      const y = b[sortCol];
      const nx = typeof x === "number" ? x : parseFloat(String(x).replace(/,/g, ""));
      const ny = typeof y === "number" ? y : parseFloat(String(y).replace(/,/g, ""));
      if (!Number.isNaN(nx) && !Number.isNaN(ny)) return (nx - ny) * sortDir;
      return String(x ?? "").localeCompare(String(y ?? "")) * sortDir;
    });
  }, [rows, sortCol, sortDir]);

  const onSort = (idx) => {
    if (sortCol === idx) setSortDir((d) => -d);
    else {
      setSortCol(idx);
      setSortDir(1);
    }
  };

  return (
    <div className="preview-block">
      <h3 className={warnRows && rows.length ? "block-title warn-title" : "block-title"}>
        {title}
        <span className="row-count">{rows.length ? ` — ${rows.length} row${rows.length > 1 ? "s" : ""}` : ""}</span>
      </h3>
      {rows.length === 0 ? (
        <p className="empty-msg">{empty}</p>
      ) : (
        <div className="table-wrap">
          <table className="preview-table">
            <thead>
              <tr>
                {headers.map((h, i) => (
                  <th key={h} onClick={() => onSort(i)} title="Click to sort">
                    {h}
                    {sortCol === i ? (sortDir === 1 ? " ▲" : " ▼") : ""}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sorted.map((row, ri) => (
                <tr key={ri} className={warnRows ? "warn-row" : ""}>
                  {row.map((cell, ci) => (
                    <td key={ci} className={typeof cell === "number" ? "num" : ""}>
                      {fmtCell(cell)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function FileRow({ label, required, accept, multiple, files, onFiles, hint, link }) {
  const inputRef = useRef(null);
  const names = files.map((f) => f.name);
  return (
    <div className="path-row">
      <label className="path-label">
        {label}
        {required ? <span className="req">*</span> : null}
        {link ? (
          <a className="row-link" href={link.href} target="_blank" rel="noreferrer">
            {link.text}
          </a>
        ) : null}
      </label>
      <div
        className={`path-entry ${names.length ? "filled" : ""}`}
        onClick={() => inputRef.current?.click()}
        title={names.join(", ")}
      >
        {names.length
          ? names.length > 2
            ? `${names.slice(0, 2).join(", ")} +${names.length - 2} more`
            : names.join(", ")
          : hint || "No file selected"}
      </div>
      <button type="button" className="browse-btn" onClick={() => inputRef.current?.click()}>
        Browse
      </button>
      {files.length > 0 && (
        <button
          type="button"
          className="clear-btn"
          onClick={() => {
            onFiles([]);
            if (inputRef.current) inputRef.current.value = "";
          }}
          title="Clear selection"
        >
          ×
        </button>
      )}
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        multiple={multiple}
        style={{ display: "none" }}
        onChange={(e) => onFiles(Array.from(e.target.files || []))}
      />
    </div>
  );
}

function downloadWorkbook(res) {
  const blob = new Blob([res.buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = res.output;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

export default function App() {
  const [trackerFiles, setTrackerFiles] = useState([]);
  const [gstFiles, setGstFiles] = useState([]);
  const [jsonFiles, setJsonFiles] = useState([]);
  const [priorFiles, setPriorFiles] = useState([]);
  const [boeFiles, setBoeFiles] = useState([]);
  const [isdFiles, setIsdFiles] = useState([]);
  const [period, setPeriod] = useState("April");
  const [fy, setFy] = useState("2026-27");
  const [log, setLog] = useState(["Ready. Select files and create the focused workbook."]);
  const [running, setRunning] = useState(false);
  const [result, setResult] = useState(null);
  const [mode, setMode] = useState("gst"); // "gst" (existing) | "daybook" (new module)

  const appendLog = (msg) => setLog((prev) => [...prev, msg]);

  const validate = () => {
    const required = [
      ["SIPL tracker/template", trackerFiles],
      ["Current GST workbook", gstFiles],
      ["Current GSTR-2B JSON", jsonFiles],
    ];
    for (const [label, files] of required) {
      if (!files.length) {
        alert(`Missing input\n\nSelect ${label}.`);
        return false;
      }
    }
    if (!period.trim() || !fy.trim()) {
      alert("Missing period\n\nEnter period/month and FY.");
      return false;
    }
    return true;
  };

  const startRun = async () => {
    if (!validate()) return;
    setRunning(true);
    setResult(null);
    appendLog("");
    appendLog("Starting workbook creation...");
    try {
      const [trackerBuffer, gstWorkbookBuffer, r2bText] = await Promise.all([
        trackerFiles[0].arrayBuffer(),
        gstFiles[0].arrayBuffer(),
        jsonFiles[0].text(),
      ]);
      const r2bJson = JSON.parse(r2bText);
      const priorData = await Promise.all(
        priorFiles.map(async (f) => ({ name: f.name, buffer: await f.arrayBuffer() }))
      );
      const boeData = await Promise.all(
        boeFiles.map(async (f) => ({
          name: f.name,
          relativePath: f.webkitRelativePath || f.name,
          data: await f.arrayBuffer(),
        }))
      );
      let isdJson = null;
      if (isdFiles.length) isdJson = JSON.parse(await isdFiles[0].text());

      const res = await buildWorkbook({
        trackerBuffer,
        gstWorkbookBuffer,
        r2bJson,
        priorFiles: priorData,
        boeFiles: boeData,
        isdJson,
        periodLabel: period.trim(),
        fyLabel: fy.trim(),
        onProgress: appendLog,
      });

      // Download the focused workbook
      downloadWorkbook(res);

      // Same success log lines as the desktop tool
      appendLog(`Created: ${res.output}`);
      appendLog(`Sanitized ITC rows: ${res.invoice_count}`);
      appendLog(`2B rows: ${res.r2b_count}`);
      appendLog(`Correction flags: ${res.correction_count} | Amount: ${fmt2(res.correction_amount)}`);
      appendLog(`Ledger mismatch invoices: ${Math.trunc(res.ledger_mismatch_count)}`);
      appendLog(
        `Prior aliases used/available: ${res.prior_alias_count} | Conflicts: ${res.prior_conflict_count}`
      );
      appendLog(`BOE PDFs parsed/matched: ${res.boe_count} | ${res.boe_matched_count}`);
      appendLog(
        `  BOE validation: GSTIN mismatch = ${res.boe_gstin_mismatch_count}, Not in 2B = ${res.boe_not_in_2b_count}`
      );
      const insp = res.boe_inspection || {};
      if (insp.pdfs_found) {
        appendLog(
          `  BOE input scan: ${insp.pdfs_found} PDF(s) found, ${insp.parsed} valid BOEs, ` +
            `${insp.format_mismatch} non-BOE (invoices/packing lists), ` +
            `${insp.text_extraction_failed} unreadable, ${insp.deduped} duplicates.`
        );
        for (const [path, reason] of (insp.skipped_samples || []).slice(0, 3)) {
          appendLog(`    skipped: ${path.split("!").pop()} — ${reason.slice(0, 80)}`);
        }
      }
      appendLog(`Eligible ISD rows/tax: ${res.isd_eligible_count} | ${fmt2(res.isd_eligible_tax)}`);
      appendLog(`ISD bill-wise rows/tax: ${res.isd_bill_count} | ${fmt2(res.isd_bill_tax)}`);
      appendLog(
        `ISD eligibility classifier: ${res.isd_eligibility_method || "n/a"} ` +
          `(${res.isd_eligible_bill_count} bills, ₹${fmt2(res.isd_eligible_bill_tax)})`
      );
      appendLog(
        `ISD supplier GSTINs in Sanitized ITC / absent: ` +
          `${res.isd_gstin_present_count} | ${res.isd_gstin_absent_count} ` +
          `(absent tax ${fmt2(res.isd_gstin_absent_tax)})`
      );
      setResult(res);
    } catch (exc) {
      console.error(exc);
      appendLog(`Failed: ${exc.message || exc}`);
      alert(`Run failed\n\n${exc.message || exc}`);
    } finally {
      setRunning(false);
    }
  };

  return (
    <div className="app">
      <header className="header">
        <div className="header-inner">
          <div className="brand-mark">SG</div>
          <div>
            <h1>Sungrow GST Processor 2026</h1>
            <p className="header-title">SIPL GST Monthly ITC and Ledger Tie-Out</p>
            <p className="header-sub">
              Create the focused workbook from current GST files. Prior files, BOE support, and ISD
              JSON are optional.
            </p>
          </div>
        </div>
      </header>

      <nav className="mode-tabs">
        {[["gst", "GST Processor"], ["daybook", "Daybook Analysis"], ["gstr1", "GSTR-1 Processor"], ["salaryoa", "Salary OA / APAC"]].map(([id, label]) => (
          <button
            key={id}
            className={`mode-tab ${mode === id ? "active" : ""}`}
            onClick={() => setMode(id)}
          >
            {label}
          </button>
        ))}
      </nav>

      <main className="main">
        {mode === "salaryoa" ? (
          <SalaryOA />
        ) : mode === "gstr1" ? (
          <Gstr1Processor />
        ) : mode === "daybook" ? (
          <DaybookAnalysis />
        ) : (
          <>
        <section className="card form-card">
          <FileRow
            label="SIPL tracker/template"
            required
            accept=".xlsx"
            files={trackerFiles}
            onFiles={setTrackerFiles}
            hint="Select tracker .xlsx — download from the template link above if you don't have it"
            link={{
              href: "https://docs.google.com/spreadsheets/d/1JjYLfd0oO7Q-lKfhH_EtVbCJ6kqA7XY9iL8cVN4JzBo/export?format=xlsx",
              text: "⬇ Download template (.xlsx)",
            }}
          />
          <FileRow
            label="Current GST workbook"
            required
            accept=".xlsx"
            files={gstFiles}
            onFiles={setGstFiles}
            hint="Select workbook with PO / Non PO / GL Summary"
          />
          <FileRow
            label="Current GSTR-2B JSON"
            required
            accept=".json"
            files={jsonFiles}
            onFiles={setJsonFiles}
            hint="Select returns_R2B_*.json"
          />
          <FileRow
            label="Prior focused output files (optional)"
            accept=".xlsx"
            multiple
            files={priorFiles}
            onFiles={setPriorFiles}
            hint="Select prior *Focused_Output*.xlsx files"
          />
          <FileRow
            label="BOE PDF zip/files (optional)"
            accept=".zip,.pdf"
            multiple
            files={boeFiles}
            onFiles={setBoeFiles}
            hint="Select BOE .zip and/or .pdf files"
          />
          <FileRow
            label="ISD JSON (optional)"
            accept=".json"
            files={isdFiles}
            onFiles={setIsdFiles}
            hint="Select GSTR-6 ISD .json"
          />

          <div className="meta-row">
            <label className="path-label">Period/month</label>
            <input className="meta-input" value={period} onChange={(e) => setPeriod(e.target.value)} />
            <label className="fy-label">FY</label>
            <input className="meta-input" value={fy} onChange={(e) => setFy(e.target.value)} />
          </div>

          <p className="note">
            <strong>How to use:</strong> 1) Download the SIPL tracker template from the{" "}
            <a
              href="https://docs.google.com/spreadsheets/d/1JjYLfd0oO7Q-lKfhH_EtVbCJ6kqA7XY9iL8cVN4JzBo/edit?gid=2126389207#gid=2126389207"
              target="_blank"
              rel="noreferrer"
            >
              Google Sheets master
            </a>{" "}
            (File → Download → .xlsx, or use the ⬇ link above) and select it here. 2) Select the
            current GST workbook (PO / Non PO / GL Summary tabs) and the GSTR-2B JSON from the GST
            portal. 3) Optional: prior focused outputs, BOE ZIP/PDFs, ISD JSON. BOE input can be a
            ZIP, a PDF, or multiple files (nested ZIPs are scanned). ISD JSON adds bill-wise b2b
            detail and uses isd.elglst only as the eligible control total. The focused workbook
            downloads automatically when ready.
          </p>

          <button className="run-btn" onClick={startRun} disabled={running}>
            {running ? "Creating Workbook…" : "Create Focused Workbook"}
          </button>
        </section>

        {result && (
          <section className="card stats-card">
            <div className="preview-head">
              <h2>Dashboard Preview</h2>
              <button className="dl-again-btn" onClick={() => downloadWorkbook(result)}>
                ⬇ Download workbook again
              </button>
            </div>
            <div className="stats-grid">
              {result.preview.cards.map((c) => (
                <div key={c.label} className={`stat ${c.warn ? "warn" : ""}`}>
                  <span className="stat-label">{c.label}</span>
                  <span className="stat-value">{c.count ? c.value : `₹${fmt2(c.value)}`}</span>
                </div>
              ))}
            </div>
            <SortableTable {...result.preview.ledgerMismatches} warnRows />
            <SortableTable {...result.preview.unresolvedB2b} warnRows />
            <SortableTable {...result.preview.corrections} warnRows />
            <SortableTable {...result.preview.ledgerSummary} />
          </section>
        )}

        <section className="card log-card">
          <h2>Status</h2>
          <pre className="log">{log.join("\n")}</pre>
        </section>
          </>
        )}
      </main>

      <footer className="footer">
        All processing runs locally in your browser — no file ever leaves this machine.
      </footer>
    </div>
  );
}
