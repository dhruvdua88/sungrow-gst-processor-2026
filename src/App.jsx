import { useRef, useState } from "react";
import { buildWorkbook } from "./engine/index.js";
import "./App.css";

const fmt2 = (n) =>
  Number(n || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function FileRow({ label, required, accept, multiple, files, onFiles, hint }) {
  const inputRef = useRef(null);
  const names = files.map((f) => f.name);
  return (
    <div className="path-row">
      <label className="path-label">
        {label}
        {required ? <span className="req">*</span> : null}
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

      <main className="main">
        <section className="card form-card">
          <FileRow
            label="SIPL tracker/template"
            required
            accept=".xlsx"
            files={trackerFiles}
            onFiles={setTrackerFiles}
            hint="Select tracker .xlsx"
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
            BOE input can be a ZIP, a PDF, or multiple files (nested ZIPs are scanned). ISD JSON adds
            bill-wise b2b detail and uses isd.elglst only as the eligible control total. The focused
            workbook downloads automatically when ready.
          </p>

          <button className="run-btn" onClick={startRun} disabled={running}>
            {running ? "Creating Workbook…" : "Create Focused Workbook"}
          </button>
        </section>

        {result && (
          <section className="card stats-card">
            <h2>Run Summary</h2>
            <div className="stats-grid">
              <div className="stat">
                <span className="stat-label">Sanitized ITC Tax</span>
                <span className="stat-value">₹{fmt2(result.books_total_tax)}</span>
              </div>
              <div className="stat">
                <span className="stat-label">2B Available Tax</span>
                <span className="stat-value">₹{fmt2(result.r2b_total_tax)}</span>
              </div>
              <div className={`stat ${result.ledger_itc_rec_difference ? "warn" : ""}`}>
                <span className="stat-label">ITC − REC Ledger Diff</span>
                <span className="stat-value">₹{fmt2(result.ledger_itc_rec_difference)}</span>
              </div>
              <div className={`stat ${result.correction_count ? "warn" : ""}`}>
                <span className="stat-label">Correction Flags</span>
                <span className="stat-value">{result.correction_count}</span>
              </div>
              <div className="stat">
                <span className="stat-label">Sanitized ITC Rows</span>
                <span className="stat-value">{result.invoice_count}</span>
              </div>
              <div className="stat">
                <span className="stat-label">2B Rows</span>
                <span className="stat-value">{result.r2b_count}</span>
              </div>
              <div className="stat">
                <span className="stat-label">BOE Parsed / Matched</span>
                <span className="stat-value">
                  {result.boe_count} / {result.boe_matched_count}
                </span>
              </div>
              <div className="stat">
                <span className="stat-label">ISD Bills / Tax</span>
                <span className="stat-value">
                  {result.isd_bill_count} / ₹{fmt2(result.isd_bill_tax)}
                </span>
              </div>
            </div>
          </section>
        )}

        <section className="card log-card">
          <h2>Status</h2>
          <pre className="log">{log.join("\n")}</pre>
        </section>
      </main>

      <footer className="footer">
        All processing runs locally in your browser — no file ever leaves this machine.
      </footer>
    </div>
  );
}
