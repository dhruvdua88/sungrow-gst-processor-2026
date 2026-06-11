import { useRef, useState } from "react";
import { build, validate, writeWorkbook, splitExports, makeExtraRow, pyRound, loadSalary } from "./engine.js";
import { isEncrypted, decryptOffice } from "./decrypt.js";
import { buildApac, saveApac, writeManualReport, writeRecoReport } from "./apac.js";
import "./salaryoa.css";

const fmt0 = (n) => Number(n || 0).toLocaleString("en-IN", { maximumFractionDigits: 0 });

function download(buffer, filename) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function FilePick({ label, required, accept, file, onFile, hint }) {
  const inputRef = useRef(null);
  return (
    <div className="soa-file-row">
      <label className="soa-file-label">
        {label}
        {required ? <span className="soa-req">*</span> : null}
      </label>
      <div
        className={`soa-file-entry ${file ? "filled" : ""}`}
        onClick={() => inputRef.current?.click()}
        title={file?.name || ""}
      >
        {file ? file.name : hint || "No file selected"}
      </div>
      <button type="button" className="soa-browse" onClick={() => inputRef.current?.click()}>
        Browse
      </button>
      {file && (
        <button
          type="button"
          className="soa-clear"
          onClick={() => {
            onFile(null);
            if (inputRef.current) inputRef.current.value = "";
          }}
        >
          ×
        </button>
      )}
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        style={{ display: "none" }}
        onChange={(e) => onFile(e.target.files?.[0] || null)}
      />
    </div>
  );
}

export default function SalaryOA() {
  // --- Step 1 inputs
  const [salaryFile, setSalaryFile] = useState(null);
  const [pfFile, setPfFile] = useState(null);
  const [monthLabel, setMonthLabel] = useState("");
  const [loans, setLoans] = useState([]); // {code, amount}
  const [stipends, setStipends] = useState([]); // {name, amount, category}
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [pfAdminOverride, setPfAdminOverride] = useState("");
  const [zeroTds, setZeroTds] = useState("");

  // --- Step 2 inputs
  const [apacFile, setApacFile] = useState(null);
  const [apacPassword, setApacPassword] = useState("Sungrow@123");

  // --- State
  const [running, setRunning] = useState(false);
  const [oaResult, setOaResult] = useState(null); // {run, report, adminTotal, edliTotal, buffers}
  const [apacResult, setApacResult] = useState(null);
  const [log, setLog] = useState(["Ready. Follow steps 1 → 2 → 3 below."]);

  const appendLog = (m) => setLog((prev) => [...prev, m]);

  // Employee dropdown for loan EMIs — parsed straight from the salary register
  // on selection, so codes are pickable before the first generate.
  const [registerEmployees, setRegisterEmployees] = useState([]);
  const onSalaryFile = async (f) => {
    setSalaryFile(f);
    setRegisterEmployees([]);
    if (!f) return;
    try {
      const emps = loadSalary(await f.arrayBuffer());
      setRegisterEmployees(emps);
      appendLog(`Salary register read: ${emps.length} employees (loan dropdown ready).`);
    } catch (exc) {
      appendLog(`Could not pre-read salary register: ${exc.message || exc}`);
    }
  };

  const employeeOptions = (oaResult ? oaResult.run.employees : registerEmployees).map((e) => ({
    code: e.code,
    label: `${e.code} — ${e.name}`,
  }));

  const monthSuffix = monthLabel.trim() ? ` - ${monthLabel.trim()}` : "";

  const runOa = async () => {
    if (!salaryFile || !pfFile) {
      alert("Missing input\n\nSelect both the Salary register and the PF MIS report.");
      return;
    }
    if (!monthLabel.trim()) {
      alert("Missing month label\n\nEnter the month, e.g. \"May 2026\".");
      return;
    }
    setRunning(true);
    setOaResult(null);
    setApacResult(null);
    try {
      appendLog("");
      appendLog(`Reading salary register: ${salaryFile.name}`);
      const [salaryBuf, pfBuf] = await Promise.all([salaryFile.arrayBuffer(), pfFile.arrayBuffer()]);

      const extras = stipends
        .filter((s) => s.name.trim() && Number(s.amount))
        .map((s) =>
          makeExtraRow({
            code: "STIPEND",
            name: `Stipend — ${s.name.trim()}`,
            department: s.category === "Service" ? "SERVICE" : "OPERATIONS",
            netPaid: Number(s.amount),
          })
        );

      const opts = {
        extras,
        monthLabel: monthLabel.trim(),
        pfAdminOverride: pfAdminOverride !== "" && Number(pfAdminOverride) > 0 ? Number(pfAdminOverride) : undefined,
        zeroTdsCodes: zeroTds.split(",").map((c) => c.trim()).filter(Boolean),
      };

      const res = build(salaryBuf, pfBuf, opts);

      // Apply loan EMIs after build (matches the desktop tool's flow)
      let loanTotal = 0;
      for (const ln of loans) {
        if (!ln.code || !Number(ln.amount)) continue;
        const emp = res.run.employees.find((e) => e.code === ln.code);
        if (emp) {
          emp.loanDeduction += Number(ln.amount);
          loanTotal += Number(ln.amount);
        }
      }
      // Loans change Net Paid → re-validate
      const report = validate(res.run, res.pfLookup, res.pfAdminSource);

      appendLog(`Employees loaded: ${res.run.employees.length}`);
      appendLog(`PF MIS codes: ${Object.keys(res.pfLookup).length} | PF Admin: ${fmt0(res.adminTotal)} | EDLI: ${fmt0(res.edliTotal)}`);
      if (loanTotal) appendLog(`Loan EMIs applied: ${fmt0(loanTotal)}`);
      if (extras.length) appendLog(`Stipends added: ${extras.length}`);

      const mainBuffer = await writeWorkbook(res.run, report);
      const splits = await splitExports(res.run);

      const failCount = report.checks.filter((c) => !c.ok).length;
      appendLog(
        failCount === 0
          ? `All ${report.checks.length} validation checks OK.`
          : `⚠ ${failCount} validation check(s) FAILED — open the Validation tab in the workbook.`
      );

      setOaResult({ ...res, report, mainBuffer, splits });
      download(mainBuffer, `Salary OA${monthSuffix}.xlsx`);
      appendLog(`Downloaded: Salary OA${monthSuffix}.xlsx`);
    } catch (exc) {
      console.error(exc);
      appendLog(`Failed: ${exc.message || exc}`);
      alert(`Run failed\n\n${exc.message || exc}`);
    } finally {
      setRunning(false);
    }
  };

  const runApac = async () => {
    if (!oaResult) {
      alert("Generate the OA workbook first (Step 2). The APAC step needs the OA data in memory.");
      return;
    }
    if (!apacFile) {
      alert("Missing input\n\nSelect last month's APAC Payroll Report (.xlsx).");
      return;
    }
    setRunning(true);
    try {
      appendLog("");
      appendLog(`Reading previous APAC: ${apacFile.name}`);
      let buf = await apacFile.arrayBuffer();
      if (isEncrypted(buf)) {
        appendLog("File is encrypted — decrypting (takes a few seconds)…");
        buf = await decryptOffice(buf, apacPassword);
        appendLog("Decrypted OK.");
      } else {
        appendLog("File is not encrypted — using as-is.");
      }

      const result = await buildApac(buf, oaResult.run);
      appendLog(
        `APAC rows written: ${result.employeesWritten} employees + ${result.internsWritten} interns. ` +
          `New joiners: ${result.newJoiners.length} | Leavers excluded: ${result.leavers.length}`
      );

      const apacBuffer = await saveApac(result.workbook);
      const manualBuffer = await writeManualReport(result.manual);
      const reco = await writeRecoReport(
        result.recoRows, result.recoSummary, oaResult.run,
        oaResult.adminTotal, oaResult.edliTotal, monthLabel.trim()
      );

      setApacResult({ ...result, apacBuffer, manualBuffer, recoBuffer: reco.buffer, totals: reco.totals });
      download(apacBuffer, `APAC Payroll Report${monthSuffix}.xlsx`);
      download(manualBuffer, `APAC Manual Fill${monthSuffix}.xlsx`);
      download(reco.buffer, `APAC Reconciliation${monthSuffix}.xlsx`);
      appendLog(`Downloaded: APAC Payroll Report / Manual Fill / Reconciliation${monthSuffix}.`);
      appendLog("⚠ Filled APAC is UNENCRYPTED — re-add the password in Excel before sending to HQ.");
    } catch (exc) {
      console.error(exc);
      appendLog(`Failed: ${exc.message || exc}`);
      alert(`APAC failed\n\n${exc.message || exc}`);
    } finally {
      setRunning(false);
    }
  };

  return (
    <div className="soa">
      {/* ------------------------------------------------ instructions */}
      <section className="card soa-howto">
        <h2>Salary OA + APAC — monthly run in 3 steps</h2>
        <ol>
          <li>
            <strong>Step 1 — Upload inputs.</strong> Select this month's <em>Salary register</em> (from the
            payroll consultant, <code>Sal-*.xls</code>) and the <em>PF MIS report</em> (
            <code>pfreport*.xls</code>). Type the month label (e.g. "June 2026"). Add any loan EMIs and
            intern stipends below if applicable.
          </li>
          <li>
            <strong>Step 2 — Generate OA workbook.</strong> Click the blue button. Check that all 9
            validation rows say OK. The main <em>Salary OA</em> workbook downloads automatically; use the
            buttons to download the 3 paste-ready files (Service / Non Service / Final Combined) for
            accounts.
          </li>
          <li>
            <strong>Step 3 — Generate APAC.</strong> Upload <em>last month's APAC Payroll Report</em>{" "}
            (password is pre-filled). Click Generate APAC — the Filled APAC, Manual Fill list and
            email-ready Reconciliation download automatically. Then: fix the yellow cells per the Manual
            Fill list, and <strong>re-encrypt the Filled APAC in Excel</strong> (File → Info → Protect
            Workbook → Encrypt with Password) before sending to HQ.
          </li>
        </ol>
        <p className="soa-note-line">
          Everything runs locally in your browser — no file is uploaded anywhere. PF Admin and EDLI are
          auto-pulled from the PF MIS; Aditya Kumar → TECHNICAL and Human Resources → OPERATIONS rules are
          applied automatically.
        </p>
      </section>

      {/* ------------------------------------------------ step 1 */}
      <section className="card">
        <h2 className="soa-step-title"><span className="soa-step-no">1</span> Inputs</h2>
        <FilePick label="Salary register" required accept=".xls,.xlsx,.xlsm" file={salaryFile}
          onFile={onSalaryFile} hint="Sal-*.xls from the payroll consultant" />
        <FilePick label="PF MIS report" required accept=".xls,.xlsx" file={pfFile}
          onFile={setPfFile} hint="pfreport*.xls — needs PF Ac. 1 / Admin Ac. 2 / EDLI Ac. 21" />
        <div className="soa-meta-row">
          <label>Month label</label>
          <input className="soa-input" placeholder="e.g. June 2026" value={monthLabel}
            onChange={(e) => setMonthLabel(e.target.value)} />
        </div>

        <details className="soa-sub">
          <summary>Loan EMIs recovered this month ({loans.filter((l) => l.code && Number(l.amount)).length})</summary>
          <p className="soa-hint">
            Pick the employee and enter the EMI amount. Becomes a negative LOANRECOVERY row in the OA and
            flows to APAC column BX. The dropdown fills as soon as the salary register is selected.
          </p>
          {loans.map((ln, i) => (
            <div key={i} className="soa-inline-row">
              {employeeOptions.length ? (
                <select className="soa-input" value={ln.code}
                  onChange={(e) => setLoans(loans.map((x, j) => (j === i ? { ...x, code: e.target.value } : x)))}>
                  <option value="">— select employee —</option>
                  {employeeOptions.map((o) => (
                    <option key={o.code} value={o.code}>{o.label}</option>
                  ))}
                </select>
              ) : (
                <input className="soa-input" placeholder="Employee code" value={ln.code}
                  onChange={(e) => setLoans(loans.map((x, j) => (j === i ? { ...x, code: e.target.value } : x)))} />
              )}
              <input className="soa-input soa-amt" type="number" placeholder="EMI ₹" value={ln.amount}
                onChange={(e) => setLoans(loans.map((x, j) => (j === i ? { ...x, amount: e.target.value } : x)))} />
              <button className="soa-clear" onClick={() => setLoans(loans.filter((_, j) => j !== i))}>×</button>
            </div>
          ))}
          <button className="soa-add" onClick={() => setLoans([...loans, { code: "", amount: "" }])}>
            + Add loan EMI
          </button>
        </details>

        <details className="soa-sub">
          <summary>Intern stipends ({stipends.filter((s) => s.name.trim() && Number(s.amount)).length})</summary>
          <p className="soa-hint">
            Name + amount + Service / Non Service. Added as STIPEND rows in the OA and as Intern rows at the
            bottom of the APAC.
          </p>
          {stipends.map((s, i) => (
            <div key={i} className="soa-inline-row">
              <input className="soa-input" placeholder="Intern name" value={s.name}
                onChange={(e) => setStipends(stipends.map((x, j) => (j === i ? { ...x, name: e.target.value } : x)))} />
              <input className="soa-input soa-amt" type="number" placeholder="Amount ₹" value={s.amount}
                onChange={(e) => setStipends(stipends.map((x, j) => (j === i ? { ...x, amount: e.target.value } : x)))} />
              <select className="soa-input soa-cat" value={s.category}
                onChange={(e) => setStipends(stipends.map((x, j) => (j === i ? { ...x, category: e.target.value } : x)))}>
                <option>Non Service</option>
                <option>Service</option>
              </select>
              <button className="soa-clear" onClick={() => setStipends(stipends.filter((_, j) => j !== i))}>×</button>
            </div>
          ))}
          <button className="soa-add"
            onClick={() => setStipends([...stipends, { name: "", amount: "", category: "Non Service" }])}>
            + Add stipend
          </button>
        </details>

        <details className="soa-sub" open={showAdvanced} onToggle={(e) => setShowAdvanced(e.target.open)}>
          <summary>Advanced options</summary>
          <div className="soa-meta-row">
            <label>Override PF Admin Charge</label>
            <input className="soa-input soa-amt" type="number" placeholder="auto from PF MIS"
              value={pfAdminOverride} onChange={(e) => setPfAdminOverride(e.target.value)} />
          </div>
          <div className="soa-meta-row">
            <label>Zero-TDS employee codes</label>
            <input className="soa-input" placeholder="comma-separated codes" value={zeroTds}
              onChange={(e) => setZeroTds(e.target.value)} />
          </div>
        </details>
      </section>

      {/* ------------------------------------------------ step 2 */}
      <section className="card">
        <h2 className="soa-step-title"><span className="soa-step-no">2</span> Generate OA workbook</h2>
        <button className="run-btn" onClick={runOa} disabled={running}>
          {running ? "Working…" : "Generate OA Workbook"}
        </button>

        {oaResult && (
          <>
            <div className="soa-dl-row">
              <button className="soa-dl" onClick={() => download(oaResult.mainBuffer, `Salary OA${monthSuffix}.xlsx`)}>
                ⬇ Salary OA (main)
              </button>
              {Object.entries(oaResult.splits).map(([name, buf]) => (
                <button key={name} className="soa-dl" onClick={() => download(buf, `${name}${monthSuffix}.xlsx`)}>
                  ⬇ {name}
                </button>
              ))}
            </div>

            <h3 className="soa-block-title">Validation — {oaResult.report.allOk ? "all OK ✓" : "FAILURES ✗"}</h3>
            <div className="table-wrap">
              <table className="preview-table">
                <thead>
                  <tr><th>Check</th><th>Source</th><th>Output</th><th>Diff</th><th>Status</th></tr>
                </thead>
                <tbody>
                  {oaResult.report.checks.map((c) => (
                    <tr key={c.name} className={c.ok ? "" : "soa-fail-row"}>
                      <td>{c.name}</td>
                      <td className="num">{fmt0(pyRound(c.sourceValue))}</td>
                      <td className="num">{fmt0(pyRound(c.outputValue))}</td>
                      <td className="num">{fmt0(pyRound(c.diff))}</td>
                      <td>{c.ok ? "OK" : "FAIL"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {(oaResult.report.missingInPfMis.length > 0 || oaResult.report.extraInPfMis.length > 0 ||
              oaResult.report.pfAmountMismatches.length > 0) && (
              <p className="soa-warn-line">
                PF MIS gaps: {oaResult.report.missingInPfMis.length} register employee(s) missing from PF MIS,{" "}
                {oaResult.report.extraInPfMis.length} extra PF MIS code(s),{" "}
                {oaResult.report.pfAmountMismatches.length} PF amount mismatch(es). Details in the Validation tab
                of the workbook.
              </p>
            )}
          </>
        )}
      </section>

      {/* ------------------------------------------------ step 3 */}
      <section className="card">
        <h2 className="soa-step-title"><span className="soa-step-no">3</span> Generate APAC Payroll Report</h2>
        <FilePick label="Previous month's APAC" accept=".xlsx" file={apacFile} onFile={setApacFile}
          hint="APAC Payroll Report-<Month>'<YY>.xlsx (encrypted is fine)" />
        <div className="soa-meta-row">
          <label>APAC password</label>
          <input className="soa-input" type="password" value={apacPassword}
            onChange={(e) => setApacPassword(e.target.value)} />
        </div>
        <button className="run-btn" onClick={runApac} disabled={running || !oaResult}
          title={!oaResult ? "Generate the OA workbook first" : ""}>
          {running ? "Working…" : "Generate APAC"}
        </button>

        {apacResult && (
          <>
            <div className="soa-dl-row">
              <button className="soa-dl" onClick={() => download(apacResult.apacBuffer, `APAC Payroll Report${monthSuffix}.xlsx`)}>
                ⬇ Filled APAC
              </button>
              <button className="soa-dl" onClick={() => download(apacResult.manualBuffer, `APAC Manual Fill${monthSuffix}.xlsx`)}>
                ⬇ Manual Fill
              </button>
              <button className="soa-dl" onClick={() => download(apacResult.recoBuffer, `APAC Reconciliation${monthSuffix}.xlsx`)}>
                ⬇ Reconciliation
              </button>
            </div>

            <div className="stats-grid">
              <div className="stat"><span className="stat-label">Employees written</span>
                <span className="stat-value">{apacResult.employeesWritten}</span></div>
              <div className="stat"><span className="stat-label">Interns added</span>
                <span className="stat-value">{apacResult.internsWritten}</span></div>
              <div className={`stat ${apacResult.newJoiners.length ? "warn" : ""}`}>
                <span className="stat-label">New joiners (manual fill)</span>
                <span className="stat-value">{apacResult.newJoiners.length}</span></div>
              <div className="stat"><span className="stat-label">Leavers excluded</span>
                <span className="stat-value">{apacResult.leavers.length}</span></div>
              <div className="stat"><span className="stat-label">Total OA Cost</span>
                <span className="stat-value">₹{fmt0(apacResult.totals.totalOa)}</span></div>
              <div className="stat"><span className="stat-label">APAC Total Labor Cost</span>
                <span className="stat-value">₹{fmt0(apacResult.totals.apacTotalLabor)}</span></div>
            </div>

            {apacResult.manual.length > 0 && (
              <>
                <h3 className="soa-block-title">Manual fill items — {apacResult.manual.length}</h3>
                <div className="table-wrap">
                  <table className="preview-table">
                    <thead>
                      <tr><th>Code</th><th>Name</th><th>Type</th><th>Columns to fill</th><th>Value</th></tr>
                    </thead>
                    <tbody>
                      {apacResult.manual.map((m, i) => (
                        <tr key={i}>
                          <td>{m.code}</td><td>{m.name}</td><td>{m.itemType}</td>
                          <td>{m.apacColumns}</td>
                          <td className="num">{typeof m.suggestedValue === "number" ? fmt0(m.suggestedValue) : m.suggestedValue}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
            <p className="soa-warn-line">
              ⚠ The Filled APAC downloads <strong>unencrypted</strong>. Re-add the password in Excel
              (File → Info → Protect Workbook → Encrypt with Password) before emailing it to HQ.
            </p>
          </>
        )}
      </section>

      {/* ------------------------------------------------ log */}
      <section className="card log-card">
        <h2>Status</h2>
        <pre className="log">{log.join("\n")}</pre>
      </section>
    </div>
  );
}
