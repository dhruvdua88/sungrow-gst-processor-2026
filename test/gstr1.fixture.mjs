// GSTR-1 Processor regression fixtures — locks the June-2026 case from the hardening spec.
//   node test/gstr1.fixture.mjs
// The two source files are REAL client data and are deliberately NOT committed. Point the test at
// them via GSTR1_FIXTURE_DIR (defaults to ~/Downloads). If they're absent the test SKIPS (exit 0),
// so CI without the data never fails — but on a machine that has them, every §8 number is asserted.
import * as XLSX from "xlsx";
import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { processGstr1, gstinChecksumValid } from "../src/gstr1/engine.js";

const DIR = process.env.GSTR1_FIXTURE_DIR || join(homedir(), "Downloads");
const EINV = join(DIR, "EINV_29AAXCS2197N1Z4_2026-27 (2).xlsx");
const SALES = join(DIR, "GSTR-1 Data June 2026.xlsx");
if (!existsSync(EINV) || !existsSync(SALES)) {
  console.log(`SKIP  June-2026 fixtures not found in ${DIR} (set GSTR1_FIXTURE_DIR to run).`);
  process.exit(0);
}

const rd = (p) => XLSX.read(readFileSync(p), { type: "buffer", cellDates: true });
const r = processGstr1(rd(EINV), rd(SALES), { gstin: "29AAXCS2197N1Z4", fp: "062026", periodLabel: "June 2026" });

let fail = 0;
const eq = (label, actual, expected) => {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  console.log(`${ok ? "PASS" : "FAIL"}  ${label}: ${JSON.stringify(actual)}${ok ? "" : ` (expected ${JSON.stringify(expected)})`}`);
  if (!ok) fail++;
};
const sortA = (a) => a.slice().sort();

// ---- §8 regression fixtures ----
eq("Book taxable total", r.fileOverview.sales.txbl, 1323060247.96);
eq("B2B / B2C / zero-value counts", [r.counts.jsonB2bInvoices, r.fileOverview.sales.byCustomer["Unregistered (B2C)"].invoices, r.counts.zeroDocs], [325, 1, 42]);
eq("Total documents (= 368)", r.counts.jsonB2bInvoices + r.fileOverview.sales.byCustomer["Unregistered (B2C)"].invoices + r.counts.zeroDocs, 368);
eq("Portal valid / cancelled", [r.counts.validInvoices, r.counts.cancelled], [328, 5]);
eq("Cancelled IRN list", r.cancelled, ["90000763", "90000765", "90000766", "90000938", "90000974"]);
eq("Portal valid taxable", r.totals.validB2bTxbl, 1339047920.72);
eq("Portal duplicates (C21) = 1", r.duplicates.map((d) => [d.portalNo, d.duplicateOf, d.taxable]), [["9000793", "90000793", 10680000]]);
eq("True missing-from-book = 0", r.recon.trueMissing.length, 0);
eq("No-IRN book invoices = 1 (90000874, 410)", r.recon.notEinvoiced.map((d) => [d.inum, d.txbl]), [["90000874", 410]]);
eq("Matched deviations = 12", r.recon.deviations.length, 12);
eq("Net portal − book = +1,18,93,815.76", r.recon.direction.netUnderReported, 11893815.76);
eq("Understated 8 / 1,29,57,915.76", [r.recon.direction.underCount, r.recon.direction.underAmt], [8, 12957915.76]);
eq("Overstated 4 / 10,64,100", [r.recon.direction.overCount, r.recon.direction.overAmt], [4, 1064100]);
eq("Zero-tax-on-taxable (C22) = 1 (90000941, 58,30,275)", r.recon.zeroTax.map((z) => [z.inum, z.txbl]), [["90000941", 5830275]]);
eq("HSN-not-in-portal-master (C23a) = 2", sortA(r.recon.hsnNotInMaster), sortA(["562123", "85041010"]));
eq("Serial gap (C11) = 1 (90001042)", r.checks.find((c) => c.id === "C11").actual, "1 90001042");
eq("Bridge residual ~ 0 (emergent)", Math.abs(r.bridge.residual) < 1 && Math.abs(r.bridge.residual0) < 1, true);

// ---- GSTIN checksum sanity ----
eq("GSTIN checksum accepts a valid GSTIN", gstinChecksumValid("29AAXCS2197N1Z4"), true);
eq("GSTIN checksum rejects a bad check digit", gstinChecksumValid("29AAXCS2197N1Z5"), false);

console.log(fail ? `\n${fail} FAILURES` : "\nALL FIXTURES PASSED");
process.exit(fail ? 1 : 0);
