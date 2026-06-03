// BOE PDF parsing + validation — mirrors _pdf_text/_iter_pdf_inputs/_parse_boe_doc/
// load_boe_docs/validate_boe_docs. Browser version: pdfjs-dist for text, JSZip for archives.
import JSZip from "jszip";
import { asText, asFloat, normDoc, parseDate, round2, RECIPIENT_GSTIN } from "./helpers.js";
import { twobBoeSet } from "./r2b.js";

// pdfjs is loaded lazily so the engine stays importable outside the browser
// (unit tests in node) — the worker URL is only resolved when a PDF is parsed.
let pdfjsPromise = null;
function getPdfjs() {
  if (!pdfjsPromise) {
    pdfjsPromise = Promise.all([
      import("pdfjs-dist"),
      import("pdfjs-dist/build/pdf.worker.min.mjs?url"),
    ]).then(([lib, worker]) => {
      lib.GlobalWorkerOptions.workerSrc = worker.default;
      return lib;
    });
  }
  return pdfjsPromise;
}

export function makeBOEDoc(fields) {
  return {
    source_path: "", file_name: "", be_no: "", be_date: null,
    port_code: "", importer_gstin: "", importer_name: "",
    invoice_no: "", invoice_amount: 0, currency: "",
    assessable_value: 0, bcd: 0, sws: 0, igst: 0, total_duty: 0, total_amount: 0,
    challan_no: "", payment_amount: 0, taxable_for_igst: 0,
    matched: false, matched_source_doc: "", matched_itc_igst: 0, igst_diff: 0,
    match_method: "", gstin_check: "", twob_check: "", validation_action: "",
    ...fields,
  };
}

async function pdfText(data) {
  const pdfjsLib = await getPdfjs();
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(data) }).promise;
  const parts = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    // Reassemble line structure: group items by rounded Y, sort by X.
    const lines = new Map();
    for (const item of content.items) {
      const y = Math.round(item.transform[5]);
      if (!lines.has(y)) lines.set(y, []);
      lines.get(y).push({ x: item.transform[4], str: item.str });
    }
    const ordered = [...lines.entries()].sort((a, b) => b[0] - a[0]);
    const pageLines = ordered.map(([, items]) =>
      items.sort((a, b) => a.x - b.x).map((i) => i.str).join(" ").trim()
    );
    parts.push(pageLines.join("\n"));
  }
  await doc.destroy();
  return parts.join("\n");
}

// Walk BOE input recursively, including nested zip archives.
// files: [{name, data: ArrayBuffer, relativePath?}] from <input multiple> / directory picker
export async function iterPdfInputs(inputFiles) {
  if (!inputFiles || !inputFiles.length) return [];
  const out = [];

  async function iterZipBytes(zipData, prefix, depth) {
    if (depth > 6) return;
    let archive;
    try {
      archive = await JSZip.loadAsync(zipData);
    } catch {
      return;
    }
    const names = Object.keys(archive.files).sort();
    for (const entry of names) {
      const file = archive.files[entry];
      if (file.dir) continue;
      const lower = entry.toLowerCase();
      if (lower.endsWith(".pdf")) {
        const data = await file.async("arraybuffer");
        out.push({ sourcePath: `${prefix}!${entry}`, fileName: entry.split("/").pop(), data });
      } else if (lower.endsWith(".zip")) {
        const data = await file.async("arraybuffer");
        await iterZipBytes(data, `${prefix}!${entry}`, depth + 1);
      }
    }
  }

  const sorted = [...inputFiles].sort((a, b) =>
    (a.relativePath || a.name).localeCompare(b.relativePath || b.name)
  );
  for (const f of sorted) {
    const lower = f.name.toLowerCase();
    const displayPath = f.relativePath || f.name;
    if (lower.endsWith(".pdf")) {
      out.push({ sourcePath: displayPath, fileName: f.name, data: f.data });
    } else if (lower.endsWith(".zip")) {
      await iterZipBytes(f.data, displayPath, 1);
    }
  }
  return out;
}

function boeInvoiceHint(sourcePath) {
  const inner = sourcePath.split("!").pop();
  const parts = inner.split("/").slice(0, -1);
  for (const part of parts.reverse()) {
    const clean = asText(part);
    if (!clean || clean.startsWith(".")) continue;
    const upper = clean.toUpperCase();
    if (upper.includes("SHIP") && upper.includes("BILL")) continue;
    if (normDoc(clean).length >= 5 && (/[A-Z]/.test(upper) || clean.includes("+") || clean.includes("-"))) {
      return clean;
    }
  }
  return "";
}

export function parseBoeDoc(text, sourcePath, fileName) {
  const normalized = text.toUpperCase().replace(/\s+/g, " ");
  // Loosened sentinel: accept any of the common ICEGATE BOE templates.
  const isBoe =
    normalized.includes("BILL OF ENTRY FOR HOME CONSUMPTION") ||
    normalized.includes("BILL OF ENTRY FOR WAREHOUSING") ||
    normalized.includes("BILL OF ENTRY FOR EX-BOND CLEARANCE");
  if (!isBoe) return null;
  const hasPart1 = /\bPART\s*-\s*I\b/.test(normalized);
  const hasSummary = normalized.includes("BILL OF ENTRY SUMMARY");
  if (!hasPart1 || !hasSummary) return null;
  const head = text.slice(0, 3500);

  let beNo = "";
  let beDate = null;
  let m = head.match(/^\s*(\d{5,8})\s+(\d{2}\/\d{2}\/\d{4})\s*$/m);
  if (m) {
    beNo = m[1];
    beDate = parseDate(m[2]);
  }
  if (!beNo && (m = fileName.match(/^(\d{5,8})/))) beNo = m[1];

  const port = head.match(/\b(IN[A-Z]{3}\d)\b/);
  const gstin = head.toUpperCase().match(/\b\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]\b/);
  const importer = head.match(/PART\s*-\s*I\s*-\s*BILL OF ENTRY SUMMARY[\s\S]*?ASSESSED\s+COPY\s+([^\n]+)/i);

  const invoiceCandidates = [...head.matchAll(/^\s*1\s+([A-Z0-9+/_-]{5,})\s+([0-9]+(?:\.[0-9]+)?)\s*$/gm)]
    .map((mm) => [mm[1], mm[2]]);
  let invoiceMatch = null;
  if (invoiceCandidates.length) {
    invoiceMatch = [...invoiceCandidates].sort((a, b) => {
      const ka = !a[0].includes("+") && !a[0].includes("-") && !/[A-Z]/.test(a[0]) ? 1 : 0;
      const kb = !b[0].includes("+") && !b[0].includes("-") && !/[A-Z]/.test(b[0]) ? 1 : 0;
      if (ka !== kb) return ka - kb;
      return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0;
    })[0];
  }
  const invoiceHint = boeInvoiceHint(sourcePath);
  const currency = head.toUpperCase().match(/^\s*(USD|EUR|CNY|INR|JPY|GBP)\s*$/m);
  const challan = head.match(/^\s*1\s+(\d{6,})\s+([0-9]+(?:\.[0-9]+)?)\s*$/m);

  let assessableValue = 0, bcd = 0, sws = 0, igst = 0, totalDuty = 0, totalAmount = 0;
  const duty = head.match(
    /^\s*\d{2}-[A-Z]{3}-\d{2}\s*$[\s\S]*?^\s*\d{2}-[A-Z]{3}-\d{2}\s*$[\s\S]*?^\s*([0-9. ]+)\s*$[\s\S]*?^\s*([0-9. ]+)\s*$[\s\S]*?^\s*([0-9.]+)\s*$/m
  );
  if (duty) {
    const first = duty[1].trim().split(/\s+/).map(asFloat);
    const second = duty[2].trim().split(/\s+/).map(asFloat);
    if (first.length >= 6) {
      bcd = first[0];
      sws = first[2];
      igst = first[3];
      assessableValue = first[first.length - 1];
    }
    if (second.length) totalDuty = second[second.length - 1];
    totalAmount = asFloat(duty[3]);
  }

  return makeBOEDoc({
    source_path: sourcePath, file_name: fileName, be_no: beNo, be_date: beDate,
    port_code: port ? port[1] : "",
    importer_gstin: gstin ? gstin[0] : "",
    importer_name: importer ? asText(importer[1]) : "",
    invoice_no: invoiceHint || (invoiceMatch ? invoiceMatch[0] : ""),
    invoice_amount: invoiceMatch ? asFloat(invoiceMatch[1]) : 0,
    currency: currency ? currency[1] : "",
    assessable_value: round2(assessableValue), bcd: round2(bcd), sws: round2(sws),
    igst: round2(igst), total_duty: round2(totalDuty), total_amount: round2(totalAmount),
    challan_no: challan ? challan[1] : "",
    payment_amount: challan ? asFloat(challan[2]) : 0,
    taxable_for_igst: assessableValue ? round2(assessableValue + bcd + sws) : 0,
  });
}

export async function loadBoeDocs(inputFiles, onProgress) {
  const docs = [];
  const seen = new Set();
  const inspection = {
    pdfs_found: 0,
    text_extraction_failed: 0,
    format_mismatch: 0,
    parsed: 0,
    deduped: 0,
    skipped_samples: [],
    all_pdf_paths: [],
  };
  const pdfs = await iterPdfInputs(inputFiles);
  for (const { sourcePath, fileName, data } of pdfs) {
    inspection.pdfs_found += 1;
    inspection.all_pdf_paths.push(sourcePath);
    if (onProgress) onProgress(`Parsing BOE PDF ${inspection.pdfs_found}/${pdfs.length}: ${fileName}`);
    let text;
    try {
      text = await pdfText(data);
    } catch (exc) {
      inspection.text_extraction_failed += 1;
      if (inspection.skipped_samples.length < 5) {
        inspection.skipped_samples.push([sourcePath, `text extraction failed: ${exc.name}: ${exc.message}`]);
      }
      continue;
    }
    const doc = parseBoeDoc(text, sourcePath, fileName);
    if (!doc) {
      inspection.format_mismatch += 1;
      if (inspection.skipped_samples.length < 5) {
        inspection.skipped_samples.push([sourcePath, `sentinel/format mismatch (text length ${text.length})`]);
      }
      continue;
    }
    const key = `${normDoc(doc.be_no)}|${normDoc(doc.invoice_no)}|${round2(doc.igst)}`;
    if (seen.has(key)) {
      inspection.deduped += 1;
      continue;
    }
    seen.add(key);
    docs.push(doc);
    inspection.parsed += 1;
  }
  return [docs, inspection];
}

export function validateBoeDocs(boeDocs, r2bPayload, recipientGstin = RECIPIENT_GSTIN) {
  const twobSet = twobBoeSet(r2bPayload);
  const targetGstin = (recipientGstin || "").toUpperCase().trim();

  const summary = {
    recipient_gstin: targetGstin,
    twob_boe_count: twobSet.size,
    gstin_mismatch: 0,
    gstin_missing: 0,
    not_in_2b: 0,
    both_flags: 0,
    fully_ok: 0,
  };

  for (const d of boeDocs) {
    const boeGstin = (d.importer_gstin || "").toUpperCase().trim();
    if (!boeGstin) {
      d.gstin_check = "GSTIN missing on BOE";
      summary.gstin_missing += 1;
    } else if (boeGstin === targetGstin) {
      d.gstin_check = "OK";
    } else {
      d.gstin_check = `GSTIN mismatch (${boeGstin})`;
      summary.gstin_mismatch += 1;
    }

    const in2b = twobSet.has(normDoc(d.be_no));
    d.twob_check = in2b ? "Reported in 2B" : "Not in 2B";
    if (!in2b) summary.not_in_2b += 1;

    const reasons = [];
    if (d.gstin_check !== "OK") {
      reasons.push(
        "BOE importer GSTIN differs from recipient - credit legally sits with a different GSTIN; reverse claim or file BOE amendment (Sec. 149 Customs Act)"
      );
    }
    if (!in2b) {
      reasons.push(
        "BOE not reflected in recipient's GSTR-2B impg - check Customs transmission to GSTN or whether BOE belongs to a different GSTIN"
      );
    }
    if (reasons.length) {
      d.validation_action = reasons.join(" | ");
      if (d.gstin_check !== "OK" && !in2b) summary.both_flags += 1;
    } else {
      d.validation_action = "OK - GSTIN matches recipient and BOE reported in 2B";
      summary.fully_ok += 1;
    }
  }

  return summary;
}

// BOE matching — mirrors apply_boe_matches
export function applyBoeMatches(invoices, boeDocs) {
  if (!boeDocs.length) return;
  for (const inv of invoices.filter((i) => i.source === "IMPORT")) {
    const keys = new Set([normDoc(inv.invoice_no), normDoc(inv.sap_invoice_no)]);
    if (inv.matched_doc) keys.add(normDoc(inv.matched_doc.doc_no));
    for (const d of inv.source_docs) keys.add(normDoc(d));
    keys.delete("");
    const candidates = boeDocs.filter((b) => !b.matched && Math.abs(inv.igst - b.igst) <= 1);
    if (!candidates.length) continue;
    const cand = [...candidates].sort((a, b) => {
      const ka1 = keys.has(normDoc(a.invoice_no)) ? 0 : 1;
      const kb1 = keys.has(normDoc(b.invoice_no)) ? 0 : 1;
      if (ka1 !== kb1) return ka1 - kb1;
      const ka2 = keys.has(normDoc(a.be_no)) ? 0 : 1;
      const kb2 = keys.has(normDoc(b.be_no)) ? 0 : 1;
      if (ka2 !== kb2) return ka2 - kb2;
      return Math.abs(inv.igst - a.igst) - Math.abs(inv.igst - b.igst);
    })[0];
    cand.matched = true;
    cand.matched_source_doc = inv.sap_invoice_no;
    cand.matched_itc_igst = round2(inv.igst);
    cand.igst_diff = round2(inv.igst - cand.igst);
    cand.match_method = keys.has(normDoc(cand.invoice_no))
      ? "IGST + import invoice/reference match"
      : keys.has(normDoc(cand.be_no))
        ? "IGST + BoE number match"
        : "IGST amount match";
    inv.support.boe_doc = cand;
    inv.support.support_reference = `BOE ${cand.be_no} dated ${fmtDateLocal(cand.be_date)}`;
    if (!inv.taxable && cand.taxable_for_igst) inv.taxable = cand.taxable_for_igst;
    inv.mapping_notes = `${inv.mapping_notes} BOE support matched: ${cand.be_no} / ${cand.invoice_no}.`.trim();
  }
}

function fmtDateLocal(value) {
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) return "";
  const dd = String(value.getDate()).padStart(2, "0");
  const mm = String(value.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${value.getFullYear()}`;
}
