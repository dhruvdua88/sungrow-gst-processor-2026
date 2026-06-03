// GSTR-2B JSON loading — mirrors load_r2b_docs / _twob_boe_set
import { asText, asFloat, parseDate, taxTotal, normDoc, STATE_CODE } from "./helpers.js";

export function makeR2BDoc(fields) {
  return {
    section: "", gstin: "", party: "", doc_no: "", doc_date: null,
    taxable: 0, igst: 0, cgst: 0, sgst: 0, cess: 0, total_tax: 0,
    reverse_charge: "N", itc_available: "Y", pos: "",
    source_period: "", raw_type: "", matched: false,
    ...fields,
  };
}

export function loadR2bDocs(payload) {
  const data = payload.data;
  if (!data) throw new Error("GSTR-2B JSON has no 'data' key.");
  const docs = [];
  const docdata = data.docdata || {};

  const add = (section, gstin, party, docNo, docDate, tx, ig, cg, sg, ce, rev, itc, pos, period, typ) => {
    docs.push(makeR2BDoc({
      section,
      gstin: asText(gstin), party: asText(party),
      doc_no: asText(docNo), doc_date: parseDate(docDate),
      taxable: asFloat(tx), igst: asFloat(ig), cgst: asFloat(cg),
      sgst: asFloat(sg), cess: asFloat(ce),
      total_tax: taxTotal(ig, cg, sg),
      reverse_charge: asText(rev || "N") || "N",
      itc_available: asText(itc || "Y") || "Y",
      pos: STATE_CODE[asText(pos)] || asText(pos),
      source_period: period, raw_type: asText(typ),
    }));
  };

  for (const sup of docdata.b2b || []) {
    for (const inv of sup.inv || []) {
      add("B2B", sup.ctin, sup.trdnm, inv.inum, inv.dt,
        inv.txval, inv.igst, inv.cgst, inv.sgst,
        inv.cess, inv.rev, inv.itcavl, inv.pos,
        asText(sup.supprd), inv.typ);
    }
  }
  for (const sup of docdata.cdnr || []) {
    for (const nt of sup.nt || []) {
      add("CDNR", sup.ctin, sup.trdnm, nt.ntnum, nt.dt,
        nt.txval, nt.igst, nt.cgst, nt.sgst,
        nt.cess, nt.rev, nt.itcavl, nt.pos,
        asText(sup.supprd), nt.typ);
    }
  }
  for (const it of docdata.impg || []) {
    add("IMPG", "IMPORTGOODS", "Import Goods", it.boenum, it.boedt,
      it.txval, it.igst, it.cgst, it.sgst,
      it.cess, "N", "Y", it.portcode,
      asText(data.rtnprd), "IMPORTGOODS");
  }
  for (const sup of docdata.isd || []) {
    for (const it of sup.doclist || []) {
      add("ISD", sup.ctin, sup.trdnm,
        it.docnum || it.inum, it.docdt || it.dt,
        it.txval, it.igst, it.cgst, it.sgst,
        it.cess, it.rev, it.itcavl, "",
        asText(sup.supprd), it.doctyp || it.typ);
    }
  }
  return docs;
}

// Set of normalised BOE numbers reported in the recipient's GSTR-2B (impg)
export function twobBoeSet(payload) {
  try {
    const impg = ((payload.data || {}).docdata || {}).impg || [];
    const out = new Set();
    for (const item of impg) {
      const be = asText(item.boenum || item.boe || "");
      if (be) out.add(normDoc(be));
    }
    return out;
  } catch {
    return new Set();
  }
}
