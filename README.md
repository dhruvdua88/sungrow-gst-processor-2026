# Sungrow GST Processor 2026

**SIPL GST Monthly ITC and Ledger Tie-Out** — a faithful React + Vite port of the desktop
(`customtkinter` + `openpyxl`) tool. Everything runs **entirely in the browser**: no server,
no upload — your GST files never leave your machine.

**Live app:** https://dhruvdua88.github.io/sungrow-gst-processor-2026/

## What it does

Creates a focused 11-sheet reconciliation workbook from current GST files:

| Sheet | Content |
|---|---|
| Dashboard | KPI cards + exception blocks (ledger mismatches, unresolved B2B, ISD GSTIN check, BOE support & validation flags, category-wise ledger tie, correction flags) |
| Sanitized ITC | PO + Non-PO + Import invoices, resolved against 2B, with ISD/BOE support columns |
| 2B | Every GSTR-2B document (B2B / CDNR / IMPG / ISD) with match status |
| ISD / ISD Eligible | Bill-wise GSTR-6 detail with Eligible/Ineligible classification (supplier-level subset-sum against `isd.elglst`) |
| BOE Detail | Parsed Bills of Entry with GSTIN + 2B-reporting validation |
| 2B vs Books Summary / Detail | Party-level and invoice-level reconciliation with suggested actions |
| ITC vs Ledger / Invoice Ledger Detail | SAP GST REC / RCM ledger movement tie-out per invoice and per category |
| Book Correction Flags | Items needing immediate correction before the 3B claim |

## Inputs

- **SIPL tracker/template** (`.xlsx`) — required (presence-validated, same as the desktop tool)
- **Current GST workbook** (`.xlsx`) — required; must contain `PO`, `Non PO`, `GL Summary` sheets
- **Current GSTR-2B JSON** — required (`returns_R2B_*.json` from the GST portal)
- **Prior focused output files** (`.xlsx`, optional) — party→GSTIN alias matching from prior months
- **BOE PDFs** (`.zip`/`.pdf`, optional) — ICEGATE Bills of Entry; nested ZIPs are scanned
- **ISD JSON** (optional) — GSTR-6 bill-wise b2b detail + `isd.elglst` eligible control total

## Engine highlights (ported 1:1 from the Python tool)

- Company-code 2081 filter, NA-supplier and zero-tax row skips
- RCM evidence from GL (`REVE CH` names, RCM accounts `2221013010/30/40`, tax code `3R`)
- 2B resolution: exact invoice match → supplier-name fuzzy match (difflib-equivalent scorer) → import IGST match to IMPG BoE
- Prior-month alias resolution with conflict detection
- BOE parsing (ICEGATE PART-I summary), dedupe, GSTIN check, 2B `impg` reporting check
- ISD bill eligibility via brute-force supplier subset-sum (±0.50 tolerance) with pro-rate fallback warning
- Ledger tie-out against REC accounts `2221013210/20/30` with PO-reference substring fallback
- Styled output workbook (ExcelJS): blue header bands, warn fills, freeze panes, auto-filter, number formats

## Develop

```bash
npm install
npm run dev        # local dev server
npm run build      # production build to dist/
node test/engine.test.mjs   # end-to-end engine test (21 assertions + parity dump)
```

The test suite was verified for **exact numeric parity** against the original Python engine
on the same synthetic dataset.

## Stack

React 19 · Vite · SheetJS (read) · ExcelJS (write) · pdfjs-dist (BOE text) · JSZip (nested archives)
