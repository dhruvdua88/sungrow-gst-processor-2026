# Ind AS Bridge

Browser-only tool that traces every ledger from the signed local GAAP workbook (Schedule III,
Division I — sheet `T_B` of the Schedule 3 Automation Tool) to the Ind AS financial statements
(Division II) built in the Ind AS working file (sheet `TB`: Local GAAP | Ind AS JV | Ind AS).

**Privacy:** files are parsed by SheetJS inside a Web Worker in the browser. There is no backend,
no upload and no analytics. Never commit client workbooks — `*.xlsb`, `*.xlsx`, `*.xls` are ignored here.

## Use
Open `<pages-root>/indas-bridge/`, drop both files (or click **Load demo data** — fully synthetic).

## Engine rules (`src/engine/`, pure TypeScript, unit-tested)
- Headers are found by text; no fixed rows or columns.
- Join on the full normalised `code/name`; code-only fallback only when the code is unique on both sides.
- **Signed books win:** the xlsb amount starts every bridge; the Ind AS file supplies the mapping and the JV layer.
  Ledgers whose Ind AS-file local GAAP differs are flagged (Differences tab).
- A ledger mapped to several Ind AS rows: keep the split if it adds up; otherwise the row equal to the
  signed amount takes it all (regroup); otherwise the largest row takes the balance.
- Balance sheet = layer G + layer A (profit for the year sits in other equity); P&L = layer G + layer P.
  Equity = assets − liabilities. OCI = movement in Ind AS other equity − Ind AS profit − other equity movements.
- Ties are checked against the Ind AS file's own Balance Sheet and Statement of Profit and Loss.

## Develop
```bash
npm install
npm run dev        # local dev server
npm test           # vitest (synthetic fixtures only)
npm run typecheck && npm run lint
```
From the repository root, `npm run build:all` builds the GST app and then this app into `dist/indas-bridge/`.
