# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm install   # Install dependencies
npm start     # Start the Express server on http://localhost:3000
```

No test or lint scripts are configured.

## Architecture

Single-file Node.js/Express backend (`server.js`) + static frontend (`public/index.html`).

**Data flow:**
1. `GET /api/process` triggers the full ETL pipeline
2. `readMaterials()` — reads material numbers from `C:/Pricelists/cloud/Material Nr Mapping.xlsx` (column A, skipping header row)
3. `discoverPricingFiles()` — finds files matching `PriceList_AU_CloudDirect_YYYYMMDD*.xlsx` in `C:/Pricelists/cloud/`, sorted by date ascending (base files before `(n)` variants)
4. `processFiles()` — iterates files/sheets in order, skipping sheets named `Preface` or `Licensing Details`, reading material from col A (index 0), price list item from col B (index 1), price per unit from col F (index 5); data rows start at index 3; stops early once all materials are found
5. `writeResults()` — writes tab-separated output to `C:/Pricelists/cloud/results.txt`; also writes raw SKU list to `C:/Pricelists/cloud/skus.txt`

**Hardcoded paths** (all under `C:/Pricelists/cloud/`):
- `Material Nr Mapping.xlsx` — input: material numbers to look up
- `PriceList_AU_CloudDirect_YYYYMMDD*.xlsx` — input: price list files
- `results.txt` — output: tab-separated results (Material, Price List Item, Price per Unit, Source File)
- `skus.txt` — output: list of material numbers that were searched

The frontend is vanilla JS/HTML in `public/index.html` — calls `/api/process`, then renders the returned `results` array as an HTML table.
