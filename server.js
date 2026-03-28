const express = require('express');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const PRICE_DIR   = 'C:/Pricelists/cloud';
const MAPPING     = 'C:/Pricelists/cloud/Material Nr Mapping.xlsx';
const RESULTS     = 'C:/Pricelists/cloud/results.txt';
const SKUS        = 'C:/Pricelists/cloud/skus.txt';
const PORT        = 3000;

const app = express();
app.use(express.static(path.join(__dirname, 'public')));

function readMaterials() {
  const wb = xlsx.readFile(MAPPING);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const materials = [];
  for (let i = 1; i < rows.length; i++) {
    const val = String(rows[i][0]).trim();
    if (val && val !== 'undefined') {
      materials.push(val);
    }
  }
  return materials;
}

function discoverPricingFiles() {
  const files = fs.readdirSync(PRICE_DIR);
  const matched = [];

  for (const name of files) {
    if (name.startsWith('~$')) continue;
    const m = name.match(/^PriceList_AU_CloudDirect_(\d{8})/);
    if (!m) continue;
    const dateInt = parseInt(m[1], 10);
    matched.push({ name, dateInt, hasParens: name.includes(' (') });
  }

  // Sort by date, then base files before parenthetical variants
  matched.sort((a, b) => {
    if (a.dateInt !== b.dateInt) return a.dateInt - b.dateInt;
    return a.hasParens - b.hasParens;
  });

  return matched.map(f => path.join(PRICE_DIR, f.name));
}

function detectPriceColumn(rows) {
  // Row 1 headers, row 2 sub-headers
  const h1 = rows[1] || [];
  const h2 = rows[2] || [];
  // Look for 'Price per Unit' in row 1
  for (let c = 0; c < h1.length; c++) {
    if (String(h1[c]).trim() === 'Price per Unit') return c;
  }
  // Fall back to 'LP Value' in row 2 (new format)
  for (let c = 0; c < h2.length; c++) {
    if (String(h2[c]).trim() === 'LP Value') return c;
  }
  return 5; // default fallback
}

function processFiles(materials, pricingFiles) {
  const SKIP_SHEETS = new Set(['Preface', 'Licensing Details']);
  const materialSet = new Set(materials);
  const results = [];

  for (const filePath of pricingFiles) {
    const wb = xlsx.readFile(filePath);
    const basename = path.basename(filePath);
    console.log(`Scanning ${basename} ...`);

    // Track which materials have been found in this file already
    const foundInFile = new Set();

    for (const sheetName of wb.SheetNames) {
      if (SKIP_SHEETS.has(sheetName)) continue;
      if (foundInFile.size === materialSet.size) break;

      const sheet = wb.Sheets[sheetName];
      const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      const priceCol = detectPriceColumn(rows);

      for (let i = 3; i < rows.length; i++) {
        const materialCell = String(rows[i][0]).trim();
        if (!materialCell || materialCell === 'undefined') continue;

        if (materialSet.has(materialCell) && !foundInFile.has(materialCell)) {
          foundInFile.add(materialCell);
          results.push({
            material:      materialCell,
            priceListItem: String(rows[i][1] ?? '').trim(),
            inBlocksOf:    String(rows[i][2] ?? '').trim(),
            metrics:       String(rows[i][4] ?? '').trim(),
            pricePerUnit:  String(rows[i][priceCol] ?? '').trim(),
            sourceFile:    basename
          });
        }
      }
    }
  }

  return results;
}

function writeResults(results) {
  const lines = ['Material\tPrice List Item\tIn Blocks Of\tMetrics\tPrice per Unit\tSource File'];
  for (const r of results) {
    lines.push(`${r.material}\t${r.priceListItem}\t${r.inBlocksOf}\t${r.metrics}\t${r.pricePerUnit}\t${r.sourceFile}`);
  }
  fs.writeFileSync(RESULTS, lines.join('\n'), 'utf8');
}

app.get('/api/process', (req, res) => {
  try {
    const materials = readMaterials();
    console.log(`Found ${materials.length} materials to look up.`);

    fs.writeFileSync(SKUS, materials.join('\n'), 'utf8');
    console.log(`SKUs written to ${SKUS}`);

    const pricingFiles = discoverPricingFiles();
    console.log(`Found ${pricingFiles.length} pricing files.`);

    const results = processFiles(materials, pricingFiles);
    console.log(`Found ${results.length} entries across ${pricingFiles.length} files.`);

    writeResults(results);
    console.log(`Results written to ${RESULTS}`);

    res.json({ success: true, count: results.length, results });
  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
