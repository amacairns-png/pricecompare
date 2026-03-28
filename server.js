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

function processFiles(materials, pricingFiles) {
  const SKIP_SHEETS = new Set(['Preface', 'Licensing Details']);
  const materialSet = new Set(materials);
  const found = new Set();
  const results = [];

  for (const filePath of pricingFiles) {
    if (found.size === materialSet.size) break;

    const wb = xlsx.readFile(filePath);
    const basename = path.basename(filePath);
    console.log(`Scanning ${basename} ...`);

    for (const sheetName of wb.SheetNames) {
      if (SKIP_SHEETS.has(sheetName)) continue;

      const sheet = wb.Sheets[sheetName];
      const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      for (let i = 3; i < rows.length; i++) {
        const materialCell = String(rows[i][0]).trim();
        if (!materialCell || materialCell === 'undefined') continue;

        if (materialSet.has(materialCell) && !found.has(materialCell)) {
          found.add(materialCell);
          results.push({
            material:      materialCell,
            priceListItem: String(rows[i][1] ?? '').trim(),
            pricePerUnit:  String(rows[i][5] ?? '').trim(),
            sourceFile:    basename
          });
        }
      }

      if (found.size === materialSet.size) break;
    }
  }

  return results;
}

function writeResults(results) {
  const lines = ['Material\tPrice List Item\tPrice per Unit\tSource File'];
  for (const r of results) {
    lines.push(`${r.material}\t${r.priceListItem}\t${r.pricePerUnit}\t${r.sourceFile}`);
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
