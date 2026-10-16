// Usage: node scripts/build-snapshot.js path/to/WOD.xlsx
// Writes data/snapshot.json — the fallback used when SharePoint can't be reached.
const fs = require('fs');
const path = require('path');
const { parseWOD } = require('../lib/parse-wod');

const file = process.argv[2];
if (!file) { console.error('Pass the path to the WOD .xlsx'); process.exit(1); }
const t = Date.now();
const data = parseWOD(fs.readFileSync(file));
data.fetchedAt = new Date().toISOString();
data.origin = 'snapshot';
const out = path.join(__dirname, '..', 'data', 'snapshot.json');
fs.writeFileSync(out, JSON.stringify(data));
console.log(`rows=${data.rows.length} asOf=${data.asOf} sheet=${data.sheet} ${Date.now() - t}ms -> ${out}`);
