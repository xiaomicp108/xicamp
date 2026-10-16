// Dependency-free XLSX reader for the WOD sheet.
// Unzips with Node's zlib, reads sharedStrings + the "WOD" sheet XML,
// and returns a compact JSON payload the dashboard consumes.
'use strict';
const zlib = require('zlib');

const ELEMENTS = ['Poster 1', 'Poster 2', 'Dangler', 'AIP 1 Phone', 'AIP 2 AIOT',
  'Tent Card 1', 'Tent Card 2', 'Cube Box', 'Broucher', 'Buntings'];

function unzip(buf) {
  // Find End Of Central Directory
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('Not a valid XLSX (zip) file');
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const files = {};
  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) break;
    const method = buf.readUInt16LE(p + 10);
    const csize = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localOff = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen);
    files[name] = { method, csize, localOff };
    p += 46 + nameLen + extraLen + commentLen;
  }
  return function read(name) {
    const f = files[name];
    if (!f) return null;
    const lo = f.localOff;
    const start = lo + 30 + buf.readUInt16LE(lo + 26) + buf.readUInt16LE(lo + 28);
    const data = buf.subarray(start, start + f.csize);
    const out = f.method === 0 ? data : zlib.inflateRawSync(data);
    return out.toString('utf8');
  };
}

function decode(s) {
  return s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (m, e) => {
    if (e === 'lt') return '<'; if (e === 'gt') return '>'; if (e === 'amp') return '&';
    if (e === 'quot') return '"'; if (e === 'apos') return "'";
    if (e[1] === 'x') return String.fromCodePoint(parseInt(e.slice(2), 16));
    return String.fromCodePoint(parseInt(e.slice(1), 10));
  });
}

function colIndex(ref) {
  let n = 0;
  for (let i = 0; i < ref.length; i++) {
    const c = ref.charCodeAt(i);
    if (c < 65 || c > 90) break;
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

function serialToISO(v) {
  const d = new Date(Math.round((v - 25569) * 86400000));
  return d.toISOString().slice(0, 10);
}

function findSheetPath(read, wanted) {
  const wb = read('xl/workbook.xml') || '';
  const rels = read('xl/_rels/workbook.xml.rels') || '';
  const sheets = [...wb.matchAll(/<sheet\b[^>]*>/g)].map(m => {
    const tag = m[0];
    const name = decode((tag.match(/\bname="([^"]*)"/) || [])[1] || '');
    const rid = (tag.match(/\br:id="([^"]*)"/) || tag.match(/\bid="([^"]*)"/) || [])[1];
    return { name, rid };
  });
  const pick = sheets.find(s => s.name.trim().toLowerCase() === wanted.toLowerCase()) || sheets[0];
  if (!pick) return 'xl/worksheets/sheet1.xml';
  const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map(m => m[0])
    .find(t => t.includes(`Id="${pick.rid}"`));
  let target = rel ? (rel.match(/Target="([^"]*)"/) || [])[1] : 'worksheets/sheet1.xml';
  target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
  return { path: 'xl/' + target, name: pick.name };
}

function parseWOD(buffer) {
  const read = unzip(Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer));
  // Shared strings
  const sst = [];
  const ssXml = read('xl/sharedStrings.xml');
  if (ssXml) {
    for (const m of ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      let t = '';
      for (const tm of m[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) t += tm[1];
      sst.push(decode(t));
    }
  }
  const sheet = findSheetPath(read, 'WOD');
  const xml = read(sheet.path);
  if (!xml) throw new Error('Sheet XML not found: ' + sheet.path);

  // Parse rows into sparse arrays
  const rows = [];
  const rowRe = /<row\b[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g;
  const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let rm;
  while ((rm = rowRe.exec(xml))) {
    const r = [];
    const body = rm[1] || '';
    let cm;
    cellRe.lastIndex = 0;
    while ((cm = cellRe.exec(body))) {
      const attrs = cm[1];
      const inner = cm[2] || '';
      const ref = (attrs.match(/\br="([A-Z]+)\d+"/) || [])[1];
      const t = (attrs.match(/\bt="([^"]*)"/) || [])[1];
      const ci = ref ? colIndex(ref) : r.length;
      let val = null;
      if (t === 'inlineStr') {
        let s = '';
        for (const tm of inner.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) s += tm[1];
        val = decode(s);
      } else {
        const v = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
        if (v != null) {
          if (t === 's') val = sst[+v];
          else if (t === 'str' || t === 'e') val = decode(v);
          else if (t === 'b') val = v === '1';
          else val = Number(v);
        }
      }
      r[ci] = val;
    }
    rows.push(r);
  }

  // A1 = Data updated upto
  const a1 = rows[0] ? rows[0][0] : null;
  let asOf = null;
  if (typeof a1 === 'number') asOf = serialToISO(a1);
  else if (a1 != null) asOf = String(a1).trim();

  // Header row: first row containing "Source" and "WOD Covered"
  let h = rows.findIndex(r => r && r.some(v => v === 'Source') && r.some(v => typeof v === 'string' && v.trim() === 'WOD Covered'));
  if (h < 0) throw new Error('Header row not found in sheet ' + sheet.name);
  const header = rows[h].map(v => (v == null ? '' : String(v).trim()));
  const col = name => header.findIndex(x => x.toLowerCase() === name.toLowerCase());
  const C = {
    source: col('Source'), partner: col('Partner Name'), region: col('Final Region'),
    zone: col('Final Zone'), cov: col('Coverage'), covered: col('WOD Covered')
  };
  for (const k in C) if (C[k] < 0) throw new Error('Missing column: ' + k);
  const elCols = ELEMENTS.map(e => col(e));

  const dims = { source: [], partner: [], region: [], zone: [], cov: [] };
  const idx = { source: new Map(), partner: new Map(), region: new Map(), zone: new Map(), cov: new Map() };
  const code = (k, v) => {
    const s = v == null ? '' : String(v).trim();
    let i = idx[k].get(s);
    if (i === undefined) { i = dims[k].length; dims[k].push(s); idx[k].set(s, i); }
    return i;
  };
  const out = [];
  for (let i = h + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r[C.source] == null || String(r[C.source]).trim() === '') continue;
    let mask = 0;
    elCols.forEach((c, b) => { if (c >= 0 && Number(r[c]) > 0) mask |= (1 << b); });
    const covered = Number(r[C.covered]) === 1 ? 1 : 0;
    out.push([code('source', r[C.source]), code('partner', r[C.partner]), code('region', r[C.region]),
      code('zone', r[C.zone]), code('cov', r[C.cov]), covered, mask]);
  }
  return { asOf, sheet: sheet.name, elements: ELEMENTS, dims, rows: out };
}

module.exports = { parseWOD, ELEMENTS };
