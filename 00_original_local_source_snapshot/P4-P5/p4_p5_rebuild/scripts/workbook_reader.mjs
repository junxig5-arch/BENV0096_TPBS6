import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

function findEocd(buf) {
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65558); i -= 1) {
    if (buf.readUInt32LE(i) === 0x06054b50) return i;
  }
  throw new Error('ZIP end-of-central-directory not found');
}

export function zipEntries(filePath) {
  const buf = fs.readFileSync(filePath);
  const eocd = findEocd(buf);
  const total = buf.readUInt16LE(eocd + 10);
  const cdOffset = buf.readUInt32LE(eocd + 16);
  const out = new Map();
  let p = cdOffset;
  for (let i = 0; i < total; i += 1) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('Bad ZIP central directory at ' + p);
    const method = buf.readUInt16LE(p + 10);
    const compressedSize = buf.readUInt32LE(p + 20);
    const fileNameLength = buf.readUInt16LE(p + 28);
    const extraLength = buf.readUInt16LE(p + 30);
    const commentLength = buf.readUInt16LE(p + 32);
    const localOffset = buf.readUInt32LE(p + 42);
    const name = buf.slice(p + 46, p + 46 + fileNameLength).toString('utf8');
    if (buf.readUInt32LE(localOffset) !== 0x04034b50) throw new Error('Bad ZIP local header for ' + name);
    const localNameLength = buf.readUInt16LE(localOffset + 26);
    const localExtraLength = buf.readUInt16LE(localOffset + 28);
    const start = localOffset + 30 + localNameLength + localExtraLength;
    const data = buf.slice(start, start + compressedSize);
    let inflated;
    if (method === 0) inflated = data;
    else if (method === 8) inflated = zlib.inflateRawSync(data);
    else inflated = null;
    out.set(name, inflated);
    p += 46 + fileNameLength + extraLength + commentLength;
  }
  return out;
}

export function xmlDecode(value = '') {
  return String(value).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

function attr(xml, name) {
  const re = new RegExp(name.replace(':', ':') + '="([^"]*)"');
  const m = xml.match(re);
  return m ? xmlDecode(m[1]) : undefined;
}

function extractText(xml, tagNames = ['t']) {
  const tags = tagNames.join('|');
  const re = new RegExp('<[^:>]*:?(' + tags + ')\\b[^>]*>([\\s\\S]*?)<\\/[^:>]*:?\\1>', 'g');
  const parts = [];
  let m;
  while ((m = re.exec(xml))) parts.push(xmlDecode(m[2]).trim());
  return parts.filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
}

function colToIndex(ref) {
  const letters = String(ref || '').match(/[A-Z]+/i)?.[0]?.toUpperCase() || 'A';
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function sheetPathFromTarget(target) {
  if (!target) return undefined;
  if (target.startsWith('/')) return target.slice(1);
  return path.posix.normalize(path.posix.join('xl', target));
}

function xlsxSharedStrings(entries) {
  const raw = entries.get('xl/sharedStrings.xml');
  if (!raw) return [];
  const xml = raw.toString('utf8');
  const strings = [];
  const re = /<si\b[\s\S]*?<\/si>/g;
  let m;
  while ((m = re.exec(xml))) strings.push(extractText(m[0], ['t']));
  return strings;
}

export function xlsxSheetNames(filePath) {
  const entries = zipEntries(filePath);
  const workbook = entries.get('xl/workbook.xml')?.toString('utf8');
  const rels = entries.get('xl/_rels/workbook.xml.rels')?.toString('utf8');
  if (!workbook) throw new Error('xl/workbook.xml not found');
  const relMap = new Map();
  if (rels) {
    const relRe = /<Relationship\b[^>]*>/g;
    let rm;
    while ((rm = relRe.exec(rels))) relMap.set(attr(rm[0], 'Id'), attr(rm[0], 'Target'));
  }
  const sheets = [];
  const sheetRe = /<sheet\b[^>]*>/g;
  let sm;
  while ((sm = sheetRe.exec(workbook))) {
    const node = sm[0];
    const relId = attr(node, 'r:id') || attr(node, 'id');
    sheets.push({ name: attr(node, 'name'), id: attr(node, 'sheetId'), relId, path: sheetPathFromTarget(relMap.get(relId)) });
  }
  return sheets;
}

export function readXlsxSheet(filePath, sheetName, options = {}) {
  const maxRows = options.maxRows ?? Infinity;
  const entries = zipEntries(filePath);
  const sheets = xlsxSheetNames(filePath);
  const sheet = sheets.find((s) => s.name === sheetName) || sheets[0];
  if (!sheet?.path) throw new Error('Sheet not found: ' + sheetName);
  const xml = entries.get(sheet.path)?.toString('utf8');
  if (!xml) throw new Error('Sheet XML not found: ' + sheet.path);
  const shared = xlsxSharedStrings(entries);
  const out = [];
  const rowRe = /<row\b[\s\S]*?<\/row>/g;
  let rowMatch;
  while ((rowMatch = rowRe.exec(xml)) && out.length < maxRows) {
    const row = [];
    const cellRe = /<c\b[\s\S]*?<\/c>/g;
    let cellMatch;
    while ((cellMatch = cellRe.exec(rowMatch[0]))) {
      const cxml = cellMatch[0];
      const ref = attr(cxml, 'r');
      const idx = ref ? colToIndex(ref) : row.length;
      const type = attr(cxml, 't');
      let value = '';
      if (type === 'inlineStr') value = extractText(cxml, ['t']);
      else {
        const v = cxml.match(/<v\b[^>]*>([\s\S]*?)<\/v>/)?.[1];
        if (v !== undefined) {
          const decoded = xmlDecode(v);
          if (type === 's') value = shared[Number(decoded)] ?? decoded;
          else if (type === 'str') value = decoded;
          else if (type === 'b') value = decoded === '1';
          else if (decoded !== '' && !Number.isNaN(Number(decoded))) value = Number(decoded);
          else value = decoded;
        }
      }
      row[idx] = value;
    }
    while (row.length && (row[row.length - 1] === undefined || row[row.length - 1] === '')) row.pop();
    out.push(row.map((v) => (v === undefined ? '' : v)));
  }
  return out;
}

export function odsSheetNames(filePath) {
  const entries = zipEntries(filePath);
  const content = entries.get('content.xml')?.toString('utf8');
  if (!content) throw new Error('content.xml not found');
  const sheets = [];
  const tableRe = /<table:table(?:\s|>)[^>]*>/g;
  let m;
  while ((m = tableRe.exec(content))) sheets.push({ name: attr(m[0], 'table:name') });
  return sheets;
}

function odsCellValue(cellXml) {
  const valueType = attr(cellXml, 'office:value-type');
  const raw = attr(cellXml, 'office:value') ?? attr(cellXml, 'office:date-value') ?? attr(cellXml, 'office:time-value') ?? attr(cellXml, 'office:boolean-value');
  if (raw !== undefined && valueType !== 'string') {
    const numeric = Number(raw);
    return Number.isNaN(numeric) ? raw : numeric;
  }
  return extractText(cellXml, ['p', 'span']).replace(/\s+/g, ' ').trim();
}

export function readOdsSheet(filePath, sheetName, options = {}) {
  const maxRows = options.maxRows ?? Infinity;
  const maxCols = options.maxCols ?? 500;
  const entries = zipEntries(filePath);
  const content = entries.get('content.xml')?.toString('utf8');
  if (!content) throw new Error('content.xml not found');
  const tableRe = /<table:table(?:\s|>)[\s\S]*?<\/table:table>/g;
  let tableXml;
  let tm;
  while ((tm = tableRe.exec(content))) {
    const head = tm[0].slice(0, tm[0].indexOf('>') + 1);
    const name = attr(head, 'table:name');
    if (!sheetName || name === sheetName) { tableXml = tm[0]; break; }
  }
  if (!tableXml) throw new Error('ODS sheet not found: ' + sheetName);
  const out = [];
  const rowRe = /<table:table-row\b[\s\S]*?<\/table:table-row>/g;
  let rm;
  while ((rm = rowRe.exec(tableXml)) && out.length < maxRows) {
    const repeatRows = Math.min(Number(attr(rm[0], 'table:number-rows-repeated') || 1), Math.max(1, maxRows - out.length));
    const row = [];
    const cellRe = /<table:table-cell\b[\s\S]*?<\/table:table-cell>|<table:covered-table-cell\b[^>]*\/>/g;
    let cm;
    while ((cm = cellRe.exec(rm[0])) && row.length < maxCols) {
      const cxml = cm[0];
      const repeated = Math.min(Number(attr(cxml, 'table:number-columns-repeated') || 1), maxCols - row.length);
      const value = cxml.startsWith('<table:covered') ? '' : odsCellValue(cxml);
      for (let i = 0; i < repeated; i += 1) row.push(value);
    }
    while (row.length && (row[row.length - 1] === undefined || row[row.length - 1] === '')) row.pop();
    if (row.length) for (let i = 0; i < repeatRows; i += 1) out.push([...row]);
  }
  return out;
}

export function sheetNames(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.xlsx' || ext === '.xlsm') return xlsxSheetNames(filePath).map((s) => s.name);
  if (ext === '.ods') return odsSheetNames(filePath).map((s) => s.name);
  throw new Error('Unsupported file type for sheet names: ' + ext);
}

export function readSheet(filePath, sheetName, options = {}) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.xlsx' || ext === '.xlsm') return readXlsxSheet(filePath, sheetName, options);
  if (ext === '.ods') return readOdsSheet(filePath, sheetName, options);
  throw new Error('Unsupported file type for reading sheet rows: ' + ext);
}
