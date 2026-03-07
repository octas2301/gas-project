/**
 * 楽天CSVの「バリエーション1選択肢定義」が32バイト以下か検証する。
 * コード.js の truncateToByteLength と同じ数え方: charCode < 128 → 1バイト、それ以外 → 2バイト
 * 使い方: node check-rakuten-32byte.js
 */
const fs = require('fs');
const path = require('path');

function byteLength(str) {
  if (str == null || str === '') return 0;
  const s = String(str);
  let bytes = 0;
  for (let i = 0; i < s.length; i++) {
    bytes += s.charCodeAt(i) < 128 ? 1 : 2;
  }
  return bytes;
}

function parseCsvLine(line) {
  const fields = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') {
      inQuotes = !inQuotes;
    } else if ((c === ',' && !inQuotes) || (c === '\n' && !inQuotes)) {
      fields.push(cur);
      cur = '';
      if (c === '\n') break;
    } else {
      cur += c;
    }
  }
  if (cur !== '' || line[line.length - 1] === ',') fields.push(cur);
  return fields;
}

const csvDir = path.join(__dirname, '楽天CSVファイル');
const logDir = path.join(__dirname, '楽天ログファイル');
const COL_NAME = 'バリエーション1選択肢定義';
const MAX_BYTES = 32;

const filesToCheck = [];
if (fs.existsSync(csvDir)) {
  const mainCsv = path.join(csvDir, 'normal-item.csv');
  if (fs.existsSync(mainCsv)) filesToCheck.push({ path: mainCsv, name: '楽天CSVファイル/normal-item.csv' });
}
if (fs.existsSync(logDir)) {
  const names = fs.readdirSync(logDir).filter(n => n.endsWith('.csv'));
  names.forEach(n => filesToCheck.push({ path: path.join(logDir, n), name: '楽天ログファイル/' + n }));
}

if (filesToCheck.length === 0) {
  console.log('CSVファイルが見つかりません。');
  process.exit(1);
}

let totalOver = 0;
let totalRows = 0;
const results = [];

for (const { path: filePath, name } of filesToCheck) {
  let content;
  try {
    content = fs.readFileSync(filePath, 'utf8');
  } catch (e) {
    try {
      content = fs.readFileSync(filePath, 'shiftjis').toString();
    } catch (e2) {
      console.warn('スキップ（読めません）:', name);
      continue;
    }
  }
  const lines = content.split(/\r?\n/).filter(l => l.trim() !== '');
  if (lines.length < 2) continue;
  const headerFields = parseCsvLine(lines[0]);
  const idx = headerFields.findIndex(h => h.trim() === COL_NAME);
  if (idx === -1) {
    results.push({ file: name, note: 'ヘッダーに「バリエーション1選択肢定義」なし', over: [], overCount: 0, dataRows: 0 });
    continue;
  }
  const over = [];
  let dataRows = 0;
  for (let r = 1; r < lines.length; r++) {
    const fields = parseCsvLine(lines[r]);
    if (fields.length <= idx) continue;
    dataRows++;
    const val = (fields[idx] || '').trim();
    if (val === '') continue;
    const bytes = byteLength(val);
    if (bytes > MAX_BYTES) {
      const itemUrl = fields[0] || ('行' + (r + 1));
      over.push({ row: r + 1, itemUrl, value: val.substring(0, 50) + (val.length > 50 ? '…' : ''), bytes });
    }
  }
  totalRows += dataRows;
  totalOver += over.length;
  results.push({ file: name, over, overCount: over.length, dataRows });
}

console.log('=== 楽天「バリエーション1選択肢定義」32バイト超過チェック ===\n');
console.log('数え方: コード.js の truncateToByteLength と同じ（ASCII=1バイト、それ以外=2バイト）\n');

for (const r of results) {
  console.log('ファイル:', r.file);
  if (r.note) {
    console.log('  ', r.note);
    continue;
  }
  console.log('  データ行数:', r.dataRows, '  32バイト超過:', r.overCount);
  if (r.over.length > 0) {
    r.over.forEach(o => {
      console.log('    行' + o.row + '  ' + o.itemUrl + '  バイト数=' + o.bytes + '  値="' + o.value + '"');
    });
  }
  console.log('');
}

console.log('--- まとめ ---');
console.log('対象データ行合計:', totalRows);
console.log('32バイト超過の行数:', totalOver);
if (totalOver === 0 && totalRows > 0) {
  console.log('→ すべて32バイト以下でした。');
} else if (totalOver > 0) {
  console.log('→ 上記の行が楽天の32バイト制限を超えています。');
}
