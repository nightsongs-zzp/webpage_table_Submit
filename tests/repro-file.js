'use strict';
/**
 * 用脚本自己的解析 + 合并逻辑，跑一个真实导出的 xlsx，定位"上传不了"的原因。
 * 用法： node tests/repro-file.js <文件.xlsx>
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const XLSX = require('xlsx');

const SRC = path.join(__dirname, '..', 'webpage-table-sync.user.js');
const src = fs.readFileSync(SRC, 'utf8');
const marker = "if (document.readyState === 'loading')";
const body = src.slice(0, src.indexOf(marker)) +
  '\nreturn { parseFileToPayload, computeMerge, buildUploadTables, cellKey, PRESET_SHEET, BASE_SHEET, BORDER_MARK, seedServerTablesFromDefs, S };\n})();';

function el() {
  return {
    style: {}, dataset: {}, classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    children: [], innerHTML: '', textContent: '', value: '',
    appendChild(c) { return c; }, removeChild() {}, addEventListener() {}, removeEventListener() {},
    querySelector() { return null; }, querySelectorAll() { return []; }, setAttribute() {}, getAttribute() { return null; }, click() {},
  };
}
const sandbox = {
  console, XLSX,
  document: { readyState: 'complete', head: el(), body: el(), createElement: () => el(), getElementById: () => null, querySelector: () => null, querySelectorAll: () => [], addEventListener() {} },
  window: { location: { href: 'http://x/public/experiment.html?id=99', search: '?id=99', pathname: '/public/experiment.html', host: 'x' }, addEventListener() {}, innerWidth: 1, innerHeight: 1 },
  location: { href: 'http://x', search: '?id=99', pathname: '/p', host: 'x' },
  localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
  URL: { createObjectURL: () => '', revokeObjectURL() {} }, Blob: function () {},
  alert() {}, confirm() { return true; }, prompt() { return '1'; },
  setTimeout, clearTimeout, Date, Math, JSON, Number, String, Array, Object, RegExp, Map, Set, isNaN, parseInt, parseFloat,
  fetch: async () => { throw new Error('no'); },
};
sandbox.globalThis = sandbox; sandbox.window.localStorage = sandbox.localStorage;
const fns = vm.runInNewContext(body, sandbox, { filename: 'u.js' });

const file = process.argv[2];
const buf = fs.readFileSync(file);
console.log('文件:', path.basename(file), buf.length, 'B');

// ---------- 1. 解析 ----------
let fp;
try {
  fp = fns.parseFileToPayload(null, buf, path.basename(file));
} catch (e) {
  console.log('✗ parseFileToPayload 抛错:', e.constructor.name, e.message);
  process.exit(1);
}
console.log('\n=== 1. 解析结果 ===');
console.log('  表数:', fp.tables.length);
fp.tables.forEach((t) => console.log(`    id=${t.id}  ${t.rows.length}行 × ${t.headers.length}列  表头[0..2]=${JSON.stringify(t.headers.slice(0, 3))}`));
console.log('  baseline:', fp.baseline ? fp.baseline.size + ' 格' : 'null');
console.log('  presets :', fp.presets ? fp.presets.size + ' 格' : 'null');
if (fp.warnings.length) { console.log('  告警:'); fp.warnings.forEach((w) => console.log('    - ' + w)); }

// ---------- 2. 用 _基线 还原"服务器已存数据" ----------
function serverFromBaseline(fp) {
  const tables = fp.tables.map((t) => {
    const rows = t.rows.map((r, ri) => {
      const cells = [];
      for (let ci = 0; ci < t.headers.length; ci++) {
        const k = fns.cellKey(t.id, ri, ci);
        const v = fp.baseline ? fp.baseline.get(k) : undefined;
        cells.push(v === undefined ? '' : v);
      }
      return { cells };
    });
    return { id: t.id, name: '', headers: t.headers.slice(), rows };
  });
  return { tables, update_time: '2026-10-15 23:14:58', table_data: null };
}
const server = serverFromBaseline(fp);
const baseIds = server.tables.map((t) => String(t.id));
console.log('\n=== 2. 用 _基线 还原的服务器表 id ===');
console.log('  ' + baseIds.join(','));

// ---------- 3. 未改动 → 应 0 改动 ----------
let m;
try {
  m = fns.computeMerge(fp, server);
} catch (e) {
  console.log('\n✗ computeMerge 抛错:', e.constructor.name, e.message);
  console.log(e.stack.split('\n').slice(0, 6).join('\n'));
  process.exit(1);
}
console.log('\n=== 3. 未改动直接上传 ===');
console.log('  changes:', m.changes.length, ' stats:', JSON.stringify(m.stats));
if (m.warnings.length) { console.log('  合并告警:'); m.warnings.forEach((w) => console.log('    - ' + w)); }
console.log('  前 12 处改动明细:');
m.changes.slice(0, 12).forEach((c) => {
  console.log(`    [${c.kind}] 表${c.t} (${c.r},${c.c})  ${JSON.stringify(c.from)} → ${JSON.stringify(c.to)}`);
});

// ---------- 4. 改一格再上传 ----------
console.log('\n=== 4. 改一格后上传 ===');
const fp2 = fns.parseFileToPayload(null, buf, path.basename(file));
const t0 = fp2.tables[0];
let changed = 0;
for (let r = 0; r < t0.rows.length && changed < 2; r++) {
  for (let c = 0; c < t0.rows[r].cells.length && changed < 2; c++) {
    const cur = t0.rows[r].cells[c];
    if (!cur) continue;
    t0.rows[r].cells[c] = String(cur) + 'X';
    console.log(`    改 id=${t0.id} (${r},${c}): "${cur}" → "${cur}X"`);
    changed++;
  }
}
const m2 = fns.computeMerge(fp2, server);
console.log('  changes:', m2.changes.length, ' stats:', JSON.stringify(m2.stats));
let up;
try {
  up = fns.buildUploadTables(server, m2.mergedCell);
  console.log('  上传体: tables=' + up.length + ' 各表行数=' + up.map((t) => t.rows.length).join(','));
} catch (e) {
  console.log('  ✗ buildUploadTables 抛错:', e.constructor.name, e.message);
  console.log(e.stack.split('\n').slice(0, 6).join('\n'));
}
// 回读校验模拟
if (up) {
  const expect = new Map();
  up.forEach((t) => t.rows.forEach((row, ri) => row.cells.forEach((v, ci) => expect.set(fns.cellKey(t.id, ri, ci), String(v)))));
  let bad = 0;
  server.tables.forEach((t) => t.rows.forEach((row, ri) => row.cells.forEach((v, ci) => {
    const k = fns.cellKey(t.id, ri, ci);
    if (expect.has(k) && String(v) !== expect.get(k)) bad++;
  })));
  console.log('  模拟回读不一致数:', bad);
}

console.log('\n=== 6. 【真实场景】服务器还没有已存数据（exp28 首次上传）===');
{
  // 模板定义 = 文件里的表（id/表头/行数），和服务器上 dataTableDefinition 一致
  const defs = fp.tables.map((t) => ({ id: t.id, name: t.name, headers: t.headers.slice(), expectedRows: t.rows.length }));
  const tables = fns.seedServerTablesFromDefs([], defs);
  console.log('  补出的表:', tables.map((t) => `id=${t.id}(${t.rows.length}行×${t.headers.length}列)`).join(' '));
  const srv = { tables, update_time: '', table_data: null };
  const m = fns.computeMerge(fp, srv);
  console.log('  changes:', m.changes.length, ' stats:', JSON.stringify(m.stats));
  const up = fns.buildUploadTables(srv, m.mergedCell, defs);
  console.log('  上传体: tables=' + up.length + ' 各表行数=' + up.map((t) => t.rows.length).join(','));
  // 找出用户真实填写的数值（表7 第1行第3列 = 24.706）
  const t7 = up.find((t) => String(t.id) === '7');
  console.log('  表7 第1行:', JSON.stringify(t7.rows[0].cells));
  console.log('  → 用户填的数值是否进入上传体:', t7.rows[0].cells.includes('24.706'));
}

// ---------- 5. 关键怀疑：服务器表 id 与文件表 id 不匹配时会发生什么 ----------
console.log('\n=== 5. 如果服务器表 id 是 1..8（与文件的 7,8,9,13… 不同）===');
const wrongServer = { tables: server.tables.map((t, i) => Object.assign({}, t, { id: i + 1 })), update_time: 't', table_data: null };
const m3 = fns.computeMerge(fp, wrongServer);
console.log('  changes:', m3.changes.length, ' stats:', JSON.stringify(m3.stats));
m3.warnings.slice(0, 10).forEach((w) => console.log('    - ' + w));
console.log('  → 这种情况下上传会是「无改动，未上传」，用户会觉得"点确认没反应"');
