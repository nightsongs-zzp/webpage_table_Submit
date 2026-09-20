/**
 * 复现：用真实导出文件走一遍 导入解析 → computeMerge → 回读校验。
 * 用法： node tests/repro-import.js <导出的.json>
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = path.join(__dirname, '..', 'webpage-table-sync.user.js');
const src = fs.readFileSync(SRC, 'utf8');
const marker = 'if (document.readyState === \'loading\')';
const body = src.slice(0, src.indexOf(marker)) + '\n' +
  'return { parseFileToPayload, computeMerge, buildUploadTables, buildStrictTables, cellKey, normCell, CLEAR_TOKEN, normalizeServerTables, sanitizeSheetNames, expCols };\n})();';

function makeEl() {
  return {
    style: {}, dataset: {}, classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    children: [], innerHTML: '', textContent: '', value: '',
    appendChild(c) { this.children.push(c); return c; }, removeChild() {}, addEventListener() {}, removeEventListener() {},
    querySelector() { return null; }, querySelectorAll() { return []; }, setAttribute() {}, getAttribute() { return null; }, click() {},
  };
}
const sandbox = {
  console,
  document: { readyState: 'complete', head: makeEl(), body: makeEl(), createElement: () => makeEl(), getElementById: () => null, querySelector: () => null, querySelectorAll: () => [], addEventListener() {} },
  window: { location: { href: 'http://x/public/experiment.html?id=22', search: '?id=22', pathname: '/public/experiment.html', host: 'x' }, addEventListener() {}, innerWidth: 1200, innerHeight: 800 },
  location: { href: 'http://x/public/experiment.html?id=22', search: '?id=22', pathname: '/public/experiment.html', host: 'x' },
  localStorage: { _d: {}, getItem(k) { return this._d[k] === undefined ? null : this._d[k]; }, setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; } },
  URL: { createObjectURL: () => 'blob:x', revokeObjectURL() {} }, Blob: function () {},
  alert: () => {}, confirm: () => true, prompt: () => '1',
  setTimeout, clearTimeout, Date, Math, JSON, Number, String, Array, Object, RegExp, Map, Set, isNaN, parseInt, parseFloat,
  fetch: async () => { throw new Error('不该发 fetch'); },
};
sandbox.globalThis = sandbox;
sandbox.window.localStorage = sandbox.localStorage;
const fns = vm.runInNewContext(body, sandbox, { filename: 'userscript.js' });

const raw = fs.readFileSync(process.argv[2], 'utf8');
const j = JSON.parse(raw);
console.log('文件：' + process.argv[2]);
console.log('顶层 tables =', j.tables.length, ' _baseline =', j._baseline ? j._baseline.length : 'null');

// ---- 步骤 1：按页面真实逻辑构造 S.server（表定义来自接口 → 已存数据来自 GET）----
function makeServerFromExport(j, rowsStyle) {
  return {
    tables: j.tables.map((t) => ({
      id: t.id,
      name: t.name,
      headers: t.headers.slice(),
      rows: t.rows.map((r) => (rowsStyle === 'array' ? r.cells.slice() : { cells: r.cells.slice() })),
    })),
    table_data: null,
    update_time: '',
  };
}

for (const style of ['obj', 'array']) {
  console.log(`\n================ S.server.rows 形态 = ${style} ================`);
  const server = makeServerFromExport(j, style);
  let filePayload, merge;
  try {
    filePayload = fns.parseFileToPayload(raw, null, 'exp22.json');
    console.log('  ① parseFileToPayload OK：tables=' + filePayload.tables.length + ' baseline=' + (filePayload.baseline ? filePayload.baseline.size : 'null'));
  } catch (e) {
    console.log('  ✗ parseFileToPayload 抛错：' + e.constructor.name + ': ' + e.message);
    console.log(e.stack.split('\n').slice(0, 6).join('\n'));
    continue;
  }
  try {
    merge = fns.computeMerge(filePayload, server);
    console.log(`  ② computeMerge OK：changes=${merge.changes.length} stats=${JSON.stringify(merge.stats)}`);
    console.log('     warnings=' + JSON.stringify(merge.warnings));
  } catch (e) {
    console.log('  ✗ computeMerge 抛错：' + e.constructor.name + ': ' + e.message);
    console.log(e.stack.split('\n').slice(0, 8).map((l) => '      ' + l.trim()).join('\n'));
    continue;
  }
  try {
    const up = fns.buildUploadTables(server, merge.mergedCell);
    console.log('  ③ buildUploadTables OK：tables=' + up.length + ' 首表行数=' + up[0].rows.length + ' 首行列数=' + up[0].rows[0].cells.length);
    console.log('     上传体行形态：' + JSON.stringify(up[0].rows[0]));
  } catch (e) {
    console.log('  ✗ buildUploadTables 抛错：' + e.constructor.name + ': ' + e.message);
    console.log(e.stack.split('\n').slice(0, 6).map((l) => '      ' + l.trim()).join('\n'));
  }
}

// ---- 步骤 2：模拟"改了一格"后上传，检查回读校验段 ----
console.log('\n================ 改一格后走完整链路 ================');
const server = makeServerFromExport(j, 'obj');
const j2 = JSON.parse(raw);
j2.tables[0].rows[0].cells[0] = '1';
const payload2 = JSON.stringify(j2);
try {
  const fp = fns.parseFileToPayload(payload2, null, 'exp22.json');
  const m = fns.computeMerge(fp, server);
  console.log('  changes=' + m.changes.length + ' stats=' + JSON.stringify(m.stats));
  const up = fns.buildUploadTables(server, m.mergedCell);
  console.log('  上传体首格 = ' + JSON.stringify(up[0].rows[0].cells[0]) + '  （应为 "1"）');
  // 模拟回读：服务器原样存回
  const after = { tables: up.map((t) => ({ id: t.id, name: t.name, headers: t.headers, rows: t.rows })) };
  const expectCell = new Map();
  up.forEach((t) => t.rows.forEach((row, ri) => row.cells.forEach((v, ci) => expectCell.set(fns.cellKey(t.id, ri, ci), String(v)))));
  let bad = 0;
  for (const t of after.tables) t.rows.forEach((row, ri) => row.cells.forEach((v, ci) => {
    const want = expectCell.get(fns.cellKey(t.id, ri, ci));
    if (want !== undefined && String(v) !== want) bad++;
  }));
  console.log('  回读校验不一致数 = ' + bad);
} catch (e) {
  console.log('  ✗ 抛错：' + e.constructor.name + ': ' + e.message);
  console.log(e.stack.split('\n').slice(0, 8).map((l) => '    ' + l.trim()).join('\n'));
}

// ---- 步骤 3：服务器返回裸数组形态（用户实际遇到的服务器实现）+ 同一份文件 ----
console.log('\n================ 场景 C：服务器 rows 为裸数组 + 改一格 ================');
{
  const arrServer = makeServerFromExport(j, 'array');
  // 直接从 JSON 文件构造"改了一格"的文件（模拟用户编辑后回传）
  const fp = fns.parseFileToPayload(payload2, null, 'exp22.json');
  const m = fns.computeMerge(fp, arrServer);
  console.log('  changes=' + m.changes.length + ' stats=' + JSON.stringify(m.stats) + ' warnings=' + m.warnings.length);
  const up = fns.buildUploadTables(arrServer, m.mergedCell);
  console.log('  上传体：tables=' + up.length + ' 首表=' + up[0].rows.length + '行×' + up[0].rows[0].cells.length + '列');
  console.log('  首格 = ' + JSON.stringify(up[0].rows[0].cells[0]) + '（应为 "1"），行形态 = ' + JSON.stringify(up[0].rows[0]));
  // 严格覆盖也要能走
  const st2 = fns.buildStrictTables(fp, arrServer);
  console.log('  严格覆盖：tables=' + st2.length + ' 首表列数=' + st2[0].headers.length);
}

// ---- 步骤 4：文件行是裸数组（手工 JSON）也不能崩 ----
console.log('\n================ 场景 D：文件行也是裸数组 ================');
{
  const arrFile = { tables: [{ id: 1, name: 't', headers: ['a', 'b'], rows: [['x', 'y']] }], baseline: null };
  const arrServer = makeServerFromExport(j, 'array');
  try {
    const m = fns.computeMerge(arrFile, arrServer);
    console.log('  OK：changes=' + m.changes.length + ' stats=' + JSON.stringify(m.stats));
  } catch (e) {
    console.log('  ✗ 抛错：' + e.constructor.name + ': ' + e.message);
  }
}
