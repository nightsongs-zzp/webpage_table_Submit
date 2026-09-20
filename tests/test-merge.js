/**
 * 逻辑自测：把 webpage-table-sync.user.js 的纯函数抽出来跑（不连网、不动服务器）。
 * 用法： node tests/test-merge.js
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = path.join(__dirname, '..', 'webpage-table-sync.user.js');
const src = fs.readFileSync(SRC, 'utf8');
const marker = 'if (document.readyState === \'loading\')';
const cut = src.indexOf(marker);
if (cut < 0) { console.error('找不到入口标记，脚本结构变了？'); process.exit(1); }
const body = src.slice(0, cut) + '\n' +
  'return { computeMerge, buildUploadTables, buildStrictTables, parseFileToPayload, cellKey, ' +
  'buildExportView, expCols, normCell, rowCells, mapLike, CLEAR_TOKEN, sanitizeSheetNames, normalizeServerTables };\n})();';

// ---- 最小 DOM / window 桩 ----
function makeEl() {
  const el = {
    style: {}, dataset: {}, classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    children: [], innerHTML: '', textContent: '', value: '',
    appendChild(c) { this.children.push(c); return c; }, removeChild() {}, addEventListener() {},
    removeEventListener() {}, querySelector() { return null; }, querySelectorAll() { return []; },
    setAttribute() {}, getAttribute() { return null; }, click() {},
  };
  return el;
}
const documentStub = {
  readyState: 'complete',
  head: makeEl(), body: makeEl(),
  createElement: () => makeEl(),
  getElementById: () => null,
  querySelector: () => null,
  querySelectorAll: () => [],
  addEventListener() {},
};
const sandbox = {
  console,
  document: documentStub,
  window: { location: { href: 'http://x/public/experiment.html?id=33', search: '?id=33', pathname: '/public/experiment.html', host: 'x' }, addEventListener() {}, innerWidth: 1200, innerHeight: 800 },
  location: { href: 'http://x/public/experiment.html?id=33', search: '?id=33', pathname: '/public/experiment.html', host: 'x' },
  localStorage: { _d: {}, getItem(k) { return this._d[k] === undefined ? null : this._d[k]; }, setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; } },
  URL: { createObjectURL: () => 'blob:x', revokeObjectURL() {} },
  Blob: function () {}, alert: () => {}, confirm: () => true, prompt: () => '1',
  setTimeout, clearTimeout, Date, Math, JSON, Number, String, Array, Object, RegExp, Map, Set, isNaN, parseInt, parseFloat,
  fetch: async () => { throw new Error('测试环境不应发起 fetch'); },
};
sandbox.globalThis = sandbox;
sandbox.window.localStorage = sandbox.localStorage;

const fns = vm.runInNewContext(body, sandbox, { filename: 'userscript.js' });

// ---- 断言小工具 ----
let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log('  ✗ ' + name + (extra ? '   → ' + extra : '')); }
}
function eq(name, actual, expect) {
  ok(name, JSON.stringify(actual) === JSON.stringify(expect),
    `actual=${JSON.stringify(actual)} expect=${JSON.stringify(expect)}`);
}

const mkTable = (id, headers, rows, name) => ({ id, name: name || ('表' + id), headers, rows: rows.map((r) => ({ cells: r.slice() })) });
const mkServer = (tables, updateTime) => ({ tables, update_time: updateTime || '2026-09-19 15:01:11', table_data: null });
const mkFile = (tables, baseline) => {
  const bm = new Map();
  if (baseline) baseline.forEach((t) => t.rows.forEach((r, ri) => r.cells.forEach((c, ci) => bm.set(fns.cellKey(t.id, ri, ci), c))));
  return { tables, baseline: baseline ? bm : null };
};

console.log('\n=== 1. 三方合并：没动的格不上传 ===');
{
  const server = mkServer([
    mkTable(0, ['a', 'b'], [['1', '2'], ['3', '4']]),
    mkTable(1, ['x'], [['9']]),
  ]);
  const file = mkFile([
    mkTable(0, ['a', 'b'], [['1', '2'], ['3', '4']]),   // 与服务器一致
    mkTable(1, ['x'], [['9']]),
  ], [mkTable(0, ['a', 'b'], [['1', '2'], ['3', '4']]), mkTable(1, ['x'], [['9']])]);
  const m = fns.computeMerge(file, server);
  eq('无改动 → changes 为空', m.changes.length, 0);
  eq('stats.mod/add/del 全 0', [m.stats.mod, m.stats.add, m.stats.del], [0, 0, 0]);
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('上传体保留服务器全部 2 张表', up.length, 2);
  eq('上传体内容与服务器完全一致', up.map((t) => t.rows), server.tables.map((t) => t.rows));
}

console.log('\n=== 2. 三方合并：改 / 新增 / 清空 ===');
{
  const server = mkServer([mkTable(0, ['a', 'b'], [['1', '2'], ['3', '4']])]);
  const base = [mkTable(0, ['a', 'b'], [['1', '2'], ['3', '4']])];
  const file = mkFile([mkTable(0, ['a', 'b'], [['1', 'X'], ['', '4']])], base);
  const m = fns.computeMerge(file, server);
  eq('识别 2 处改动', m.changes.length, 2);
  eq('1 改 1 清空', [m.stats.mod, m.stats.add, m.stats.del], [1, 0, 1]);
  const byKey = Object.fromEntries(m.changes.map((c) => [fns.cellKey(c.t, c.r, c.c), c]));
  eq('(0,1) 2→X 为 mod', [byKey['0|0|1'].kind, byKey['0|0|1'].from, byKey['0|0|1'].to], ['mod', '2', 'X']);
  eq('(1,0) 3→空 为 del', [byKey['0|1|0'].kind, byKey['0|1|0'].to], ['del', '']);
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('上传后 (0,1)=X', up[0].rows[0].cells[1], 'X');
  eq('上传后 (1,0) 被清空', up[0].rows[1].cells[0], '');
  eq('未改动的 (1,1) 保持 4', up[0].rows[1].cells[1], '4');
}

console.log('\n=== 3. 显式清空标记 __CLEAR__ ===');
{
  const server = mkServer([mkTable(0, ['a'], [['keep'], ['wipe']])]);
  const file = mkFile([mkTable(0, ['a'], [['keep'], [fns.CLEAR_TOKEN]])], [mkTable(0, ['a'], [['keep'], ['wipe']])]);
  const m = fns.computeMerge(file, server);
  eq('只产生一处 del', [m.changes.length, m.stats.del], [1, 1]);
  eq('被清空单元格不含 __CLEAR__ 字样', m.mergedCell.get('0|1|0'), '');
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('上传体：第 1 行保留、第 2 行清空', up[0].rows.map((r) => r.cells[0]), ['keep', '']);
}

console.log('\n=== 4. 无基线（手工 Excel）：绝不覆盖服务器已有值 ===');
{
  const server = mkServer([mkTable(0, ['a', 'b'], [['sv1', ''], ['sv2', '']])]);
  const file = mkFile([mkTable(0, ['a', 'b'], [['my1', 'new1'], ['sv2', 'new2']])], null);
  const m = fns.computeMerge(file, server);
  eq('只新增 2 格（行1列0 冲突不算改动），0 覆盖 0 清空', [m.stats.add, m.stats.mod, m.stats.del], [2, 0, 0]);
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('冲突格保留服务器值 sv1', up[0].rows[0].cells[0], 'sv1');
  eq('空缺格写入 new1', up[0].rows[0].cells[1], 'new1');
  eq('产生冲突告警', m.warnings.some((w) => w.includes('冲突')), true);
  eq('上传体不会把空值抹掉', up[0].rows[1].cells[0], 'sv2');
}

console.log('\n=== 5. 文件里没有的表 / 服务器上没有的表 ===');
{
  const server = mkServer([mkTable(0, ['a'], [['s0']]), mkTable(1, ['b'], [['s1']])]);
  const file = mkFile([mkTable(0, ['a'], [['f0']]), mkTable(9, ['z'], [['zz']])], [mkTable(0, ['a'], [['s0']]), mkTable(9, ['z'], [['']])]);
  const m = fns.computeMerge(file, server);
  eq('只动表 0', m.changes.length, 1);
  eq('表 9 被跳过并告警', m.warnings.some((w) => w.includes('id=9')), true);
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('上传体仍含服务器表 1 且值不变', [up.length, up[1].rows[0].cells[0]], [2, 's1']);
}

console.log('\n=== 6. 严格覆盖模式 ===');
{
  const server = mkServer([mkTable(0, ['a', 'b'], [['s1', 's2']])]);
  const file = mkFile([mkTable(0, ['a', 'b'], [['new', '']])], [mkTable(0, ['a', 'b'], [['s1', 's2']])]);
  const strict = fns.buildStrictTables(file, server);
  eq('严格模式：空即清空', strict[0].rows[0].cells, ['new', '']);
  eq('严格模式：__CLEAR__ 也被解析为空', fns.buildStrictTables(mkFile([mkTable(0, ['a'], [[fns.CLEAR_TOKEN]])], null), server)[0].rows[0].cells, ['']);
}

console.log('\n=== 7. 行列扩展（文件比服务器多行多列） ===');
{
  const server = mkServer([mkTable(0, ['a'], [['1']])]);
  const file = mkFile([mkTable(0, ['a', 'b'], [['1', '2'], ['', '3']])], [mkTable(0, ['a'], [['1']])]);
  const m = fns.computeMerge(file, server);
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('新增 (0,1)=2', up[0].rows[0].cells[1], '2');
  eq('新增行 (1,1)=3', up[0].rows[1].cells[1], '3');
  eq('服务器原值 (0,0) 保留 1', up[0].rows[0].cells[0], '1');
  eq('行数扩展到 2', up[0].rows.length, 2);
  eq('新增 (1,0) 为空字符串而非 undefined', up[0].rows[1].cells[0], '');
}

console.log('\n=== 8. JSON 往返（导出 → 解析） ===');
{
  const payload = {
    file: { app: 'webpage_table_sync/v1', expId: '33' },
    tables: [
      { id: 0, name: '表A', headers: ['h $I_f$', 'b'], rows: [{ cells: ['$1.5\\times10^{-4}$', ''] }, { cells: ['x|y,z', ' 空格保留 '] }] },
      { id: 1, name: '表B', headers: ['c'], rows: [{ cells: ['v'] }] },
    ],
    _baseline: [
      { id: 0, name: '表A', headers: ['h $I_f$', 'b'], rows: [{ cells: ['old', ''] }, { cells: ['', ''] }] },
      { id: 1, name: '表B', headers: ['c'], rows: [{ cells: ['v'] }] },
    ],
    _clearToken: '__CLEAR__',
  };
  const parsed = fns.parseFileToPayload(JSON.stringify(payload), null, 'exp33.json');
  eq('解析出 2 张表', parsed.tables.length, 2);
  eq('LaTeX 原文无损', parsed.tables[0].rows[0].cells[0], '$1.5\\times10^{-4}$');
  eq('含 | 与 , 的单元格无损', parsed.tables[0].rows[1].cells[0], 'x|y,z');
  eq('前后空格保留', parsed.tables[0].rows[1].cells[1], ' 空格保留 ');
  eq('带出基线', parsed.baseline.size, 5);
  const server = mkServer([mkTable(0, ['h $I_f$', 'b'], [['old', '']]), mkTable(1, ['c'], [['v']])]);
  const m = fns.computeMerge(parsed, server);
  eq('合并改 3 格：表A (0,0) 改公式、(1,0) 补 x|y,z、(1,1) 补「 空格保留 」', m.changes.length, 3);
  eq('1 处修改 + 2 处新增，无清空', [m.stats.mod, m.stats.add, m.stats.del], [1, 2, 0]);
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('含 | 与 , 的值原样上传', up[0].rows[1].cells[0], 'x|y,z');
  eq('前后空格原样上传', up[0].rows[1].cells[1], ' 空格保留 ');
  eq('未改动的表 1 保持 v', up[1].rows[0].cells[0], 'v');
}

console.log('\n=== 9. 导出视图：表定义 + 服务器数据合体 ===');
{
  const savedS = sandbox.__S;
  // 用 vm 里的 S 不方便，这里直接验证 expCols / normCell 两个辅助
  eq('expCols 兼容空行表', fns.expCols({ headers: ['a', 'b'], rows: [] }), 2);
  eq('expCols 取最宽行', fns.expCols({ headers: ['a'], rows: [{ cells: ['1', '2', '3'] }] }), 3);
  eq('normCell(number 0.00015)', fns.normCell(0.00015), '0.00015');
  eq('normCell(null)', fns.normCell(null), '');
  eq('normCell(整数)', fns.normCell(1800), '1800');
}

console.log('\n=== 10. 工作表名净化 ===');
{
  const names = fns.sanitizeSheetNames([
    { id: 0, name: 'T0_第1张/表:测试[1]' },
    { id: 1, name: 'T1_第1张/表:测试[1]' },
    { id: 2, name: 'T2_' + 'x'.repeat(60) },
  ]);
  eq('非法字符被替换', /[\[\]\*\/\\\?:]/.test(names[0].sheet), false);
  eq('长度 ≤ 31', names.every((n) => n.sheet.length <= 31), true);
  eq('名称不重复', new Set(names.map((n) => n.sheet)).size, 3);
  eq('保留 T{id}_ 前缀（导入靠它认 id）', names.map((n) => /^T\d+_/.test(n.sheet)), [true, true, true]);
}

console.log('\n=== 11. 回归：服务器 rows 是「裸数组」形态（v1.0.0 崩溃点） ===');
{
  // 这正是 exp22 崩溃的形态：rows: [ ["1","2"], ["3","4"] ] —— 没有 .cells
  const server = mkServer([
    { id: 0, name: '表0', headers: ['a', 'b'], rows: [['1', '2'], ['3', '4']] },
  ]);
  eq('rowCells 兼容裸数组行', fns.rowCells(['1', '2']), ['1', '2']);
  eq('rowCells 兼容对象行', fns.rowCells({ cells: ['1'] }), ['1']);
  eq('rowCells 兼容垃圾输入', fns.rowCells(null), []);
  eq('expCols 兼容裸数组行（旧版在此崩）', fns.expCols(server.tables[0]), 2);

  const file = mkFile([mkTable(0, ['a', 'b'], [['1', 'X'], ['3', '4']])], [mkTable(0, ['a', 'b'], [['1', '2'], ['3', '4']])]);
  let m;
  try { m = fns.computeMerge(file, server); } catch (e) { m = null; }
  ok('computeMerge 不再抛错', !!m, m ? '' : '仍然抛错');
  if (m) {
    eq('识别 1 处改动', m.changes.length, 1);
    const up = fns.buildUploadTables(server, m.mergedCell);
    eq('上传体写回对象形态 {cells:[...]}', Array.isArray(up[0].rows[0].cells), true);
    eq('改动生效', up[0].rows[0].cells[1], 'X');
    eq('未改动保留', up[0].rows[0].cells[0], '1');
  }

  // normalizeServerTables：接口不管给哪种行形态，出口必须统一成 {cells:[]}
  const norm = fns.normalizeServerTables({ tables: [{ id: 3, name: 'n', headers: ['h'], rows: [['x'], { cells: ['y'] }] }] });
  eq('normalizeServerTables 出口统一', norm[0].rows, [{ cells: ['x'] }, { cells: ['y'] }]);
  eq('normalizeServerTables 出口 expCols 可用', fns.expCols(norm[0]), 1);

  // mergedCell 若是普通对象（外部调用）也不能崩
  const upObj = fns.buildUploadTables(server, { '0|0|0': 'Z' });
  eq('buildUploadTables 接受普通对象 mergedCell', upObj[0].rows[0].cells[0], 'Z');
}

console.log('\n=== 12. 表集合不一致：以服务器已存数据为准，不凭空加表 ===');
{
  // 接口模板 2 张（id 0,9），服务器已存只有 1 张（id 0）→ 导出/上传都只应出现 id 0
  const server = mkServer([mkTable(0, ['a'], [['s']])]);
  const file = mkFile([mkTable(0, ['a'], [['f']]), mkTable(9, ['z'], [['zz']])],
    [mkTable(0, ['a'], [['s']]), mkTable(9, ['z'], [['']])]);
  const m = fns.computeMerge(file, server);
  eq('表 9 被跳过', m.warnings.some((w) => w.includes('id=9')), true);
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('上传体只有服务器认得的那 1 张表', up.map((t) => t.id), [0]);
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败\n`);
process.exit(fail ? 1 : 0);
