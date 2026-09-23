/**
 * XLSX 往返自测：按脚本里 exportXLSX 的同一套结构造一个工作簿 →
 * 写盘 → 用 parseFileToPayload 读回来 → 校验表格/基线/三方合并是否正确。
 * 用法： node tests/test-xlsx-roundtrip.js
 */
'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const XLSX = require('xlsx');

// ---------- 载入 userscript 的纯函数 ----------
const SRC = path.join(__dirname, '..', 'webpage-table-sync.user.js');
const src = fs.readFileSync(SRC, 'utf8');
const marker = 'if (document.readyState === \'loading\')';
const cut = src.indexOf(marker);
const body = src.slice(0, cut) + '\nreturn { parseFileToPayload, computeMerge, buildUploadTables, cellKey, normCell, CLEAR_TOKEN, sanitizeSheetNames, PRESET_SHEET, BASE_SHEET, BORDER_MARK };\n})();';

function makeEl() {
  return {
    style: {}, dataset: {}, classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    children: [], innerHTML: '', textContent: '', value: '',
    appendChild(c) { this.children.push(c); return c; }, removeChild() {}, addEventListener() {}, removeEventListener() {},
    querySelector() { return null; }, querySelectorAll() { return []; }, setAttribute() {}, getAttribute() { return null; }, click() {},
  };
}
const sandbox = {
  console, XLSX,
  document: { readyState: 'complete', head: makeEl(), body: makeEl(), createElement: () => makeEl(), getElementById: () => null, querySelector: () => null, querySelectorAll: () => [], addEventListener() {} },
  window: { location: { href: 'http://test-host/public/experiment.html?id=33', search: '?id=33', pathname: '/public/experiment.html', host: 'test-host' }, addEventListener() {}, innerWidth: 1200, innerHeight: 800 },
  location: { href: 'http://test-host/public/experiment.html?id=33', search: '?id=33', pathname: '/public/experiment.html', host: 'test-host' },
  localStorage: { _d: {}, getItem(k) { return this._d[k] === undefined ? null : this._d[k]; }, setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; } },
  URL: { createObjectURL: () => 'blob:x', revokeObjectURL() {} }, Blob: function () {},
  alert: () => {}, confirm: () => true, prompt: () => '1',
  setTimeout, clearTimeout, Date, Math, JSON, Number, String, Array, Object, RegExp, Map, Set, isNaN, parseInt, parseFloat,
  fetch: async () => { throw new Error('测试环境不应发起 fetch'); },
};
sandbox.globalThis = sandbox;
sandbox.window.localStorage = sandbox.localStorage;
const fns = vm.runInNewContext(body, sandbox, { filename: 'userscript.js' });

let pass = 0, fail = 0;
function eq(name, actual, expect) {
  const a = JSON.stringify(actual), e = JSON.stringify(expect);
  if (a === e) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log(`  ✗ ${name}   → actual=${a} expect=${e}`); }
}
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ✓ ' + name); }
  else { fail++; console.log('  ✗ ' + name + (extra ? '  → ' + extra : '')); }
}

// ---------- 造 payload（等价于真实导出：5 张表，列数 6/11/7/8/5） ----------
const HEADERS = {
  0: ['序号', '灯丝电流 $I_f$ (A)', '灯丝温度 $T$ ($10^{3}$ K)', '绝对温度 $T$ (K)', '$1/T$ ($10^{-4}$ K$^{-1}$)', '本次实验是否使用'],
  1: ['灯丝电流 $I_f$ (A)', '灯丝温度 $T$ (K)', '$U_a=16$ V', '$U_a=25$ V'],
};
const ROWS = { 0: 7, 1: 6 };
const tables = Object.keys(HEADERS).map((id) => ({
  id: parseInt(id, 10),
  name: `表${id}`,
  headers: HEADERS[id],
  rows: Array.from({ length: ROWS[id] }, (_, r) => ({
    cells: HEADERS[id].map((_, c) => (r === 0 && c === 0 ? String(r + 1) : (r < 2 && c < 2 ? `$1.5\\times10^{-${r + 3}}$` : ''))),
  })),
}));

// 故意放几个 Excel 会「自作主张」的值
tables[0].rows[2].cells[0] = 1800;               // 整数
tables[0].rows[2].cells[1] = 0.00015;            // 会被存成数字
tables[0].rows[2].cells[2] = '007';              // 看起来像数字的文本
tables[0].rows[3].cells[0] = 'x|y,z 空格 结尾 '; // 含 | , 与空格
const baseline = tables.map((t) => ({ id: t.id, name: t.name, headers: t.headers.slice(), rows: t.rows.map((r) => ({ cells: r.cells.slice() })) }));

// ---------- 用与 exportXLSX 相同的结构写工作簿 ----------
const sheetFromAoa = (aoa, widths) => {
  const rows = Array.isArray(aoa) ? aoa : [];
  let maxCols = 0;
  rows.forEach((r) => { if (Array.isArray(r) && r.length > maxCols) maxCols = r.length; });
  const padded = rows.map((r) => { const row = Array.isArray(r) ? r.slice() : []; while (row.length < maxCols) row.push(''); return row; });
  const ws = XLSX.utils.aoa_to_sheet(padded, { cellDates: false });
  if (rows.length && maxCols) ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length - 1, c: maxCols - 1 } });
  ws['!cols'] = widths.map((w) => ({ wch: Math.min(40, Math.max(8, w)) }));
  return ws;
};
const wb = XLSX.utils.book_new();
wb.SheetNames.push('说明_README');
wb.Sheets['说明_README'] = sheetFromAoa([['实验数据表格同步器 — 说明'], ['文件标记', 'webpage_table_sync/v1']], [14, 100]);
const baseAoa = [['表ID', '行', '列', '服务器原值（导出时快照，请勿修改）']];
baseline.forEach((t) => t.rows.forEach((r, ri) => r.cells.forEach((c, ci) => baseAoa.push([t.id, ri, ci, c]))));
wb.SheetNames.push('_基线');
wb.Sheets['_基线'] = sheetFromAoa(baseAoa, [8, 6, 6, 60]);
// 预设页（与脚本 exportXLSX 同名同结构）
const presetAoa = [['表ID', '行', '列', '模板预置值（导出时快照，请勿修改）']];
tables.forEach((t) => t.rows.forEach((r, ri) => r.cells.forEach((c, ci) => presetAoa.push([t.id, ri, ci, String(c)]))));
wb.SheetNames.push(fns.PRESET_SHEET);
wb.Sheets[fns.PRESET_SHEET] = sheetFromAoa(presetAoa, [8, 6, 6, 60]);
const names = fns.sanitizeSheetNames(tables.map((t) => ({ id: t.id, name: `T${t.id}_${t.name || '表' + t.id}` })));
const BORDER = fns.BORDER_MARK;
const bStart = (id) => `\u25AC\u25AC\u25AC 数据区 (表 id=${id}) 开始 \u25AC\u25AC\u25AC`;
const bEnd = (id) => `\u25AC\u25AC\u25AC 数据区 (表 id=${id}) 结束\u25AC\u25AC\u25AC`;
/** 与脚本 exportXLSX 完全一致的表页结构：上横幅 / 表名 / 列名+边界列 / 数据+边界列 / 下横幅 / 提示行 */
const dataRows = [];
tables.forEach((t, idx) => {
  const cols = t.headers.length;
  const aoa = [[bStart(t.id)], [`表 id=${t.id}  ${t.name || ''}`.trim()], t.headers.concat([BORDER])];
  t.rows.forEach((r) => {
    const cells = [];
    for (let c = 0; c < cols; c++) cells.push(c < r.cells.length ? r.cells[c] : '');
    cells.push(BORDER);
    aoa.push(cells);
  });
  aoa.push([bEnd(t.id)]);
  aoa.push([`↑ 上面是「表 id=${t.id}」的数据区。`]);
  dataRows.push(aoa);
  wb.SheetNames.push(names[idx].sheet);
  wb.Sheets[names[idx].sheet] = sheetFromAoa(aoa, t.headers.map((h) => Math.max(String(h || '').length + 3, 11)).concat([3]));
});
// Excel 行号：上横幅=1，表名=2，列名=3，数据从第 4 行起
const FIRST_DATA_EXCEL_ROW = 4;
const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
const tmp = path.join(os.tmpdir(), 'wts-roundtrip.xlsx');
fs.writeFileSync(tmp, Buffer.from(buf));
console.log(`\n临时文件：${tmp}（${fs.statSync(tmp).size} B）\n`);

// ---------- 模拟用户：手工改一个格 + 清空一个格 ----------
const modified = XLSX.read(buf, { type: 'array', cellDates: false });
{
  const sn = names[0].sheet;
  const ws = modified.Sheets[sn];
  // 数据第 3 行（0-based r=2）→ Excel 第 4+2 = 6 行
  const excelRow = FIRST_DATA_EXCEL_ROW + 2;
  ws['B' + excelRow] = { t: 's', v: '$9.99\\times10^{-9}$' };  // 改成新值
  delete ws['A' + excelRow];                                    // 第 1 列删空 → 视为清空
  modified.Sheets[sn] = ws;
}
const buf2 = XLSX.write(modified, { bookType: 'xlsx', type: 'array', compression: true });

// ---------- 读回校验 ----------
console.log('=== A. 原样读回（未改动） ===');
{
  const p = fns.parseFileToPayload(null, Buffer.from(buf), 'exp33.xlsx');
  eq('识别 2 张表', p.tables.length, 2);
  eq('表 id 从工作表名还原', p.tables.map((t) => t.id), [0, 1]);
  eq('表头无损（含 LaTeX 与 $ 符号）', p.tables[0].headers[1], '灯丝电流 $I_f$ (A)');
  // 关键：即使数据行全空，也必须原样回来（靠行尾边界列撑住，否则 Excel 会整行省略 → 导入时误判删除）
  if (JSON.stringify(p.tables.map((t) => t.rows.length)) !== '[7,6]') {
    const rx = XLSX.read(Buffer.from(buf), { type: 'array' });
    const wsx = rx.Sheets[names[0].sheet];
    console.log('  [dbg] !ref=' + wsx['!ref'] + ' 造表行数=' + dataRows[0].length);
    for (let i = 3; i < dataRows[0].length; i++) {
      const aoa = XLSX.utils.sheet_to_json(wsx, { header: 1, defval: '<空>', blankrows: true });
      console.log(`  [dbg] 行${i + 1} 写的是 ${JSON.stringify(dataRows[0][i])} → 读回 ${JSON.stringify(aoa[i])}`);
    }
  }
  eq('行数正确（含末尾空行，一列不少）', p.tables.map((t) => t.rows.length), [7, 6]);
  eq('空行内容为空而不是缺失', p.tables[0].rows[6].cells, ['', '', '', '', '', '']);
  eq('列数正确', p.tables.map((t) => t.headers.length), [6, 4]);
  eq('LaTeX 单元格无损', p.tables[0].rows[1].cells[0], '$1.5\\times10^{-4}$');
  eq('含 | 与 , 的单元格无损', p.tables[0].rows[3].cells[0], 'x|y,z 空格 结尾 ');
  eq('带出基线', !!p.baseline, true);
  eq('带出预设层', !!p.presets, true);
  eq('预设层内容正确（第1行第1列=1）', p.presets.get(fns.cellKey(0, 0, 0)), '1');
  ok('行尾边界列已被剥掉（不应出现在数据里）', !JSON.stringify(p.tables).includes(BORDER), '还有 ' + BORDER + ' 残留');
  const server = { tables: JSON.parse(JSON.stringify(tables)), update_time: 't' };
  const m = fns.computeMerge(p, server);
  console.log(`     （数字/文本还原：[1800]=${JSON.stringify(p.tables[0].rows[2].cells[0])}, [0.00015]=${JSON.stringify(p.tables[0].rows[2].cells[1])}, [007]=${JSON.stringify(p.tables[0].rows[2].cells[2])}）`);
  eq('未改动 → 0 处改动（数字类型转换不产生误报）', m.changes.length, 0);
}

console.log('\n=== B. 改一格 + 清空一格 ===');
{
  const p = fns.parseFileToPayload(null, Buffer.from(buf2), 'exp33.xlsx');
  eq('改后 LaTeX 值读回正确', p.tables[0].rows[2].cells[1], '$9.99\\times10^{-9}$');
  const server = { tables: JSON.parse(JSON.stringify(tables)), update_time: 't' };
  const m = fns.computeMerge(p, server);
  eq('识别 2 处改动（1 改 + 1 清空）', m.changes.length, 2);
  eq('mod=1 del=1', [m.stats.mod, m.stats.del], [1, 1]);
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('上传体：改的格是新值', up[0].rows[2].cells[1], '$9.99\\times10^{-9}$');
  eq('上传体：删空的格变成空字符串', up[0].rows[2].cells[0], '');
  eq('上传体：同行其它格不受影响', up[0].rows[2].cells[2], '007');
  eq('上传体：另一张表完全未动', up[1].rows, JSON.parse(JSON.stringify(tables[1].rows)));
  eq('上传体：表数量与服务器一致', up.length, 2);
}

console.log('\n=== C. 手工 Excel（无 _基线 工作表） ===');
{
  const wb2 = XLSX.utils.book_new();
  wb2.SheetNames.push('T0_手工表');
  wb2.Sheets['T0_手工表'] = sheetFromAoa([
    ['表 id=0 手工'],
    ['a', 'b'],
    ['server-cell', 'new-from-user'],
  ], [20, 20]);
  const b2 = XLSX.write(wb2, { bookType: 'xlsx', type: 'array' });
  const p = fns.parseFileToPayload(null, Buffer.from(b2), 'manual.xlsx');
  eq('无基线 → baseline 为 null', p.baseline, null);
  eq('产生「只补空缺」告警', p.warnings.some((w) => w.includes('只补空缺')), true);
  const server = { tables: [{ id: 0, name: '表0', headers: ['a', 'b'], rows: [{ cells: ['server-cell', ''] }] }], update_time: 't' };
  const m = fns.computeMerge(p, server);
  eq('只补空缺：1 处新增', [m.stats.add, m.stats.mod, m.stats.del], [1, 0, 0]);
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('绝不覆盖服务器已有值', up[0].rows[0].cells[0], 'server-cell');
  eq('空缺补上', up[0].rows[0].cells[1], 'new-from-user');
}

console.log('\n=== D. 含 __CLEAR__ 的 JSON 清空 ===');{
  const payload = {
    tables: [{ id: 0, name: '表0', headers: ['a'], rows: [{ cells: [fns.CLEAR_TOKEN] }] }],
    _baseline: [{ id: 0, name: '表0', headers: ['a'], rows: [{ cells: ['keep'] }] }],
  };
  const p = fns.parseFileToPayload(JSON.stringify(payload), null, 'x.json');
  const server = { tables: [{ id: 0, name: '表0', headers: ['a'], rows: [{ cells: ['keep'] }] }], update_time: 't' };
  const m = fns.computeMerge(p, server);
  eq('__CLEAR__ → 1 处清空', [m.stats.del, m.stats.mod], [1, 0]);
  const up = fns.buildUploadTables(server, m.mergedCell);
  eq('清空后为空字符串', up[0].rows[0].cells[0], '');
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败\n`);
try { fs.unlinkSync(tmp); } catch (e) {}
process.exit(fail ? 1 : 0);
