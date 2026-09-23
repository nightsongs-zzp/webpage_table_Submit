/**
 * 表格分区边界 + 公式取值 的针对性自测（真 SheetJS）。
 * 用法： node tests/test-boundary-formula.js
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const XLSX = require('xlsx');

const SRC = path.join(__dirname, '..', 'webpage-table-sync.user.js');
const src = fs.readFileSync(SRC, 'utf8');
const cut = src.indexOf("if (document.readyState === 'loading')");
const body = src.slice(0, cut) + '\nreturn { parseFileToPayload, computeMerge, normCell, cellKey, BORDER_MARK };\n})();';

function makeEl() {
  return {
    style: {}, dataset: {}, classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    children: [], innerHTML: '', textContent: '', value: '',
    appendChild(c) { return c; }, removeChild() {}, addEventListener() {}, removeEventListener() {},
    querySelector() { return null; }, querySelectorAll() { return []; }, setAttribute() {}, getAttribute() { return null; }, click() {},
  };
}
const sandbox = {
  console, XLSX,
  document: { readyState: 'complete', head: makeEl(), body: makeEl(), createElement: () => makeEl(), getElementById: () => null, querySelector: () => null, querySelectorAll: () => [], addEventListener() {} },
  window: { location: { href: 'http://x/public/experiment.html?id=22', search: '?id=22', pathname: '/public/experiment.html', host: 'x' }, addEventListener() {}, innerWidth: 1, innerHeight: 1 },
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

let pass = 0, fail = 0;
function eq(name, actual, expect) {
  const a = JSON.stringify(actual), e = JSON.stringify(expect);
  if (a === e) { pass++; console.log('  ✓ ' + name); } else { fail++; console.log(`  ✗ ${name}   → actual=${a} expect=${e}`); }
}
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ✓ ' + name); } else { fail++; console.log('  ✗ ' + name + (extra ? '  → ' + extra : '')); }
}

const B1 = '\u25AC\u25AC 数据区 (表 id=1) 开始 \u25AC\u25AC';
const B1E = '\u25AC\u25AC 数据区 (表 id=1) 结束 \u25AC\u25AC';
const BM = fns.BORDER_MARK;
console.log('边界标记字符 = ' + JSON.stringify(BM) + '\n');

// ============ A. 公式：读缓存值 ============
console.log('=== A. XLSX 公式格：必须取"最终值" ===');
{
  const wb = XLSX.utils.book_new();
  const ws = {};
  // 手排单元格：A1 表名，A2 表头行，A3..A5 数据
  const put = (addr, cell) => { ws[addr] = cell; };
  put('A1', { t: 's', v: '表 id=1  表1' });
  put('A2', { t: 's', v: '序号' });
  put('B2', { t: 's', v: '电压 $V$' });
  put('C2', { t: 's', v: '两倍电压' });
  put('D2', { t: 's', v: BM });
  put('A3', { t: 's', v: '1' });
  put('B3', { t: 'n', v: 1.234 });
  // C3 是公式，缓存值 2.468
  put('C3', { t: 'n', v: 2.468, f: 'B3*2' });
  put('D3', { t: 's', v: BM });
  put('A4', { t: 's', v: '2' });
  put('B4', { t: 'n', v: 3 });
  // C4 是公式但没有缓存值（文件被程序生成、Excel 没算过）
  put('C4', { t: 'n', v: 0, f: 'B4*2' });
  delete ws.C4.v;
  ws.C4 = { t: 'n', f: 'B4*2' };
  put('D4', { t: 's', v: BM });
  ws['!ref'] = 'A1:D4';
  wb.SheetNames.push('T1_表1');
  wb.Sheets['T1_表1'] = ws;
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

  const p = fns.parseFileToPayload(null, Buffer.from(buf), 'x.xlsx');
  eq('识别 1 张表', p.tables.length, 1);
  eq('表头正确', p.tables[0].headers, ['序号', '电压 $V$', '两倍电压']);
  eq('公式格取到缓存值 2.468（不是 "B3*2"）', p.tables[0].rows[0].cells[2], '2.468');
  ok('没有把公式文本当成数据', !JSON.stringify(p.tables).includes('B3*2'), JSON.stringify(p.tables));
  eq('普通数字格照常', p.tables[0].rows[0].cells[1], '1.234');
  eq('无缓存值的公式格 → 空字符串', p.tables[0].rows[1].cells[2], '');
  ok('无缓存值时给出告警', p.warnings.some((w) => w.includes('公式')), JSON.stringify(p.warnings));
  eq('右侧边界标记列已剥掉（3 列）', p.tables[0].headers.length, 3);
}

// ============ B. 分区边界 ============
console.log('\n=== B. 分区边界：横幅之外的内容不得进入表格 ===');
{
  const wb = XLSX.utils.book_new();
  const aoa = [
    [B1],                                   // 上边界横幅
    ['表 id=1  表1：测试表'],                // 表名
    ['序号', '电压 $V$', BM],               // 表头
    ['1', '1.1', BM],
    ['2', '2.2', BM],
    [B1E],                                  // 下边界横幅
    ['↑ 上面是表格区，不要在下方继续写数据'],
    ['这是表外的备注，绝不能进表格'],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  wb.SheetNames.push('T1_表1');
  wb.Sheets['T1_表1'] = ws;
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

  const p = fns.parseFileToPayload(null, Buffer.from(buf), 'x.xlsx');
  const t = p.tables[0];
  eq('表头正确（自动跳过横幅与表名行）', t.headers, ['序号', '电压 $V$']);
  eq('只收表格区内的 2 行数据', t.rows.length, 2);
  eq('数据内容正确', t.rows.map((r) => r.cells), [['1', '1.1'], ['2', '2.2']]);
  ok('横幅之外的内容没有混进表格', !JSON.stringify(t.rows).includes('表外'), JSON.stringify(t.rows));
  ok('横幅文字没进表头', !t.headers.some((h) => h.includes('数据区')), JSON.stringify(t.headers));
}

// ============ C. 老文件（无横幅、无边界列）仍然能读 ============
console.log('\n=== C. 向后兼容：旧格式（无横幅）仍可读 ===');
{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['表 id=7  旧格式表'],
    ['a', 'b'],
    ['1', '2'],
    ['3', '4'],
  ]);
  wb.SheetNames.push('T7_旧格式表');
  wb.Sheets['T7_旧格式表'] = ws;
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const p = fns.parseFileToPayload(null, Buffer.from(buf), 'old.xlsx');
  eq('表 id 从工作表名还原', p.tables[0].id, 7);
  eq('旧格式表头正确', p.tables[0].headers, ['a', 'b']);
  eq('旧格式 2 行数据', p.tables[0].rows.map((r) => r.cells), [['1', '2'], ['3', '4']]);
}

// ============ D. 只有下边界（上半部分被用户删了）也能收住 ============
console.log('\n=== D. 只保留下边界也能正确截断 ===');
{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['表 id=3  表3'],
    ['x', 'y'],
    ['1', '2'],
    [B1E.replace('id=1', 'id=3')],
    ['表外内容'],
  ]);
  wb.SheetNames.push('T3_表3');
  wb.Sheets['T3_表3'] = ws;
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const p = fns.parseFileToPayload(null, Buffer.from(buf), 'x.xlsx');
  eq('在结束横幅处截断', p.tables[0].rows.map((r) => r.cells), [['1', '2']]);
}

// ============ E. 合并后不影响上传形态 ============
console.log('\n=== E. 解析结果可直接进合并 ===');
{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    [B1],
    ['表 id=1  表1'],
    ['序号', '电压 $V$', BM],
    ['1', '9.9', BM],
    [B1E],
  ]);
  wb.SheetNames.push('T1_表1');
  wb.Sheets['T1_表1'] = ws;
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const p = fns.parseFileToPayload(null, Buffer.from(buf), 'x.xlsx');
  const server = { tables: [{ id: 1, name: '表1', headers: ['序号', '电压 $V$'], rows: [['1', '']] }], update_time: 't' };
  const m = fns.computeMerge(p, server);
  ok('能算出改动', m.changes.length >= 1, JSON.stringify(m.stats));
  eq('改动是新填的 9.9', m.changes.some((c) => c.to === '9.9'), true);
  ok('BORDER 标记没有进入改动', !m.changes.some((c) => String(c.to).includes(BM)), JSON.stringify(m.changes));
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败\n`);
process.exit(fail ? 1 : 0);
