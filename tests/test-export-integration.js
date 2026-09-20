/**
 * 导出→导入 全链路集成测试（伪造网络层，不连真服务器）。
 *
 * 覆盖 v1.1.0 修掉的两个真实 bug 形态：
 *   A. 服务器 rows 是「裸数组」时导出崩溃（exp22 现场）
 *   B. 模板预置单元格（如序号列 1..12）在导出里丢失、以及导入时被误判成"用户填的"
 *
 * 用法： node tests/test-export-integration.js
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = path.join(__dirname, '..', 'webpage-table-sync.user.js');
const src = fs.readFileSync(SRC, 'utf8');
const marker = 'if (document.readyState === \'loading\')';
const cut = src.indexOf(marker);
const body = src.slice(0, cut) + '\n' +
  'return { S, refreshAll, buildPayload, buildExportView, computeMerge, buildUploadTables, parseFileToPayload, cellKey };\n})();';

// ---------------- 网络伪造 ----------------
let apiHandler = () => { throw new Error('未设置 handler'); };
function setApiHandler(fn) { apiHandler = fn; }

function makeEl() {
  const el = {
    style: {}, dataset: {}, classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    children: [], innerHTML: '', textContent: '', value: '', checked: false,
    appendChild(c) { this.children.push(c); return c; }, removeChild() {}, addEventListener() {},
    removeEventListener() {}, querySelector() { return null; }, querySelectorAll() { return []; },
    setAttribute() {}, getAttribute() { return null; }, click() {},
  };
  return el;
}
const dom = {
  readyState: 'complete', head: makeEl(), body: makeEl(),
  createElement: () => makeEl(),
  getElementById: () => null,
  querySelector: () => null,
  querySelectorAll: () => [],
  addEventListener() {},
};
const sandbox = {
  console, document: dom,
  window: { location: { href: 'http://h/public/experiment.html?id=22', search: '?id=22', pathname: '/public/experiment.html', host: 'h' }, addEventListener() {}, innerWidth: 1200, innerHeight: 800 },
  location: { href: 'http://h/public/experiment.html?id=22', search: '?id=22', pathname: '/public/experiment.html', host: 'h' },
  localStorage: { _d: {}, getItem(k) { return this._d[k] === undefined ? null : this._d[k]; }, setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; } },
  URL: { createObjectURL: () => 'blob:x', revokeObjectURL() {} }, Blob: function () {},
  alert: () => {}, confirm: () => true, prompt: () => '1',
  setTimeout, clearTimeout, Date, Math, JSON, Number, String, Array, Object, RegExp, Map, Set, isNaN, parseInt, parseFloat,
  fetch: async (url, opt) => {
    const r = apiHandler(String(url), opt || {});
    return {
      ok: r.status >= 200 && r.status < 300,
      status: r.status,
      text: async () => JSON.stringify(r.json),
    };
  },
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
  if (cond) { pass++; console.log('  ✓ ' + name); } else { fail++; console.log('  ✗ ' + name + (extra ? '  → ' + extra : '')); }
}

// ---------------- 伪数据：exp22 的真实形态 ----------------
// 接口模板：2 张表，序号列有预置值（页面上学生看得到）
const DEF = [
  {
    id: 1, name: '表1：升温过程原始数据记录',
    tableDef: {
      headers: ['序号', '铂电阻温度计电压 $V_{Pt}$ (mV)', '样品电压 $V_{s}$ (V)', '备注'],
      rows: [
        { cells: ['1', '', '', ''] },
        { cells: ['2', '', '', ''] },
        { cells: ['3', '', '', ''] },
        { cells: ['4', '', '', ''] },
      ],
    },
  },
  {
    id: 2, name: '表2：降温过程原始数据记录',
    tableDef: {
      headers: ['序号', '样品电压 $V_{s}$ (V)'],
      rows: [{ cells: ['1', ''] }, { cells: ['2', ''] }],
    },
  },
];

// 服务器已存数据：**rows 故意用裸数组形态**（exp22 崩溃现场）
function serverPayload(rowsStyle, edits) {
  const mk = (id, name, headers, rows) => ({
    id, name, headers,
    rows: rowsStyle === 'array' ? rows.map((r) => r.slice()) : rows.map((r) => ({ cells: r.slice() })),
  });
  const t1 = [['1', '', '', ''], ['2', '', '', ''], ['3', '', '', ''], ['4', '', '', '']];
  const t2 = [['1', ''], ['2', '']];
  if (edits) edits(t1, t2);
  return { code: 1, data: { table_data: { tables: [mk(1, '表1：升温过程原始数据记录', DEF[0].tableDef.headers, t1), mk(2, '表2：降温过程原始数据记录', DEF[1].tableDef.headers, t2)] }, update_time: '2026-09-20 18:00:00' } };
}

const routes = {
  exp: () => ({ status: 200, json: { code: 1, data: { id: 22, name: '高温超导', dataTableDefinition: JSON.stringify(DEF) } } }),
  data: () => ({ status: 200, json: serverPayload('array', null) }),
  post: (opt) => ({ status: 200, json: { code: 1, msg: 'ok', _echo: JSON.parse(opt.body || '{}') } }),
};
let lastPost = null;
apiHandler = (u, o) => {
  if (/\/api\/experiment\//.test(u)) return routes.exp();
  if (/\/api\/student\/data-record$/.test(u) && o.method === 'POST') {
    const r = routes.post(o); lastPost = r.json._echo; return r;
  }
  if (/\/api\/student\/data-record\//.test(u)) return routes.data();
  if (/\/api\/student\/info/.test(u)) return { status: 200, json: { code: 1, data: { username: '测试' } } };
  throw new Error('未预期的请求：' + u);
};

(async () => {
  console.log('\n=== A. 服务器 rows 为裸数组：导出不再崩 ===');
  let payload = null;
  try {
    await fns.refreshAll();
    payload = fns.buildPayload();
    ok('导出成功（旧版在此崩溃）', true);
  } catch (e) {
    ok('导出成功（旧版在此崩溃）', false, e.constructor.name + ': ' + e.message);
    console.log(`\n结果：${pass} 通过 / ${fail} 失败\n`);
    process.exit(1);
  }
  eq('导出 2 张表', payload.tables.length, 2);
  eq('表1 尺寸 4×4', [payload.tables[0].rows.length, payload.tables[0].headers.length], [4, 4]);

  console.log('\n=== B. 模板预置单元格必须体现在导出里 ===');
  eq('表1 第1行 = 预置序号 1', payload.tables[0].rows[0].cells, ['1', '', '', '']);
  eq('表1 序号列 = 1,2,3,4（模板预置）', payload.tables[0].rows.map((r) => r.cells[0]), ['1', '2', '3', '4']);
  eq('_presets 也带上了', payload._presets.map((t) => t.rows.map((r) => r.cells[0])), [['1', '2', '3', '4'], ['1', '2']]);
  eq('summary.filled 把预置格也算进去', payload.summary.map((s) => s.filled), [4, 2]);

  console.log('\n=== C. 预置格没被动 → 导入必须 0 改动（关键回归） ===');
  {
    const filePayload = fns.parseFileToPayload(JSON.stringify(payload), null, 'exp22.json');
    eq('解析出预设层', !!filePayload.presets, true);
    const merge = fns.computeMerge(filePayload, fns.S.server);
    eq('未改动 → changes 为 0', merge.changes.length, 0);
    eq('stats 全 0', [merge.stats.mod, merge.stats.add, merge.stats.del], [0, 0, 0]);
  }

  console.log('\n=== D. 用户填数 → 只上传他填的那些格 ===');
  {
    const edited = JSON.parse(JSON.stringify(payload));
    edited.tables[0].rows[0].cells[1] = '1.234';   // 用户填
    edited.tables[0].rows[0].cells[3] = '备注A';    // 用户填
    edited.tables[0].rows[1].cells[1] = '$2.5\\times10^{-3}$';
    const fp = fns.parseFileToPayload(JSON.stringify(edited), null, 'exp22.json');
    const merge = fns.computeMerge(fp, fns.S.server);
    eq('恰好 3 处改动', merge.changes.length, 3);
    eq('序号列预置值没被当成用户数据上传', merge.changes.some((c) => c.c === 0), false);
    const up = fns.buildUploadTables(fns.S.server, merge.mergedCell, fns.S.defs);
    eq('上传体保留服务器全部 2 张表', up.length, 2);
    eq('上传体：用户填的值到位', up[0].rows[0].cells.slice(0, 2), ['1', '1.234']);
    eq('上传体：备注到位', up[0].rows[0].cells[3], '备注A');
    eq('上传体：LaTeX 原样', up[0].rows[1].cells[1], '$2.5\\times10^{-3}$');
    eq('上传体：行形态统一为 {cells:[...]}', Array.isArray(up[0].rows[0].cells), true);
    eq('上传体：表2 原封不动', up[1].rows.map((r) => r.cells), [['1', ''], ['2', '']]);
  }

  console.log('\n=== E. 【预置格永不入库】导出时模板预置格有值、服务器上没有 → 不上传 ===');
  {
    // 服务器上是"干净"状态：序号列没有存过任何值 —— 这正是 exp22 真机的样子
    const cleanServer = () => ({
      tables: [{ id: 1, name: '表1', headers: ['序号', '电压', '备注', 'x'], rows: [['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', '']] }],
      update_time: 't', table_data: null,
    });
    const p = JSON.parse(JSON.stringify(payload));
    eq('导出视图里序号列 = 模板预置 1,2,3,4', p.tables[0].rows.map((r) => r.cells[0]), ['1', '2', '3', '4']);

    // ① 一个字都不改 → 0 上传，且全部记入 presetSkipped
    {
      const fp = fns.parseFileToPayload(JSON.stringify(p), null, 'exp22.json');
      const merge = fns.computeMerge(fp, cleanServer());
      eq('未改动 → 0 处上传', merge.changes.length, 0);
      eq('4 个预置格被记为 presetSkipped', merge.stats.presetSkipped, 4);
      ok('没有预置格进入上传体', !merge.mergedCell.has('1|0|0') && !merge.mergedCell.has('1|1|0'));
    }
    // ② 把预置的"1"删成空 → 依然不上传（模板回落会照旧显示 1）
    {
      const e = JSON.parse(JSON.stringify(p));
      e.tables[0].rows[0].cells[0] = '';
      const fp = fns.parseFileToPayload(JSON.stringify(e), null, 'exp22.json');
      const merge = fns.computeMerge(fp, cleanServer());
      eq('删预置格 → 0 处上传', [merge.changes.length, merge.stats.del], [0, 0]);
      eq('仍记入 presetSkipped', merge.stats.presetSkipped > 0, true);
    }
    // ③ 真的把预置格改成 99 → 正常上传
    {
      const e = JSON.parse(JSON.stringify(p));
      e.tables[0].rows[0].cells[0] = '99';
      const fp = fns.parseFileToPayload(JSON.stringify(e), null, 'exp22.json');
      const merge = fns.computeMerge(fp, cleanServer());
      eq('改成 99 → 1 处新增（服务器原来没值）', [merge.changes.length, merge.stats.add], [1, 1]);
      const up = fns.buildUploadTables(cleanServer(), merge.mergedCell, fns.S.defs);
      eq('上传体该格是 99', up[0].rows[0].cells[0], '99');
      eq('其它预置格仍不入库', up[0].rows.slice(1).map((r) => r.cells[0]), ['', '', '']);
    }
    // ④ 用户填了非预置格（第 2 列）→ 照常上传
    {
      const e = JSON.parse(JSON.stringify(p));
      e.tables[0].rows[0].cells[1] = '9.9';
      const fp = fns.parseFileToPayload(JSON.stringify(e), null, 'exp22.json');
      const merge = fns.computeMerge(fp, cleanServer());
      eq('填非预置格 → 1 处新增', [merge.changes.length, merge.stats.add], [1, 1]);
      const up = fns.buildUploadTables(cleanServer(), merge.mergedCell, fns.S.defs);
      eq('上传体：9.9 到位、预置格仍为空', up[0].rows[0].cells.slice(0, 2), ['', '9.9']);
    }
  }

  console.log('\n=== E2. 服务器上真的存过序号（旧数据遗留）→ 用户清空它才叫删除 ===');
  {
    // 这一格既然服务器存过，就不再是"纯预置格"，用户删空就应该真的删
    const srv = {
      tables: [{ id: 1, name: '表1', headers: ['序号', '电压', '备注', 'x'], rows: [['1', '', '', '']] }],
      update_time: 't', table_data: null,
    };
    const p = JSON.parse(JSON.stringify(payload));
    const fp = fns.parseFileToPayload(JSON.stringify(p), null, 'exp22.json');
    eq('服务器已有值与文件一致 → 0 上传', fns.computeMerge(fp, srv).changes.length, 0);
    const e = JSON.parse(JSON.stringify(p));
    e.tables[0].rows[0].cells[0] = '';
    const fp2 = fns.parseFileToPayload(JSON.stringify(e), null, 'exp22.json');
    const m2 = fns.computeMerge(fp2, srv);
    eq('服务器存过的值被清空 → 1 处删除', [m2.stats.del, m2.changes.length], [1, 1]);
    const up = fns.buildUploadTables(srv, m2.mergedCell, fns.S.defs);
    eq('上传体该格为空', up[0].rows[0].cells[0], '');
  }

  console.log('\n=== E3. 模板改过（导出快照为空、当前模板有值）→ 也不上传 ===');
  {
    const e = JSON.parse(JSON.stringify(payload));
    e._presets[0].rows[0].cells[0] = '';   // 快照里为空
    const srv = {
      tables: [{ id: 1, name: '表1', headers: ['序号', '电压', '备注', 'x'], rows: [['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', '']] }],
      update_time: 't', table_data: null,
    };
    const fp = fns.parseFileToPayload(JSON.stringify(e), null, 'exp22.json');
    const merge = fns.computeMerge(fp, srv);
    eq('模板新加的预置格不会被当成用户数据', merge.changes.length, 0);
  }

  console.log('\n=== F. 服务器已有值的格：用户清空 → 产生删除 ===');
  {
    // 换一份"服务器已有填写值"的数据
    apiHandler = (u, o) => {
      if (/\/api\/experiment\//.test(u)) return routes.exp();
      if (/\/api\/student\/data-record$/.test(u) && o.method === 'POST') { const r = routes.post(o); lastPost = r.json._echo; return r; }
      if (/\/api\/student\/data-record\//.test(u)) {
        return { status: 200, json: serverPayload('array', (t1) => { t1[0][1] = '1.234'; }) };
      }
      return { status: 200, json: { code: 1, data: { username: 'x' } } };
    };
    await fns.refreshAll();
    const p2 = fns.buildPayload();
    eq('服务器值 1.234 已体现在导出', p2.tables[0].rows[0].cells[1], '1.234');
    eq('_baseline 记录的是服务器值', p2._baseline[0].rows[0].cells[1], '1.234');
    const edited = JSON.parse(JSON.stringify(p2));
    edited.tables[0].rows[0].cells[1] = '';   // 用户清空
    const fp = fns.parseFileToPayload(JSON.stringify(edited), null, 'exp22.json');
    const merge = fns.computeMerge(fp, fns.S.server);
    eq('识别 1 处清空', [merge.stats.del, merge.changes.length], [1, 1]);
    const up = fns.buildUploadTables(fns.S.server, merge.mergedCell, fns.S.defs);
    eq('上传体该格变空', up[0].rows[0].cells[1], '');
  }

  console.log('\n=== G. 真 POST 一次：请求体形态符合页面口径 ===');
  {
    const merge = fns.computeMerge(
      fns.parseFileToPayload(JSON.stringify(payload), null, 'x.json'), fns.S.server);
    const up = fns.buildUploadTables(fns.S.server, merge.mergedCell, fns.S.defs);
    const resp = await sandbox.fetch('/api/student/data-record', { method: 'POST', body: JSON.stringify({ expId: 22, tableData: { tables: up } }) });
    const j = JSON.parse(await resp.text());
    eq('POST 返回 code=1', j.code, 1);
    eq('请求体 expId 是数字 22', lastPost.expId, 22);
    ok('请求体 tableData.tables 是数组', Array.isArray(lastPost.tableData.tables), JSON.stringify(lastPost.tableData).slice(0, 80));
    eq('每张表都带 id/name/headers/rows', Object.keys(lastPost.tableData.tables[0]).sort(), ['headers', 'id', 'name', 'rows']);
    eq('每行都是 {cells:[...]}', Object.keys(lastPost.tableData.tables[0].rows[0]), ['cells']);
  }

  console.log(`\n结果：${pass} 通过 / ${fail} 失败\n`);
  process.exit(fail ? 1 : 0);
})();
