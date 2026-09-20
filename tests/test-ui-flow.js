/**
 * 用真 DOM（linkedom）复现「导入上传」全流程，专门抓"弹窗卡住/按钮没反应"这类问题。
 *
 * 用法： node tests/test-ui-flow.js [可选的 .xlsx 路径]
 *       不给路径就用内置造出来的表；给了路径就用真实文件（更好复现现场）。
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { parseHTML } = require('linkedom');
let XLSX = null;
try { XLSX = require('xlsx'); } catch (e) { /* 没装就只测 JSON 路径 */ }

const SRC = path.join(__dirname, '..', 'webpage-table-sync.user.js');
const src = fs.readFileSync(SRC, 'utf8');

// ---------------- 网络伪造 ----------------
// 刻意造得"像真实的"：5 列，序号列有模板预置值，服务器已存数据是裸数组形态
const HEADERS_1 = ['序号', '铂电阻温度计电压 $V_{Pt}$ (mV)', '样品电压 $V_{s}$ (V)', '所处区段', '备注'];
const PRESET_1 = [['1', '', '', '', ''], ['2', '', '', '', '']];
const DEF = [
  { id: 1, name: '表1', tableDef: { headers: HEADERS_1, rows: PRESET_1.map((r) => ({ cells: r.slice() })) } },
];
// 服务器已存数据：rows 用**裸数组**形态（真机实测就是这个形态）
const serverTables = [{ id: 1, name: '表1', headers: HEADERS_1, rows: [['1', '', '', '', ''], ['2', '', '', '', '']] }];
let posts = [];

// ---------------- 真 DOM ----------------
const { document, window } = parseHTML('<!DOCTYPE html><html><head></head><body></body></html>');
window.location = { href: 'http://h/public/experiment.html?id=22', search: '?id=22', pathname: '/public/experiment.html', host: 'h' };

const alerts = [];
const sandbox = {
  console, document, window, XLSX,
  location: window.location,
  localStorage: { _d: {}, getItem(k) { return this._d[k] === undefined ? null : this._d[k]; }, setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; } },
  URL: { createObjectURL: () => 'blob:x', revokeObjectURL() {} },
  Blob: function (parts, opts) { this.parts = parts; this.type = (opts || {}).type; },
  alert: (m) => { alerts.push(String(m)); console.log('  [alert] ' + m); },
  confirm: () => true, prompt: () => '1',
  setTimeout, clearTimeout, Date, Math, JSON, Number, String, Array, Object, RegExp, Map, Set, isNaN, parseInt, parseFloat,
  DOMException: class DOMException extends Error {},
  fetch: async (url, opt) => {
    const u = String(url), method = (opt && opt.method) || 'GET';
    let json;
    if (/\/api\/experiment\//.test(u)) json = { code: 1, data: { id: 22, name: '高温超导', dataTableDefinition: JSON.stringify(DEF) } };
    else if (/\/api\/student\/data-record$/.test(u) && method === 'POST') {
      const body = JSON.parse(opt.body);
      posts.push(body);
      // 让"服务器"真的把数据存下来，这样回读校验才有意义（否则会误报不一致）
      serverTables.length = 0;
      body.tableData.tables.forEach((t) => serverTables.push({
        id: t.id, name: t.name, headers: t.headers.slice(),
        rows: t.rows.map((r) => r.cells.slice()),
      }));
      json = { code: 1, msg: 'ok' };
    }
    else if (/\/api\/student\/data-record\//.test(u)) json = { code: 1, data: { table_data: { tables: serverTables }, update_time: '2026-09-20 10:32:14' } };
    else if (/\/api\/student\/info/.test(u)) json = { code: 1, data: { username: '测试' } };
    else json = { code: 0 };
    return { ok: true, status: 200, text: async () => JSON.stringify(json) };
  },
};
sandbox.globalThis = sandbox;
sandbox.window.localStorage = sandbox.localStorage;
sandbox.window.document = document;

vm.runInNewContext(src, sandbox, { filename: 'userscript.js' });

// ---------------- 断言 ----------------
let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ✓ ' + name); } else { fail++; console.log('  ✗ ' + name + (extra ? '  → ' + extra : '')); }
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  await sleep(120); // 等 main() 的首次预读取完

  console.log('\n=== 1. 面板是否真的插进了 DOM ===');
  const panel = document.getElementById('wts-panel');
  ok('面板存在', !!panel);
  const modal = document.getElementById('wts-modal');
  ok('弹窗容器存在', !!modal);
  ok('预览区存在', !!document.getElementById('wts-preview'));
  const fileInput = document.getElementById('wts-file');
  ok('文件选择框存在', !!fileInput);
  if (!panel || !modal || !fileInput) { console.log(`\n结果：${pass} 通过 / ${fail} 失败\n`); process.exit(1); }

  console.log('\n=== 2. 模拟真实「选择文件」事件 ===');
  const xlsxPath = process.argv[2];
  let fileObj;
  if (xlsxPath) {
    const buf = fs.readFileSync(xlsxPath);
    fileObj = {
      name: path.basename(xlsxPath), size: buf.length,
      arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
      text: async () => buf.toString('utf8'),
    };
  } else {
    // 内置：造一份"用户只把服务器上没值的那格填了值"的 JSON
    // 第 1 列保留预置值 '1'（不该上传），第 2 列 9.9 是用户填的（该上传）
    const payload = {
      tables: [{ id: 1, name: '表1', headers: HEADERS_1, rows: [{ cells: ['1', '9.9', '', '', ''] }, { cells: ['2', '', '', '', ''] }] }],
      _baseline: [{ id: 1, name: '表1', headers: HEADERS_1, rows: [{ cells: ['1', '', '', '', ''] }, { cells: ['2', '', '', '', ''] }] }],
      _presets: [{ id: 1, rows: [{ cells: ['1', '', '', '', ''] }, { cells: ['2', '', '', '', ''] }] }],
    };
    const text = JSON.stringify(payload);
    fileObj = { name: 'test.json', size: text.length, text: async () => text, arrayBuffer: async () => new ArrayBuffer(0) };
  }

  // jsdom/linkedom 里 FileList 不能直接塞，用 defineProperty 模拟
  Object.defineProperty(fileInput, 'files', { value: [fileObj], configurable: true, writable: true });
  fileInput.dispatchEvent(new window.Event('change'));
  await sleep(400);

  console.log('\n=== 3. 弹窗有没有真的显示出来 ===');
  const style = modal.getAttribute('style') || '';
  ok('modal.style.display 被设为 flex', /display:\s*flex/.test(style), 'style=' + JSON.stringify(style));
  const inlineDisplay = modal.style.display;
  ok('modal.style.display === "flex"', inlineDisplay === 'flex', 'display=' + JSON.stringify(inlineDisplay));
  const previewHtml = (document.getElementById('wts-preview').innerHTML || '');
  ok('预览区渲染出了内容', previewHtml.length > 0, '长度=' + previewHtml.length);
  const previewText = previewHtml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  console.log('  预览摘要：' + previewText.slice(0, 160));
  if (!xlsxPath) {
    ok('只报了 1 处改动（预置的序号列没被当成用户数据）', /将修改 0 格 · 新增 1 格 · 清空 0 格/.test(previewText), previewText.slice(0, 90));
  } else {
    // 真文件：文件名前缀是 T1_，只从工作表名推 id，所以这里只校验"弹得出预览"
    ok('真文件也弹出了差异预览', /涉及 \d+ 张表/.test(previewText), previewText.slice(0, 90));
  }
  const okBtn = document.getElementById('wts-modal-ok');
  ok('确认按钮存在', !!okBtn);
  ok('确认按钮文案已设置', /确认|覆盖/.test(okBtn.textContent || ''), 'text=' + JSON.stringify(okBtn.textContent));

  console.log('\n=== 4. 点「确认上传改动」能不能走通 ===');
  const before = posts.length;
  okBtn.dispatchEvent(new window.Event('click'));
  await sleep(700);
  ok('真的发出了 POST', posts.length > before, 'POST 次数=' + (posts.length - before));
  if (posts.length > before) {
    const body = posts[posts.length - 1];
    ok('请求体带 expId', body.expId === 22, 'expId=' + JSON.stringify(body.expId));
    ok('请求体带 tableData.tables', Array.isArray(body.tableData && body.tableData.tables));
    const t1 = body.tableData.tables[0];
    ok('只写了服务器认得的那张表', body.tableData.tables.length === serverTables.length);
    console.log('  上传体表1 第1行：' + JSON.stringify(t1.rows[0]));
    if (!xlsxPath) {
      ok('上传体保留模板预置的序号 1', t1.rows[0].cells[0] === '1', JSON.stringify(t1.rows[0].cells[0]));
      ok('上传体带上了用户填的 9.9', t1.rows[0].cells[1] === '9.9', JSON.stringify(t1.rows[0].cells[1]));
      ok('每一行都补齐成 5 列', t1.rows.every((r) => r.cells.length === 5), JSON.stringify(t1.rows.map((r) => r.cells.length)));
    }
  }
  console.log('\n=== 5. 上传后弹窗是否收起、busy 是否释放 ===');
  ok('弹窗已隐藏', modal.style.display === 'none', 'display=' + JSON.stringify(modal.style.display));
  // busy 释放的验证：再点一次「导入上传」不应再报"已有操作在进行中"
  const logsBefore = (document.getElementById('wts-log').textContent || '');
  fileInput.dispatchEvent(new window.Event('change'));
  await sleep(300);
  const logsAfter = (document.getElementById('wts-log').textContent || '');
  ok('busy 已释放（重复点击不会被挡）', !/已有操作在进行中/.test(logsAfter.slice(logsBefore.length)), '新增日志=' + JSON.stringify(logsAfter.slice(logsBefore.length)));

  console.log(`\n结果：${pass} 通过 / ${fail} 失败\n`);
  process.exit(fail ? 1 : 0);
})().catch((e) => {
  console.error('\n测试自身崩了：', e);
  process.exit(1);
});
