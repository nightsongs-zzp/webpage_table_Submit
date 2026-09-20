// ==UserScript==
// @name         实验数据表格 同步器（导出/导入 · 近代物理实验平台）
// @namespace    https://github.com/nightsongs-zzp/webpage_table_Submit
// @version      1.1.2
// @description  把 experiment.html?id=NN 里的 4.1 数据记录表（多张）导出为 JSON(主存档)+XLSX(方便手算)，本地改完后可安全回传：GET→三方合并→POST→回读校验，上传前自动备份，支持一键回滚。
// @author       nightsongs-zzp
// @match        http://119.91.120.143:8081/public/experiment.html*
// @match        http://119.91.120.143:8081/experiment.html*
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @grant        GM_xmlhttpRequest
// @grant        GM_setClipboard
// @grant        GM_registerMenuCommand
// @connect      119.91.120.143
// @run-at       document-idle
// ==/UserScript==

/* eslint-disable no-console */
(function () {
  'use strict';

  // ============================================================
  // 0. 常量与全局状态
  // ============================================================
  const API_DATA_RECORD = '/api/student/data-record';
  const API_EXPERIMENT = (id) => `/api/experiment/${encodeURIComponent(id)}`;
  const CLEAR_TOKEN = '__CLEAR__';
  const XLSX_MARK = 'webpage_table_sync/v1';
  const BASE_SHEET = '_基线';
  const PRESET_SHEET = '_预设';
  const BAK_KEY = 'wts_backup_history_v1';
  const CFG_KEY = 'wts_config_v1';

  const S = {
    expId: null,
    defs: null,          // [{id, name, headers:[], expectedRows:int, presetRows:[[..]]}] | null
    defsSource: '',
    defRowsById: new Map(), // id → 模板预置行（学生没填时页面显示的兜底值）
    server: null,        // { tables:[{id,name,headers,rows}], table_data:{...}, update_time }
    lastResponse: null,  // 上次 POST 的响应
    busy: false,
    logLines: [],
  };

  // ============================================================
  // 1. 配置持久化
  // ============================================================
  function loadCfg() {
    try { return JSON.parse(localStorage.getItem(CFG_KEY) || '{}'); } catch (e) { return {}; }
  }
  function saveCfg(patch) {
    const c = Object.assign(loadCfg(), patch || {});
    try { localStorage.setItem(CFG_KEY, JSON.stringify(c)); } catch (e) { /* ignore */ }
    return c;
  }
  function filePrefix() {
    const c = loadCfg();
    return (c.prefix && String(c.prefix).trim()) || `exp${S.expId}_`;
  }
  /** 导出文件名：自定义名优先，否则「前缀 + 时间戳」 */
  function buildFileName(ext) {
    const c = loadCfg();
    const custom = (c.fileName && String(c.fileName).trim()) || '';
    if (custom) {
      const base = custom.replace(/[\\/:*?"<>|]/g, '_').replace(/\.(json|xlsx)$/i, '');
      return `${base}.${ext}`;
    }
    return `${filePrefix()}${timestamp()}.${ext}`;
  }

  // ============================================================
  // 2. 小工具
  // ============================================================
  function getUrlParam(name) {
    const reg = new RegExp('(^|&)' + name + '=([^&]*)(&|$)');
    const r = window.location.search.substr(1).match(reg);
    return r ? decodeURIComponent(r[2]) : null;
  }
  function pad2(n) { return String(n).padStart(2, '0'); }
  function timestamp() {
    const d = new Date();
    return `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}_${pad2(d.getHours())}${pad2(d.getMinutes())}${pad2(d.getSeconds())}`;
  }
  function toStr(v) { return v === null || v === undefined ? '' : String(v); }
  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function cellKey(tid, r, c) { return `${tid}|${r}|${c}`; }
  /**
   * 取一行的 cells。**服务器侧 / 文件侧两种行形态都必须能吃**：
   *   - "对象形态"：{ cells: [...] }  ← 页面 saveDataRecord 写回服务器的形态（教学模板就是它）
   *   - "裸数组形态"：[...]           ← 有些接口返回 / 老数据直接是数组
   * 之前只认对象形态，裸数组会在 expCols() 的 r.cells.length 处抛
   * TypeError: Cannot read properties of undefined (reading 'length')。
   */
  function rowCells(row) {
    if (Array.isArray(row)) return row;
    if (row && Array.isArray(row.cells)) return row.cells;
    return [];
  }
  /**
   * 统一 Map / 普通对象 两种"键值容器"：mergedCell 是 Map，但离线自测或外部调用
   * 可能给普通对象，这里一次收口，避免 .has is not a function。
   */
  function mapLike(map) {
    if (map && typeof map.has === 'function' && typeof map.get === 'function') return map;
    const src = map || {};
    return {
      has(k) { return Object.prototype.hasOwnProperty.call(src, k); },
      get(k) { return src[k]; },
      forEach(fn) { Object.keys(src).forEach((k) => fn(src[k], k)); },
    };
  }
  function safeSheetName(raw, fallback) {
    let s = String(raw || '').replace(/[\[\]\*\/\\\?:]/g, '_').replace(/^\s+|\s+$/g, '');
    if (!s) s = fallback;
    return s.slice(0, 31);
  }
  function sanitizeSheetNames(pairs) {
    const used = new Set([BASE_SHEET, PRESET_SHEET, '说明_README']);
    const out = [];
    for (const p of pairs) {
      let base = safeSheetName(p.name, 'T' + p.id).slice(0, 28);
      let name = base, i = 2;
      while (used.has(name)) { name = `${base}(${i++})`.slice(0, 31); }
      used.add(name);
      out.push({ id: p.id, sheet: name });
    }
    return out;
  }
  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { try { document.body.removeChild(a); URL.revokeObjectURL(url); } catch (e) {} }, 4000);
  }
  function diffMs(t) { return Date.now() - t; }
  /** 当前脚本版本（用来判断浏览器里装的是不是最新那版） */
  function scriptVersion() {
    try {
      if (typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version) return GM_info.script.version;
    } catch (e) {}
    return '1.1.2';
  }

  // ============================================================
  // 3. 网络层（fetch 为主；被 CSP 拦时退回 GM_xmlhttpRequest）
  // ============================================================
  function fetchViaGM(method, url, body) {
    return new Promise((resolve, reject) => {
      if (typeof GM_xmlhttpRequest !== 'function') {
        reject(new Error('无可用网络通道（fetch 被拦截且无 GM_xmlhttpRequest）'));
        return;
      }
      GM_xmlhttpRequest({
        method,
        url: new URL(url, location.href).href,
        headers: body ? { 'Content-Type': 'application/json' } : undefined,
        data: body ? JSON.stringify(body) : undefined,
        withCredentials: true,
        onload: (res) => {
          try { resolve({ ok: res.status >= 200 && res.status < 300, status: res.status, json: JSON.parse(res.responseText) }); }
          catch (e) { reject(new Error('响应不是 JSON（HTTP ' + res.status + '）')); }
        },
        onerror: () => reject(new Error('GM_xmlhttpRequest 网络错误')),
        ontimeout: () => reject(new Error('GM_xmlhttpRequest 超时')),
      });
    });
  }

  async function api(method, url, body) {
    const t0 = Date.now();
    let out;
    try {
      const res = await fetch(url, {
        method,
        credentials: 'include',
        headers: body ? { 'Content-Type': 'application/json' } : undefined,
        body: body ? JSON.stringify(body) : undefined,
      });
      let json = null;
      const text = await res.text();
      try { json = text ? JSON.parse(text) : null; } catch (e) { throw new Error('响应不是 JSON（HTTP ' + res.status + '）'); }
      out = { ok: res.ok, status: res.status, json };
    } catch (e) {
      // fetch 抛出的既有网络错误，也可能是"页面脚本自己"的类型错误 —— 都退回 GM 通道再试一次。
      // （这样即使页面注入的 fetch 包装出问题，脚本仍有一条路可走。）
      log('warn', `fetch 失败（${e.name}: ${e.message}），改用 GM_xmlhttpRequest 重试…`);
      out = await fetchViaGM(method, url, body);
    }
    log('net', `${method} ${url} → HTTP ${out.status}（${diffMs(t0)}ms）`);
    return out;
  }

  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

  // ============================================================
  // 4. 数据模型：定义 / 服务器现状 / 文件
  // ============================================================
  function normalizeTableDef(t) {
    const td = (t && (t.tableDef || t)) || {};
    const headers = Array.isArray(td.headers) ? td.headers.map(toStr) : [];
    const rows = Array.isArray(td.rows) ? td.rows : [];
    return {
      id: t && t.id !== undefined && t.id !== null ? t.id : 0,
      name: toStr(t && t.name),
      headers,
      // 模板自带的"预置单元格"：学生没填时页面上显示的就是它，导出必须回落到这里
      presetRows: rows.map((r) => rowCells(r).map(toStr)),
      expectedRows: rows.length,
    };
  }

  function extractDefs(expData) {
    // 主键名：源码里是 expData.dataTableDefinition；再兜底几个可能的名字
    const candidates = ['dataTableDefinition', 'data_table_definition', 'dataTableDefine', 'dataTable', 'data_table'];
    for (const k of candidates) {
      const v = expData && expData[k];
      if (v === undefined || v === null || v === '') continue;
      let parsed = v;
      if (typeof v === 'string') {
        try { parsed = JSON.parse(v); } catch (e) { continue; }
      }
      let tables = null;
      if (Array.isArray(parsed)) tables = parsed;
      else if (parsed && parsed.tableDef) tables = [{ id: 0, name: '数据表格', tableDef: parsed.tableDef }];
      if (!tables) continue;
      const defs = tables.map(normalizeTableDef).filter((d) => d.headers.length > 0);
      if (defs.length) return { defs, source: k };
    }
    return { defs: null, source: '' };
  }

  /** 弹窗里 input 的几何信息兜底（拿不到 dataTableDefinition 时用） */
  function scanDomInputs() {
    const inputs = document.querySelectorAll('#fillableDataTableEditor input.data-table-input');
    if (!inputs.length) return null; // 弹窗没打开就扫不到，属正常
    const byId = new Map();
    inputs.forEach((inp) => {
      const id = inp.dataset.tableId !== undefined ? inp.dataset.tableId : '0';
      const r = parseInt(inp.dataset.row, 10), c = parseInt(inp.dataset.col, 10);
      if (isNaN(r) || isNaN(c)) return;
      if (!byId.has(id)) byId.set(id, { id, maxRow: -1, maxCol: -1 });
      const g = byId.get(id);
      g.maxRow = Math.max(g.maxRow, r);
      g.maxCol = Math.max(g.maxCol, c);
    });
    const defs = [...byId.values()].map((g) => ({
      id: /^\d+$/.test(g.id) ? parseInt(g.id, 10) : g.id,
      name: '',
      headers: new Array(g.maxCol + 1).fill(''),
      expectedRows: g.maxRow + 1,
    }));
    return defs.length ? defs : null;
  }

  /** 从服务器 table_data 里取表结构（行形态统一收口成 {cells:[...]}） */
  function normalizeServerTables(td) {
    if (!td) return [];
    let obj = td;
    if (typeof obj === 'string') {
      try { obj = JSON.parse(obj); } catch (e) { return []; }
    }
    const tables = (obj && Array.isArray(obj.tables)) ? obj.tables : [];
    return tables.map((t) => {
      const rows = Array.isArray(t.rows) ? t.rows : [];
      return {
        id: t && t.id !== undefined && t.id !== null ? t.id : 0,
        name: toStr(t && t.name),
        headers: Array.isArray(t.headers) ? t.headers.map(toStr) : [],
        rows: rows.map((r) => ({ cells: rowCells(r).map(toStr) })),
      };
    });
  }

  // ============================================================
  // 5. 读服务器状态
  // ============================================================
  async function loadExperimentDefs() {
    const r = await api('GET', API_EXPERIMENT(S.expId));
    if (r.json && r.json.code === 1 && r.json.data) {
      const { defs, source } = extractDefs(r.json.data);
      S.defs = defs;
      S.defsSource = defs ? source : '';
      // 预置单元格单独存一份：导出要回落、导入要拿它判断"这一格用户到底动没动"
      S.defRowsById = new Map();
      if (defs) {
        defs.forEach((d) => S.defRowsById.set(String(d.id), d.presetRows || []));
        const presetCells = defs.reduce((a, d) => a + (d.presetRows || []).reduce((b, r) => b + r.filter((c) => toStr(c).trim() !== '').length, 0), 0);
        log('ok', `读到表格定义：${defs.length} 张表（字段 ${source}），模板预置格 ${presetCells} 个`);
      } else {
        log('warn', '实验接口里没有找到 dataTableDefinition（字段名可能变了）——将用服务器已存数据和弹窗 DOM 兜底。');
      }
      return r.json.data;
    }
    log('err', `读取实验定义失败：HTTP ${r.status} code=${r.json && r.json.code} msg=${r.json && r.json.msg}`);
    return null;
  }

  async function loadServerData() {
    const r = await api('GET', `${API_DATA_RECORD}/${encodeURIComponent(S.expId)}`);
    const data = r.json && r.json.code === 1 ? r.json.data : null;
    const tables = normalizeServerTables(data && data.table_data);
    S.server = {
      tables,
      table_data: data ? data.table_data : null,
      update_time: data ? (data.update_time || data.create_time || '') : '',
      raw: r.json,
      rawTables: (data && data.table_data && Array.isArray(data.table_data.tables))
        ? data.table_data.tables.map((t) => ({
            id: t.id,
            rowsIsArray: Array.isArray(t.rows),
            row0IsArray: Array.isArray(t.rows && t.rows[0]),
            row0Keys: (t.rows && t.rows[0] && !Array.isArray(t.rows[0])) ? Object.keys(t.rows[0]) : null,
          }))
        : null,
    };
    if (!tables.length) {
      log('warn', `服务器暂无已保存数据（code=${r.json && r.json.code}）——首份文件将由页面表格定义生成。`);
    } else {
      log('ok', `服务器现有 ${tables.length} 张表，最后更新：${S.server.update_time || '(无时间)'}`);
      if (S.server.rawTables) {
        const shapes = S.server.rawTables.slice(0, 3).map((t) => {
          const shape = t.row0IsArray ? '裸数组' : (t.row0Keys ? `对象{${t.row0Keys.join(',')}}` : '空/未知');
          return `id=${t.id}:rows为${t.rowsIsArray ? '数组' : '非数组'}/行0为${shape}`;
        }).join('；');
        log('net', `原始行形态：${shapes}${S.server.rawTables.length > 3 ? ' …' : ''}`);
      }
    }
    return S.server;
  }

  /**
   * 确保 S.defs 可用。**优先级：服务器已存数据 > 接口定义 > 弹窗 DOM**。
   *
   * 为什么服务器已存数据优先？—— 它才是"服务器现在真正认的表"（表数、表头、行列数）。
   * 接口里的 dataTableDefinition 是"模板"，模板被老师改过之后可能与已存数据不一致；
   * 用模板当基准会导出服务器上并不存在的表（实测 exp22：模板 6 张 vs 已存 8 张），
   * 一旦把这些多出来的表 POST 回去，就可能在服务器上凭空长出表来。
   */
  async function ensureDefs() {
    if (S.server && S.server.tables.length) {
      S.defs = S.server.tables.map((t) => ({
        id: t.id,
        name: t.name,
        headers: t.headers.slice(),
        expectedRows: t.rows.length,
        _fromServer: true,
      }));
      if (!S.defsSource) S.defsSource = 'server';
      return S.defs;
    }
    if (S.defs && S.defs.length) return S.defs;
    const dom = scanDomInputs();
    if (dom) {
      S.defs = dom;
      if (!S.defsSource) S.defsSource = 'dom';
      log('warn', '表格定义来自弹窗 DOM（请确认已打开一次「填写数据」弹窗）。');
      return S.defs;
    }
    return null;
  }

  /** 接口定义与服务器已存数据的差异检查（只在日志里提示，不改变基准） */
  function checkDefsConsistency() {
    if (!S.defs || !S.server || !S.server.tables.length) return;
    const defIds = S.defs.map((d) => String(d.id));
    const srvIds = S.server.tables.map((t) => String(t.id));
    const onlyDef = defIds.filter((i) => !srvIds.includes(i));
    const onlySrv = srvIds.filter((i) => !defIds.includes(i));
    log('net', `表基准 = 服务器已存数据（${srvIds.length} 张：${srvIds.join(',')}）；接口模板 ${defIds.length} 张：${defIds.join(',')}`);
    if (onlyDef.length) log('warn', `接口模板里有、服务器已存数据里没有的表 id=${onlyDef.join(',')} → 本次不会导出，也不会写回（避免凭空加表）`);
    if (onlySrv.length) log('warn', `服务器已存数据里有、接口模板里没有的表 id=${onlySrv.join(',')} → 以服务器为准一并导出`);
  }

  async function refreshAll() {
    await loadExperimentDefs();
    await loadServerData();
    await ensureDefs();
    checkDefsConsistency();
  }

  // ============================================================
  // 6. 合并视图（导出用）
  // ============================================================
  /**
   * 导出视图：每张表的每一格取「服务器已保存值」，没有则空。
   * 同时给出 baseline（文件里存一份服务器快照，导入时做三方合并的基准）。
   */
  function buildExportView() {
    // 基准就是 S.defs（refreshAll 已保证它优先取"服务器已存数据"，与服务器表集合一致）
    const defs = (S.defs && S.defs.length) ? S.defs.slice() : [];
    const serverTables = (S.server && S.server.tables) ? S.server.tables : [];
    // 兜底：定义里没有、但服务器上有的表也一并导出（避免接口字段变动导致漏表）
    serverTables.forEach((st) => {
      if (!defs.some((d) => String(d.id) === String(st.id))) {
        defs.push({ id: st.id, name: st.name, headers: (st.headers || []).slice(), expectedRows: st.rows.length, _fromServer: true });
      }
    });
    const tables = [];
    const warnings = [];
    for (const d of defs) {
      const st = serverTables.find((t) => String(t.id) === String(d.id));
      if (!d.headers || !d.headers.length) {
        warnings.push(`表 id=${d.id} 没有表头信息，已跳过（先打开一次「填写数据」弹窗、或先刷新服务器数据）`);
        continue;
      }
      const expectedRows = Math.max(d.expectedRows || 0, st ? st.rows.length : 0);
      const cols = Math.max(d.headers.length, st ? Math.max(0, ...st.rows.map((r) => rowCells(r).length)) : 0);
      const headers = d.headers.slice();
      while (headers.length < cols) headers.push('');
      const presets = S.defRowsById.get(String(d.id)) || [];
      const rows = [];
      for (let r = 0; r < expectedRows; r++) {
        const sCells = st && st.rows[r] ? rowCells(st.rows[r]) : [];
        const pCells = presets[r] || [];
        const cells = [];
        for (let c = 0; c < cols; c++) {
          // 取值优先级与页面一致：服务器已存值 > 模板预置值 > 空
          // （学生没填的那些"序号/参数名"预置格，页面上是看得到的，导出也必须带上）
          const sv = sCells[c];
          if (sv !== undefined && sv !== null && toStr(sv) !== '') { cells.push(toStr(sv)); continue; }
          const pv = pCells[c];
          cells.push(pv === undefined || pv === null ? '' : toStr(pv));
        }
        rows.push({ cells });
      }
      if (st && st.rows.length > expectedRows) {
        warnings.push(`表 id=${d.id}：服务器有 ${st.rows.length} 行，超出定义的 ${expectedRows} 行也一并保留`);
      }
      tables.push({ id: d.id, name: st && st.name ? st.name : (d.name || ''), headers, rows });
    }
    return { tables, warnings };
  }

  function filledCount(t) {
    let n = 0;
    t.rows.forEach((r) => r.cells.forEach((c) => { if (toStr(c).trim() !== '') n++; }));
    return n;
  }

  // ============================================================
  // 7. 导出
  // ============================================================
  function buildPayload() {
    const { tables, warnings } = buildExportView();
    if (!tables.length) throw new Error('没有可导出的表（缺少表头信息）');
    const presets = tables.map((t) => {
      const pr = S.defRowsById.get(String(t.id)) || [];
      const cols = t.headers.length;
      const rows = [];
      for (let r = 0; r < t.rows.length; r++) {
        const pc = pr[r] || [];
        const cells = [];
        for (let c = 0; c < cols; c++) cells.push(pc[c] === undefined || pc[c] === null ? '' : toStr(pc[c]));
        rows.push({ cells });
      }
      return { id: t.id, rows };
    });
    return {
      file: {
        app: XLSX_MARK,
        exportedAt: new Date().toISOString(),
        host: location.host,
        expId: S.expId,
        url: location.href,
        serverUpdateTime: S.server ? S.server.update_time : '',
      },
      summary: tables.map((t) => ({
        id: t.id, name: t.name, rows: t.rows.length, cols: t.headers.length,
        filled: filledCount(t), headers: t.headers,
      })),
      tables,
      // 两层快照，缺一不可：
      //   _baseline = 导出那一刻"服务器已存值"（本轮上传的写回基准）
      //   _presets  = 模板自带的预置单元格（用来判断"用户到底动没动这一格"，
      //              否则预置格会被误判成"用户填的"，导入时就会把模板数据写进服务器）
      _baseline: tables.map((t) => ({
        id: t.id, name: t.name, headers: t.headers,
        rows: t.rows.map((r) => ({ cells: r.cells.slice() })),
      })),
      _presets: presets,
      _clearToken: CLEAR_TOKEN,
      _warnings: warnings,
    };
  }

  async function exportJSON() {
    if (S.busy) return; setBusy(true);
    try {
      await refreshAll();
      const payload = buildPayload();
      const name = buildFileName('json');
      downloadBlob(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' }), name);
      payload.warnings.forEach((w) => log('warn', w));
      const total = payload.summary.reduce((a, s) => a + s.rows * s.cols, 0);
      const filled = payload.summary.reduce((a, s) => a + s.filled, 0);
      log('ok', `已导出 ${name}（${payload.tables.length} 张表 / ${total} 格，其中已填 ${filled} 格）`);
      renderSummary(payload);
    } catch (e) {
      log('err', '导出 JSON 失败：' + e.message);
      alert('导出失败：' + e.message);
    } finally { setBusy(false); }
  }

  function sheetFromAoa(aoa, widths) {
    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: false });
    ws['!cols'] = widths.map((w) => ({ wch: Math.min(40, Math.max(8, w)) }));
    return ws;
  }

  async function exportXLSX() {
    if (S.busy) return; setBusy(true);
    try {
      const XLSX = getXLSX();
      if (!XLSX) throw new Error('SheetJS 未加载（@require 失败？）。可改用「导出 JSON」。');
      await refreshAll();
      const payload = buildPayload();
      const wb = XLSX.utils.book_new();

      // ① 说明页
      const info = [
        ['实验数据表格同步器 — 说明'],
        ['文件标记', XLSX_MARK],
        ['导出时间', new Date().toLocaleString()],
        ['服务器', location.host],
        ['实验 ID', String(S.expId)],
        ['服务器最后更新', payload.file.serverUpdateTime || '(无)'],
        [],
        ['怎么用'],
        ['1', '每张表一个独立工作表，表头第 1 行是表名，第 2 行是列名，第 3 行起是数据。'],
        ['2', '直接改数据行即可；公式请填 $LaTeX$ 原文（如 $1.23\\times10^{-4}$），不要填 Excel 公式。'],
        ['3', `想清空服务器上的某一格：把该格内容删成空。写 JSON 时也可以显式写 ${CLEAR_TOKEN}。`],
        ['4', `不要改表名、列名、以及「${BASE_SHEET}」「${PRESET_SHEET}」两张工作表，否则只能按「只补空缺」方式上传。`],
        ['5', '改完回到网页，点「导入上传」，脚本会先 GET 服务器最新数据、三方合并、再 POST 覆盖你改动的格。'],
        [],
        ['关于"模板预置值"（重要）'],
        ['P1', '有些格是老师模板里印好的（序号列、表4/表5 里印好的参数名等），它们在服务器上没有存过。'],
        ['P2', '这类格**永远不会被上传**：页面上靠"服务器值 → 模板预置值"的回落自动显示，不需要你填，也别指望改了能生效。'],
        ['P3', '想知道哪些格属于这一类，看「_预置值清单」工作表。'],
        ['P4', '如果你确实要覆盖某个预置值，就直接把那一格改成你要的值 —— 改动过的格会被正常上传。'],
        [],
        ['保真提醒'],
        ['a', '本文件同时导出 JSON（主存档，无损）。做精确计算请以 JSON 为准。'],
        ['b', 'Excel 可能把 1.5e-4 之类的文本自动转成数字/日期，读回时会做尽力还原，但不保证 100%。'],
        ['c', '列宽不够时 Excel 会显示 ####，那只是显示问题，不影响内容。'],
      ];
      wb.SheetNames.push('说明_README');
      wb.Sheets['说明_README'] = sheetFromAoa(info, [14, 100]);

      // ② 基线页 + ③ 预设页（都是三方合并的基准，勿改）
      const baseAoa = [['表ID', '行', '列', '服务器原值（导出时快照，请勿修改）']];
      payload._baseline.forEach((t) => {
        t.rows.forEach((r, ri) => r.cells.forEach((c, ci) => {
          baseAoa.push([t.id, ri, ci, toStr(c)]);
        }));
      });
      wb.SheetNames.push(BASE_SHEET);
      wb.Sheets[BASE_SHEET] = sheetFromAoa(baseAoa, [8, 6, 6, 60]);

      const presetAoa = [['表ID', '行', '列', '模板预置值（导出时快照，请勿修改）']];
      (payload._presets || []).forEach((t) => {
        t.rows.forEach((r, ri) => r.cells.forEach((c, ci) => {
          presetAoa.push([t.id, ri, ci, toStr(c)]);
        }));
      });
      wb.SheetNames.push(PRESET_SHEET);
      wb.Sheets[PRESET_SHEET] = sheetFromAoa(presetAoa, [8, 6, 6, 60]);

      // ③ 每张表一个 sheet
      const names = sanitizeSheetNames(payload.tables.map((t) => ({
        id: t.id,
        name: `T${t.id}_${t.name || '表' + t.id}`,
      })));
      payload.tables.forEach((t, idx) => {
        const sheetName = names[idx].sheet;
        const aoa = [[`表 id=${t.id}  ${t.name || ''}`.trim()]];
        aoa.push(t.headers.slice());
        t.rows.forEach((r) => aoa.push(r.cells.slice()));
        wb.SheetNames.push(sheetName);
        wb.Sheets[sheetName] = sheetFromAoa(aoa, t.headers.map((h) => Math.max(String(h || '').length + 4, 12)));
      });
      // ④ 预置值清单页：让人一眼看出"哪些格是模板自带的、永远不会被上传"
      const presetList = [['表ID', '行(从1数)', '列(从1数)', '列名', '模板预置值', '说明']];
      (payload._presets || []).forEach((t) => {
        const tbl = payload.tables.find((x) => String(x.id) === String(t.id));
        t.rows.forEach((r, ri) => r.cells.forEach((c, ci) => {
          if (toStr(c).trim() === '') return;
          presetList.push([t.id, ri + 1, ci + 1, (tbl && tbl.headers[ci]) || '', toStr(c), '模板预置：改不改都不会上传']);
        }));
      });
      wb.SheetNames.push('_预置值清单');
      wb.Sheets['_预置值清单'] = sheetFromAoa(presetList, [8, 10, 10, 30, 30, 26]);

      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
      const name = buildFileName('xlsx');
      downloadBlob(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name);
      log('ok', `已导出 ${name}（工作表：说明_README / ${BASE_SHEET} / ${names.map((n) => n.sheet).join(' / ')}）`);
      renderSummary(payload);
    } catch (e) {
      log('err', '导出 XLSX 失败：' + e.message);
      alert('导出 XLSX 失败：' + e.message);
    } finally { setBusy(false); }
  }

  function getXLSX() {
    try {
      if (typeof XLSX !== 'undefined' && XLSX && XLSX.utils) return XLSX;
    } catch (e) {}
    try {
      if (typeof unsafeWindow !== 'undefined' && unsafeWindow.XLSX && unsafeWindow.XLSX.utils) return unsafeWindow.XLSX;
    } catch (e) {}
    return null;
  }

  // ============================================================
  // 8. 导入解析（JSON / XLSX → 统一结构）
  // ============================================================
  function aoaFromSheet(XLSX, ws) {
    const ref = ws['!ref'];
    if (!ref) return [];
    const range = XLSX.utils.decode_range(ref);
    const out = [];
    for (let R = range.s.r; R <= range.e.r; R++) {
      const row = [];
      for (let C = range.s.c; C <= range.e.c; C++) {
        const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        row.push(cell ? (cell.v === null || cell.v === undefined ? '' : cell.v) : '');
      }
      out.push(row);
    }
    return out;
  }

  function normCell(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') {
      if (Number.isFinite(v) && !Number.isInteger(v)) {
        // 兜一层：Excel 把 1.5e-4 存成 0.00015 时，尽量还原成好看的文本
        const s = String(v);
        return s;
      }
      return String(v);
    }
    return String(v);
  }

  function parseFileToPayload(text, arrayBuffer, filename) {
    const isXlsx = /\.xlsx$/i.test(filename) || (arrayBuffer && !text);
    if (isXlsx) {
      const XLSX = getXLSX();
      if (!XLSX) throw new Error('SheetJS 未加载，无法读取 .xlsx；请改用 JSON 文件导入。');
      const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
      const warnings = [];
      const dataSheets = wb.SheetNames.filter((n) => /^T1?_?\d+_/.test(n) || /^T\d+_/.test(n));
      const candidates = dataSheets.length ? dataSheets : wb.SheetNames.filter((n) => n !== '说明_README' && n !== BASE_SHEET);
      const tables = [];
      candidates.forEach((sn) => {
        const aoa = aoaFromSheet(XLSX, wb.Sheets[sn]);
        let id = null;
        const m = sn.match(/^T(\d+)_/);
        if (m) id = parseInt(m[1], 10);
        if (id === null && aoa[0] && aoa[0][0] !== undefined) {
          const m2 = String(aoa[0][0]).match(/id\s*=\s*(\d+)/);
          if (m2) id = parseInt(m2[1], 10);
        }
        if (id === null) { warnings.push(`工作表「${sn}」认不出表 id，已跳过`); return; }
        const headers = (aoa[1] || []).map(normCell);
        const rows = aoa.slice(2).map((r) => ({ cells: r.map(normCell) }));
        tables.push({ id, name: '', headers, rows });
      });
      let baseline = null;
      if (wb.Sheets[BASE_SHEET]) {
        const b = aoaFromSheet(XLSX, wb.Sheets[BASE_SHEET]);
        baseline = new Map();
        b.slice(1).forEach((r) => {
          if (r.length < 4) return;
          const id = String(r[0]).trim();
          const ri = parseInt(r[1], 10), ci = parseInt(r[2], 10);
          if (id === '' || isNaN(ri) || isNaN(ci)) return;
          baseline.set(cellKey(id, ri, ci), normCell(r[3]));
        });
        if (!baseline.size) baseline = null;
      } else {
        warnings.push(`没有找到「${BASE_SHEET}」工作表（可能是别人手工做的表）：本次导入只补空缺，不会清空任何服务器数据。`);
      }
      let presets = null;
      if (wb.Sheets[PRESET_SHEET]) {
        const b = aoaFromSheet(XLSX, wb.Sheets[PRESET_SHEET]);
        presets = new Map();
        b.slice(1).forEach((r) => {
          if (r.length < 4) return;
          const id = String(r[0]).trim();
          const ri = parseInt(r[1], 10), ci = parseInt(r[2], 10);
          if (id === '' || isNaN(ri) || isNaN(ci)) return;
          presets.set(cellKey(id, ri, ci), normCell(r[3]));
        });
        if (!presets.size) presets = null;
      }
      if (!tables.length) throw new Error('这个 Excel 里没有识别到数据表（工作表名需形如 T0_表名）');
      return { source: 'xlsx', tables, baseline, presets, warnings };
    }

    // JSON
    let obj;
    try { obj = JSON.parse(text); } catch (e) { throw new Error('JSON 解析失败：' + e.message); }
    let tables = [];
    if (Array.isArray(obj)) tables = obj;
    else if (Array.isArray(obj.tables)) tables = obj.tables;
    else if (obj && obj.table_data) {
      const inner = typeof obj.table_data === 'string' ? JSON.parse(obj.table_data) : obj.table_data;
      tables = (inner && inner.tables) || [];
    }
    if (!tables.length) throw new Error('JSON 里没有 tables 数组');
    const norm = tables.map((t) => {
      const rows = Array.isArray(t.rows) ? t.rows : [];
      return {
        id: (t.id !== undefined && t.id !== null) ? t.id : 0,
        name: toStr(t.name),
        headers: Array.isArray(t.headers) ? t.headers.map(normCell) : [],
        rows: rows.map((r) => ({ cells: (Array.isArray(r && r.cells) ? r.cells : []).map(normCell) })),
      };
    });
    let baseline = null;
    const rawBase = Array.isArray(obj._baseline) ? obj._baseline : null;
    if (rawBase) {
      baseline = new Map();
      rawBase.forEach((t) => {
        (t.rows || []).forEach((r, ri) => (r.cells || []).forEach((c, ci) => {
          baseline.set(cellKey(String(t.id), ri, ci), normCell(c));
        }));
      });
      if (!baseline.size) baseline = null;
    }
    const warnings = [];
    let presets = null;
    const rawPresets = Array.isArray(obj._presets) ? obj._presets : null;
    if (rawPresets) {
      presets = new Map();
      rawPresets.forEach((t) => {
        (t.rows || []).forEach((r, ri) => (r.cells || []).forEach((c, ci) => {
          presets.set(cellKey(String(t.id), ri, ci), normCell(c));
        }));
      });
      if (!presets.size) presets = null;
    }
    if (!baseline) warnings.push('文件里没有 _baseline 快照：本次导入按「只补空缺」处理，不会清空服务器数据。');
    return { source: 'json', tables: norm, baseline, presets, warnings };
  }

  // ============================================================
  // 9. 三方合并
  // ============================================================
  /**
   * 纵深防御：无论上游给了什么形态的行（`{cells:[]}` 或裸数组 `[]`），统一归一成
   * `{cells:[...]}` —— 下游一律按 `.cells` 读，两侧形态差异到此为止。
   */
  function normRows(rows) {
    if (!Array.isArray(rows)) return [];
    return rows.map((r) => ({ cells: rowCells(r).map(toStr) }));
  }
  function serverRow(t, r) { return (t && t.rows && t.rows[r]) ? t.rows[r] : null; }
  function expCols(t) {
    if (!t) return 0;
    const headerLen = Array.isArray(t.headers) ? t.headers.length : 0;
    const rows = Array.isArray(t.rows) ? t.rows : [];
    if (!rows.length) return headerLen;
    return Math.max(...rows.map((r) => rowCells(r).length), headerLen, 0);
  }

  function computeMerge(file, server) {
    const changes = [];
    const mergedCell = new Map();
    const warnings = [];
    let stats = { mod: 0, add: 0, del: 0, tables: 0, skipped: 0, presetSkipped: 0 };

    const fileTables = (file && Array.isArray(file.tables)) ? file.tables : [];
    const serverTables = (server && Array.isArray(server.tables)) ? server.tables : [];

    for (const ft of fileTables) {
      const st = serverTables.find((t) => String(t.id) === String(ft.id));
      if (!st) {
        warnings.push(`表 id=${ft.id} 在服务器上不存在，整表跳过（不动服务器）`);
        stats.skipped++;
        continue;
      }
      const fHeaders = Array.isArray(ft.headers) ? ft.headers : [];
      const sHeaders = Array.isArray(st.headers) ? st.headers : [];
      if (fHeaders.length && sHeaders.length && fHeaders.length !== sHeaders.length) {
        warnings.push(`表 id=${ft.id}：列数不一致（文件 ${fHeaders.length} / 服务器 ${sHeaders.length}），多出的列会被忽略`);
      }
      stats.tables++;
      const fRows = normRows(ft.rows);
      const sRows = normRows(st.rows);
      // curPresetRows = **当前模板**的预置行（预置格判定的权威来源）；
      // ft.presetRows 只在拿不到当前模板时兜底。
      const curPresetRows = S.defRowsById.get(String(ft.id)) || ft.presetRows || [];
      const rows = Math.max(fRows.length, sRows.length);
      const cols = Math.max(expCols({ headers: fHeaders, rows: fRows }), expCols({ headers: sHeaders, rows: sRows }));
      for (let r = 0; r < rows; r++) {
        const frow = fRows[r];
        const srow = serverRow({ rows: sRows }, r);
        const prow = curPresetRows[r] || [];
        for (let c = 0; c < cols; c++) {
          const fv = frow && c < frow.cells.length ? toStr(frow.cells[c]) : '';
          const sv = srow && c < srow.cells.length ? toStr(srow.cells[c]) : '';
          const curPv = prow[c] === undefined || prow[c] === null ? '' : toStr(prow[c]);
          const key = cellKey(ft.id, r, c);
          const bv = file.baseline ? file.baseline.get(key) : undefined;
          // 文件里带的 _presets 快照（= 导出那一刻模板在这一格的值）。取不到就当空：
          // 只有"导出时为空、现在模板有值"才是模板新加的预置格，不能拿当前模板值反过来冒充快照。
          const fpv = (file.presets && mapLike(file.presets).has(key)) ? toStr(mapLike(file.presets).get(key)) : '';
          const hasBaseline = !!file.baseline;

          if (fv === CLEAR_TOKEN) {
            // 显式清空同样遵守"预置格不入库"：服务器没存过、只是模板预置值 → 清空无意义，不写
            if (sv !== '') { mergedCell.set(key, ''); changes.push({ t: ft.id, r, c, from: sv, to: '', kind: 'del' }); stats.del++; }
            continue;
          }
          if (hasBaseline) {
            const base = bv === undefined ? '' : bv;
            // 【预置格永不入库】这一格属于"模板预置、服务器上从没存过"时，文件里的值不管是什么，
            // 只要不是用户改过的痕迹，就一律不上传：页面上靠"服务器值 → 模板预置值"回落自然显示。
            // 判定必须用**当前模板**（curPv）而不是文件里的 _预设 快照，因为模板可能在导出之后被改过。
            // ⚠️ 必须排在 `fv === base` 之前：_baseline 存的是"导出视图"（含预置回落值），
            //    若先比基线，未改动的预置格会被当成"没动"、连计数都记不上。
            if (sv === '' && curPv !== '' && (fv === '' || fv === curPv)) {
              stats.presetSkipped++;
              continue;
            }
            if (fv === base) continue;                 // 与「服务器快照」一致 → 用户没动过
            // 兜底：导出之后模板才加上的预置值（此时 _baseline 里是空的）也算模板数据
            if (sv === '' && curPv !== '' && fpv === '') {
              stats.presetSkipped++;
              continue;
            }
            if (fv === '') {
              // 改成空 → 「清空这一格」。只清服务器上真实存过的值：
              if (sv !== '') { mergedCell.set(key, ''); changes.push({ t: ft.id, r, c, from: sv, to: '', kind: 'del' }); stats.del++; }
              // 若 sv === ''（这一格只是模板预置值），上面的预置分支已 continue；这里到不了。
              continue;
            }
            if (fv === sv) continue;                   // 服务器已是最新，无需上传
            mergedCell.set(key, fv);
            // 归类看"服务器上原来有没有值"，而不是看基线（基线里可能是预置回落值）
            changes.push({ t: ft.id, r, c, from: sv, to: fv, kind: sv === '' ? 'add' : 'mod' });
            stats[sv === '' ? 'add' : 'mod']++;
          } else {
            // 无基线（手工 Excel）：绝不覆盖服务器已有值，只补空缺
            if (sv === '' && curPv !== '' && fv === curPv) {
              // 模板预置格不算"空缺"：它本来就不该入库
              stats.presetSkipped++;
              continue;
            }
            if (fv !== '' && sv === '') {
              mergedCell.set(key, fv);
              changes.push({ t: ft.id, r, c, from: sv, to: fv, kind: 'add' });
              stats.add++;
            } else if (fv !== '' && sv !== '' && fv !== sv) {
              warnings.push(`表 id=${ft.id} 第 ${r + 1} 行第 ${c + 1} 列：文件值「${fv}」与服务器值「${sv}」冲突，已保留服务器值（无基线，不覆盖）`);
            }
          }
        }
      }
    }
    return { changes, mergedCell, warnings, stats };
  }

  /** 取某张表在表头上该用的 headers：定义优先，其次服务器，最后文件 */
  function pickHeaders(defsHeaders, serverHeaders, fileHeaders) {
    const cands = [defsHeaders, serverHeaders, fileHeaders];
    let best = [];
    for (const h of cands) {
      if (Array.isArray(h) && h.length > best.length) best = h;
    }
    return best.map(toStr);
  }

  function buildUploadTables(server, mergedCell, defs) {
    const defsList = Array.isArray(defs) ? defs : [];
    const cell = mapLike(mergedCell);
    // 以服务器全量为底，逐格覆盖改动 → 绝不触碰未改动/其他表
    return server.tables.map((t) => {
      const rawRows = (Array.isArray(t.rows) ? t.rows : []).map((r) => rowCells(r).map(toStr));
      const defHeaders = (() => {
        const d = defsList.find((x) => String(x.id) === String(t.id));
        return d && Array.isArray(d.headers) ? d.headers : null;
      })();
      const headers = pickHeaders(defHeaders, t.headers, null);
      // 宽度/高度要取「服务器已有」和「本次要写」的大者，否则文件新增的行列会被静默丢掉
      let maxKeyCol = -1, maxKeyRow = -1;
      const tid = String(t.id);
      cell.forEach((_v, k) => {
        const p = String(k).split('|');
        if (p.length === 3 && p[0] === tid) {
          const rr = parseInt(p[1], 10), cc = parseInt(p[2], 10);
          if (!isNaN(rr)) maxKeyRow = Math.max(maxKeyRow, rr);
          if (!isNaN(cc)) maxKeyCol = Math.max(maxKeyCol, cc);
        }
      });
      const cols = Math.max(headers.length, maxKeyCol + 1, ...rawRows.map((r) => r.length), 0);
      const rowCount = Math.max(rawRows.length, maxKeyRow + 1, 0);
      const rows = [];
      for (let ri = 0; ri < rowCount; ri++) {
        const src = rawRows[ri] || [];
        const cells = [];
        for (let ci = 0; ci < cols; ci++) {
          const k = cellKey(t.id, ri, ci);
          cells.push(cell.has(k) ? cell.get(k) : toStr(src[ci]));
        }
        rows.push({ cells });
      }
      while (headers.length < cols) headers.push('');
      return { id: t.id, name: toStr(t.name), headers, rows };
    });
  }

  function buildStrictTables(file, server, defs) {
    const serverTables = (server && Array.isArray(server.tables)) ? server.tables : [];
    const defsList = Array.isArray(defs) ? defs : [];
    return (file && Array.isArray(file.tables) ? file.tables : []).map((ft) => {
      const st = serverTables.find((t) => String(t.id) === String(ft.id));
      const d = defsList.find((x) => String(x.id) === String(ft.id));
      // 严格覆盖：列数以**文件**为准（定义只用来补表头文字），否则会把定义里有、文件里没有的列一起清空
      const fRows = normRows(ft.rows);
      const fileCols = Math.max(
        Array.isArray(ft.headers) ? ft.headers.length : 0,
        ...fRows.map((r) => r.cells.length), 0);
      const headers = pickHeaders(d ? d.headers : null, st ? st.headers : null, ft.headers).slice(0, fileCols);
      while (headers.length < fileCols) headers.push('');
      const cols = fileCols;
      return {
        id: ft.id,
        name: toStr(ft.name) || (st ? toStr(st.name) : ''),
        headers,
        rows: fRows.map((r) => {
          const cells = [];
          for (let c = 0; c < cols; c++) {
            const v = c < r.cells.length ? toStr(r.cells[c]) : '';
            cells.push(v === CLEAR_TOKEN ? '' : v);
          }
          return { cells };
        }),
      };
    });
  }

  // ============================================================
  // 10. 备份 / 恢复
  // ============================================================
  function backupHistory() {
    try { return JSON.parse(localStorage.getItem(BAK_KEY) || '[]'); } catch (e) { return []; }
  }
  function pushBackup(expId, server, note) {
    const list = backupHistory();
    list.unshift({
      at: new Date().toISOString(),
      expId: String(expId),
      note: note || '',
      updateTime: server ? server.update_time : '',
      tables: server ? server.tables : [],
    });
    // 同一实验只留 6 份，全局只留 20 份
    const sameExp = list.filter((b) => b.expId === String(expId)).slice(0, 6);
    const others = list.filter((b) => b.expId !== String(expId)).slice(0, 14);
    try { localStorage.setItem(BAK_KEY, JSON.stringify(sameExp.concat(others))); } catch (e) { log('warn', '备份写入 localStorage 失败（可能已满）'); }
  }

  async function rollback() {
    const list = backupHistory().filter((b) => b.expId === String(S.expId));
    if (!list.length) { alert('没有找到本实验的备份。'); return; }
    const lines = list.map((b, i) => `${i + 1}. ${new Date(b.at).toLocaleString()}  ${b.note || ''}  服务器时间:${b.updateTime || '-'}`);
    const pick = prompt('选择要回滚到的备份序号（会把该备份的表格数据整份写回服务器）：\n\n' + lines.join('\n'), '1');
    const idx = parseInt(pick, 10) - 1;
    if (isNaN(idx) || idx < 0 || idx >= list.length) return;
    const b = list[idx];
    if (!confirm(`确认把 ${new Date(b.at).toLocaleString()} 的备份整份写回服务器？\n（当前服务器数据会先被自动备份）`)) return;

    setBusy(true);
    try {
      await refreshAll();
      pushBackup(S.expId, S.server, '回滚前的自动备份');
      const r = await api('POST', API_DATA_RECORD, { expId: parseInt(S.expId, 10) || S.expId, tableData: { tables: b.tables } });
      if (r.json && r.json.code === 1) {
        log('ok', '回滚成功。正在回读校验…');
        await loadServerData();
        alert('回滚成功。');
      } else {
        log('err', `回滚失败：code=${r.json && r.json.code} msg=${r.json && r.json.msg}`);
        alert('回滚失败：' + ((r.json && r.json.msg) || ('HTTP ' + r.status)));
      }
    } catch (e) {
      log('err', '回滚异常：' + e.message);
      alert('回滚异常：' + e.message);
    } finally { setBusy(false); }
  }

  // ============================================================
  // 11. 导入 + 上传 + 校验
  // ============================================================
  let pendingImport = null;
  // 防止"确认上传"被连点两次：必须与会话级的 S.busy 分开 —— 弹窗期间的 S.busy 是故意保持 true 的
  // （用来挡住重复点「导入上传」），若用它守卫 commitImport，就会把自己的确认点击挡回去。
  let committing = false;

  async function importFile(file) {
    if (S.busy) return;
    if (!file) return;
    if (file.size > 8 * 1024 * 1024) { alert('文件太大（>8MB），已拒绝。'); return; }
    setBusy(true);
    try {
      log('net', `读取文件 ${file.name}（${file.size} B）`);
      const isXlsx = /\.xlsx$/i.test(file.name);
      let text = null, buf = null;
      if (isXlsx) buf = await file.arrayBuffer();
      else text = await file.text();

      await refreshAll();
      const filePayload = parseFileToPayload(text, buf, file.name);
      filePayload.warnings.forEach((w) => log('warn', w));
      log('ok', `解析出 ${filePayload.tables.length} 张表${filePayload.baseline ? '（带基线，三方合并）' : '（无基线，只补空缺）'}`);

      if (strictMode()) {
        log('warn', '已开启「严格覆盖」：将整份按文件内容写入服务器（含空单元格）！');
        pendingImport = { file: filePayload, mode: 'strict', server: S.server };
      } else {
        const merge = computeMerge(filePayload, S.server);
        merge.warnings.forEach((w) => log('warn', w));
        pendingImport = { file: filePayload, mode: 'merge', server: S.server, merge };
      }
      renderImportPreview(pendingImport);
      // 注意：这里 setBusy(false) 在 finally 里执行 —— 弹窗期间仍保持"忙"，防止重复点「导入上传」。
      // 用户在弹窗上的点击不受影响（按钮不在 wts-btn-* 名单里）。
      await showPreviewAndCommit(file.name);
    } catch (e) {
      log('err', '导入失败：' + e.message);
      alert('导入失败：' + e.message);
    } finally { setBusy(false); }
  }

  /** 弹出一个纯 HTML 的差异预览，确认后才真正 POST */
  function renderImportPreview(p) {
    const el = document.getElementById('wts-preview');
    if (!el) return;
    let html = '';
    if (p.mode === 'strict') {
      html += `<div class="wts-warn">严格覆盖模式：将把文件里的 ${p.file.tables.length} 张表整份写入（文件里空的格会被清空）。</div>`;
    } else {
      const s = p.merge.stats;
      html += `<div>将修改 <b>${s.mod}</b> 格 · 新增 <b>${s.add}</b> 格 · 清空 <b>${s.del}</b> 格，涉及 <b>${s.tables}</b> 张表。</div>`;
      if (s.presetSkipped) {
        html += `<div class="wts-warn">另有 <b>${s.presetSkipped}</b> 格是「模板预置值」（如序号列、模板里印好的参数名），
          按约定<b>不入库</b>：页面上会靠模板自动回落显示，不需要上传。</div>`;
      }
      if (!p.merge.changes.length) html += `<div class="wts-warn">没有检测到任何改动 —— 无需上传。</div>`;
      const byTable = new Map();
      p.merge.changes.forEach((c) => {
        if (!byTable.has(c.t)) byTable.set(c.t, []);
        byTable.get(c.t).push(c);
      });
      byTable.forEach((list, tid) => {
        html += `<div class="wts-t">表 id=${tid}：${list.length} 处改动</div><ul>`;
        list.slice(0, 40).forEach((c) => {
          const kind = c.kind === 'del' ? '清空' : (c.kind === 'add' ? '新增' : '修改');
          html += `<li>[${kind}] (第${c.r + 1}行,第${c.c + 1}列)　${esc(c.from) || '<i>(空)</i>'} → <b>${esc(c.to) || '<i>(空)</i>'}</b></li>`;
        });
        if (list.length > 40) html += `<li>…还有 ${list.length - 40} 处</li>`;
        html += '</ul>';
      });
    }
    el.innerHTML = html;
  }

  function showPreviewAndCommit(filename) {
    return new Promise((resolve) => {
      const modal = document.getElementById('wts-modal');
      const title = document.getElementById('wts-modal-title');
      const okBtn = document.getElementById('wts-modal-ok');
      const cancelBtn = document.getElementById('wts-modal-cancel');
      if (title) title.textContent = `导入预览：${filename}`;
      modal.style.display = 'flex';
      // 按钮文案按模式设置（必须在绑定前设好，避免"点太快读到旧文案"的错觉）
      okBtn.textContent = strictMode() ? '确认严格覆盖上传' : '确认上传改动';
      let settled = false;
      const onOk = async () => {
        if (settled) return;
        settled = true;
        okBtn.disabled = true;
        cancelBtn.disabled = true;
        okBtn.textContent = '上传中…';
        try {
          await commitImport();          // 注意：这里不能用 S.busy 当守门员 ——
        } finally {                      // importFile 期间 S.busy 故意是 true
          cleanup();
          resolve(true);
        }
      };
      const onCancel = () => {
        if (settled) return;
        settled = true;
        cleanup();
        log('warn', '已取消导入，服务器未改动。');
        resolve(false);
      };
      function cleanup() {
        modal.style.display = 'none';
        okBtn.disabled = false;
        cancelBtn.disabled = false;
        okBtn.removeEventListener('click', onOk);
        cancelBtn.removeEventListener('click', onCancel);
      }
      okBtn.addEventListener('click', onOk);
      cancelBtn.addEventListener('click', onCancel);
    });
  }

  async function commitImport() {
    const p = pendingImport;
    if (!p) return;
    if (committing) { log('warn', '正在上传中，忽略重复点击。'); return; }
    committing = true;
    setBusy(true);
    try {
      // 1) 备份服务器现状
      pushBackup(S.expId, p.server, p.mode === 'strict' ? '严格覆盖前' : '合并上传前');
      log('ok', '已保存上传前备份（可用「回滚上次」恢复）');

      // 2) 组装要写的表
      let tables;
      if (p.mode === 'strict') {
        tables = buildStrictTables(p.file, p.server, S.defs);
        if (!tables.length) throw new Error('严格覆盖：文件里没有有效的表');
        if (!confirm(`严格覆盖会写 ${tables.length} 张表，文件里为空的格将被清空。确定继续？`)) {
          log('warn', '已放弃严格覆盖。');
          return;
        }
      } else {
        if (!p.merge.changes.length) { log('warn', '无改动，未上传。'); return; }
        tables = buildUploadTables(p.server, p.merge.mergedCell, S.defs);
      }

      // 3) POST
      const body = { expId: parseInt(S.expId, 10) || S.expId, tableData: { tables } };
      const r = await api('POST', API_DATA_RECORD, body);
      S.lastResponse = r.json;
      if (!(r.json && r.json.code === 1)) {
        log('err', `上传失败：HTTP ${r.status} code=${r.json && r.json.code} msg=${r.json && r.json.msg}`);
        alert('上传失败：' + ((r.json && r.json.msg) || ('HTTP ' + r.status)));
        return;
      }
      log('ok', `上传成功（code=1）${r.json.msg ? ' msg=' + r.json.msg : ''}`);

      // 4) 回读校验
      log('net', '正在回读服务器数据做校验…');
      await sleep(350);
      const after = await loadServerData();
      const expectCell = new Map();
      tables.forEach((t) => t.rows.forEach((row, ri) => row.cells.forEach((v, ci) => expectCell.set(cellKey(t.id, ri, ci), toStr(v)))));
      let bad = 0; const samples = [];
      for (const t of after.tables) {
        t.rows.forEach((row, ri) => row.cells.forEach((v, ci) => {
          const want = expectCell.get(cellKey(t.id, ri, ci));
          if (want === undefined) return;
          if (toStr(v) !== want) { bad++; if (samples.length < 5) samples.push(`表${t.id}(${ri + 1},${ci + 1}) 期望「${want}」实际「${v}」`); }
        }));
      }
      if (after.tables.length < tables.length) {
        bad += 1; samples.push(`回读只有 ${after.tables.length} 张表，期望 ${tables.length} 张`);
      }
      if (bad === 0) {
        log('ok', '✅ 回读校验通过：服务器数据与上传内容一致。');
        alert('上传成功，且回读校验通过。');
      } else {
        log('err', `⚠️ 回读校验发现 ${bad} 处不一致：\n` + samples.join('\n'));
        alert(`上传返回成功，但回读发现 ${bad} 处不一致（详见面板日志）。建议用「回滚上次」恢复。`);
      }
      pendingImport = null;
    } catch (e) {
      log('err', '上传异常：' + e.message);
      alert('上传异常：' + e.message);
    } finally { committing = false; setBusy(false); }
  }

  function strictMode() {
    const cb = document.getElementById('wts-strict');
    return !!(cb && cb.checked);
  }

  // ============================================================
  // 12. 面板 UI
  // ============================================================
  const CSS = `
  #wts-panel, #wts-modal { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif; }
  #wts-panel { position: fixed; top: 90px; right: 18px; width: 400px; max-height: 72vh; z-index: 2147483000;
    background: #fff; border: 1px solid #d7dbe0; border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,.18);
    display: flex; flex-direction: column; overflow: hidden; font-size: 13px; color: #1f2937; }
  #wts-panel.wts-collapsed { max-height: 44px; width: 250px; }
  #wts-head { display: flex; align-items: center; gap: 8px; padding: 10px 12px; background: #165DFF; color: #fff;
    cursor: move; user-select: none; font-weight: 600; }
  #wts-head .wts-dot { width: 8px; height: 8px; border-radius: 50%; background: #a7f3d0; flex: none; }
  #wts-head .wts-exp { font-weight: 400; opacity: .9; font-size: 12px; }
  #wts-head button { margin-left: auto; background: rgba(255,255,255,.18); border: 0; color: #fff; border-radius: 6px;
    padding: 2px 8px; cursor: pointer; font-size: 12px; }
  #wts-body { padding: 10px 12px; overflow: auto; }
  #wts-panel.wts-collapsed #wts-body { display: none; }
  .wts-row { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; }
  .wts-row button { flex: 1 1 auto; border: 1px solid #cbd5e1; background: #f8fafc; border-radius: 8px; padding: 6px 8px;
    cursor: pointer; font-size: 12.5px; color: #0f172a; }
  .wts-row button:hover { background: #eef2ff; border-color: #93c5fd; }
  .wts-row button.wts-primary { background: #165DFF; border-color: #165DFF; color: #fff; }
  .wts-row button.wts-primary:hover { background: #1249c9; }
  .wts-row button.wts-danger { color: #b91c1c; border-color: #fca5a5; background: #fef2f2; }
  .wts-row button:disabled { opacity: .5; cursor: not-allowed; }
  #wts-summary { font-size: 12px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 8px; margin-bottom: 8px; }
  #wts-summary table { border-collapse: collapse; width: 100%; }
  #wts-summary td, #wts-summary th { border-bottom: 1px solid #e5e7eb; padding: 2px 4px; text-align: left; font-size: 11.5px; }
  #wts-log { font-family: ui-monospace, Consolas, monospace; font-size: 11px; line-height: 1.5; background: #0f172a; color: #cbd5e1;
    border-radius: 8px; padding: 8px; height: 150px; overflow: auto; white-space: pre-wrap; word-break: break-all; }
  #wts-log .l-warn { color: #fbbf24; } #wts-log .l-err { color: #f87171; } #wts-log .l-ok { color: #4ade80; } #wts-log .l-net { color: #7dd3fc; }
  #wts-conf { display: flex; align-items: center; gap: 6px; font-size: 11.5px; color: #475569; margin-bottom: 6px; flex-wrap: wrap; }
  #wts-conf input[type=text] { width: 96px; padding: 3px 6px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 11.5px; }
  #wts-modal { position: fixed; inset: 0; background: rgba(15,23,42,.55); z-index: 2147483001; display: none;
    align-items: center; justify-content: center; padding: 20px; }
  #wts-modal .wts-box { background: #fff; border-radius: 12px; width: min(760px, 96vw); max-height: 86vh; display: flex; flex-direction: column; overflow: hidden; }
  #wts-modal h3 { margin: 0; padding: 12px 16px; border-bottom: 1px solid #e5e7eb; font-size: 15px; }
  #wts-preview { padding: 12px 16px; overflow: auto; font-size: 13px; }
  #wts-preview .wts-warn { color: #b45309; background: #fffbeb; border: 1px solid #fde68a; border-radius: 6px; padding: 6px 8px; margin: 6px 0; }
  #wts-preview .wts-t { margin-top: 10px; font-weight: 600; }
  #wts-preview ul { margin: 4px 0 0 18px; padding: 0; }
  #wts-preview li { margin: 2px 0; }
  #wts-modal .wts-foot { display: flex; gap: 8px; justify-content: flex-end; padding: 12px 16px; border-top: 1px solid #e5e7eb; }
  #wts-modal button { border-radius: 8px; padding: 8px 14px; cursor: pointer; border: 1px solid #cbd5e1; background: #f8fafc; font-size: 13px; }
  #wts-modal button.wts-primary { background: #165DFF; border-color: #165DFF; color: #fff; }
  `;

  function log(level, msg) {
    const time = new Date().toLocaleTimeString();
    S.logLines.push({ level, msg, time });
    if (S.logLines.length > 400) S.logLines.shift();
    const el = document.getElementById('wts-log');
    if (el) {
      const line = document.createElement('div');
      line.className = 'l-' + level;
      line.textContent = `[${time}] ${msg}`;
      el.appendChild(line);
      el.scrollTop = el.scrollHeight;
    }
    if (level === 'err') console.error('[WTS]', msg); else console.log('[WTS]', msg);
  }

  function setBusy(b) {
    S.busy = b;
    ['wts-btn-json', 'wts-btn-xlsx', 'wts-btn-import', 'wts-btn-refresh', 'wts-btn-rollback'].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.disabled = b;
    });
  }

  function renderSummary(payload) {
    const el = document.getElementById('wts-summary');
    if (!el || !payload) return;
    let html = '<table><tr><th>表</th><th>尺寸</th><th>已填</th></tr>';
    payload.summary.forEach((s) => {
      html += `<tr><td>id=${s.id}${s.name ? ' ' + esc(s.name.slice(0, 12)) : ''}</td><td>${s.rows}×${s.cols}</td><td>${s.filled}</td></tr>`;
    });
    html += '</table>';
    if (S.server && S.server.update_time) html += `<div style="margin-top:4px;color:#64748b">服务器最后更新：${esc(S.server.update_time)}</div>`;
    if (payload.file && payload.file.serverUpdateTime) html += `<div style="color:#64748b">导出时服务器快照：${esc(payload.file.serverUpdateTime)}</div>`;
    el.innerHTML = html;
  }

  function buildPanel() {
    if (document.getElementById('wts-panel')) return;
    const style = document.createElement('style');
    style.textContent = CSS;
    document.head.appendChild(style);

    const panel = document.createElement('div');
    panel.id = 'wts-panel';
    panel.innerHTML = `
      <div id="wts-head">
        <span class="wts-dot"></span>
        <span>表格同步器</span>
        <span class="wts-exp">exp ${esc(S.expId || '?')} · v${esc(scriptVersion())}</span>
        <button id="wts-toggle" title="折叠/展开">—</button>
      </div>
      <div id="wts-body">
        <div id="wts-conf">
          <label>文件名前缀 <input type="text" id="wts-prefix" value="${esc(filePrefix())}"></label>
          <label>自定义文件名 <input type="text" id="wts-filename" placeholder="留空=带时间戳" style="width:118px" value="${esc((loadCfg().fileName) || '')}"></label>
          <label><input type="checkbox" id="wts-strict"> 严格覆盖（危险）</label>
        </div>
        <div class="wts-row">
          <button id="wts-btn-json" class="wts-primary">导出 JSON（主存档）</button>
          <button id="wts-btn-xlsx">导出 XLSX</button>
        </div>
        <div class="wts-row">
          <button id="wts-btn-import" class="wts-primary">导入上传</button>
          <button id="wts-btn-refresh">刷新服务器数据</button>
        </div>
        <div class="wts-row">
          <button id="wts-btn-rollback" class="wts-danger">回滚上次</button>
          <button id="wts-btn-diag">诊断</button>
          <button id="wts-btn-copy">复制日志</button>
        </div>
        <div id="wts-summary">尚未读取。</div>
        <div id="wts-log"></div>
        <input type="file" id="wts-file" accept=".json,.xlsx,application/json" style="display:none">
      </div>`;
    document.body.appendChild(panel);

    const modal = document.createElement('div');
    modal.id = 'wts-modal';
    modal.innerHTML = `
      <div class="wts-box">
        <h3 id="wts-modal-title">导入预览</h3>
        <div id="wts-preview"></div>
        <div class="wts-foot">
          <button id="wts-modal-cancel">取消</button>
          <button id="wts-modal-ok" class="wts-primary">确认上传改动</button>
        </div>
      </div>`;
    document.body.appendChild(modal);

    // 事件
    document.getElementById('wts-toggle').onclick = () => panel.classList.toggle('wts-collapsed');
    document.getElementById('wts-btn-json').onclick = exportJSON;
    document.getElementById('wts-btn-xlsx').onclick = exportXLSX;
    document.getElementById('wts-btn-refresh').onclick = async () => {
      setBusy(true);
      try { await refreshAll(); const p = buildPayload(); renderSummary(p); log('ok', '已刷新服务器数据。'); }
      catch (e) { log('err', '刷新失败：' + e.message); }
      finally { setBusy(false); }
    };
    document.getElementById('wts-btn-rollback').onclick = rollback;
    document.getElementById('wts-btn-copy').onclick = () => {
      const txt = S.logLines.map((l) => `[${l.time}] ${l.msg}`).join('\n');
      if (typeof GM_setClipboard === 'function') { GM_setClipboard(txt, 'text'); log('ok', '日志已复制到剪贴板。'); }
      else { console.log(txt); alert('日志已打印到控制台。'); }
    };
    document.getElementById('wts-btn-diag').onclick = runDiagnostics;
    document.getElementById('wts-btn-import').onclick = () => document.getElementById('wts-file').click();
    document.getElementById('wts-file').onchange = (e) => {
      const f = e.target.files && e.target.files[0];
      e.target.value = '';
      if (f) importFile(f);
    };
    const prefixInput = document.getElementById('wts-prefix');
    prefixInput.onchange = () => { saveCfg({ prefix: prefixInput.value.trim() }); log('ok', '文件名前缀已保存：' + filePrefix()); };
    const fileNameInput = document.getElementById('wts-filename');
    fileNameInput.onchange = () => {
      const v = fileNameInput.value.trim();
      saveCfg({ fileName: v });
      log('ok', v ? `自定义文件名已保存：${v}.json / ${v}.xlsx（同名会覆盖旧文件，不会再堆一堆时间戳副本）`
                  : '已恢复「前缀 + 时间戳」命名。');
    };
    const strict = document.getElementById('wts-strict');
    strict.checked = !!loadCfg().strict;
    strict.onchange = () => { saveCfg({ strict: strict.checked }); if (strict.checked) log('warn', '严格覆盖已开启：导入时会把文件内容整份写入（空单元格会清空服务器数据）。'); };

    // 拖动
    (function makeDraggable() {
      const head = document.getElementById('wts-head');
      let sx = 0, sy = 0, ox = 0, oy = 0, dragging = false;
      const saved = loadCfg().panelPos;
      if (saved) { panel.style.right = 'auto'; panel.style.left = saved.left + 'px'; panel.style.top = saved.top + 'px'; }
      head.addEventListener('mousedown', (e) => {
        if (e.target.tagName === 'BUTTON') return;
        dragging = true;
        const r = panel.getBoundingClientRect();
        sx = e.clientX; sy = e.clientY; ox = r.left; oy = r.top;
        e.preventDefault();
      });
      window.addEventListener('mousemove', (e) => {
        if (!dragging) return;
        const left = Math.max(4, Math.min(window.innerWidth - 80, ox + e.clientX - sx));
        const top = Math.max(4, Math.min(window.innerHeight - 40, oy + e.clientY - sy));
        panel.style.right = 'auto';
        panel.style.left = left + 'px';
        panel.style.top = top + 'px';
      });
      window.addEventListener('mouseup', () => {
        if (!dragging) return;
        dragging = false;
        saveCfg({ panelPos: { left: parseInt(panel.style.left, 10) || 4, top: parseInt(panel.style.top, 10) || 4 } });
      });
    })();
  }

  // ============================================================
  // 13. 诊断
  // ============================================================
  async function runDiagnostics() {
    setBusy(true);
    log('ok', `================ 诊断开始（脚本 v${typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version ? GM_info.script.version : '?'}）================`);
    try {
      // 版本自检：老版本没有这些函数，一眼就能看出装的是旧脚本
      const feats = {
        rowCells: typeof rowCells === 'function',
        mapLike: typeof mapLike === 'function',
        checkDefsConsistency: typeof checkDefsConsistency === 'function',
      };
      log(Object.values(feats).every(Boolean) ? 'ok' : 'err',
        `本脚本能力自检：${Object.entries(feats).map(([k, v]) => k + '=' + (v ? '有' : '缺')).join(' ')}`);
      if (!feats.rowCells) log('err', '！！ 你在浏览器里装的是**旧版脚本**（缺 rowCells/mapLike 修复）——请重新复制本文件并覆盖旧脚本后刷新页面。');
      log('net', `页面 URL：${location.href}`);
      log('net', `URL 里的 id = ${S.expId}`);
      log('net', `SheetJS = ${getXLSX() ? '已加载 (' + (getXLSX().version || '?') + ')' : '未加载'}`);
      log('net', `GM_xmlhttpRequest = ${typeof GM_xmlhttpRequest === 'function' ? '可用' : '不可用'}`);
      const me = await api('GET', '/api/student/info');
      const okLogin = me.json && me.json.code === 1;
      log(okLogin ? 'ok' : 'err', `登录状态：${okLogin ? '已登录（' + (me.json.data && me.json.data.username) + '）' : '未登录 / 会话过期'}`);

      const expData = await loadExperimentDefs();
      if (expData) {
        const keys = Object.keys(expData).filter((k) => /table|data/i.test(k));
        log('net', `实验接口里与表格相关的字段：${keys.join(', ') || '(无)'}`);
        if (!S.defs) {
          const cands = keys.filter((k) => typeof expData[k] === 'string' && expData[k].indexOf('headers') >= 0);
          log(cands.length ? 'warn' : 'err', cands.length ? `疑似表格定义字段（未识别）：${cands.join(', ')}` : '没找到任何含 headers 的字符串字段');
        }
      }
      await loadServerData();
      if (S.server && S.server.rawTables) {
        S.server.rawTables.forEach((t) => {
          const shape = t.row0IsArray ? '裸数组 [...]' : (t.row0Keys ? `对象 {${t.row0Keys.join(',')}}` : '空/未知');
          log('net', `  服务器表 id=${t.id}：rows 是数组=${t.rowsIsArray}，第 0 行形态=${shape}`);
        });
      }
      const domInputs = document.querySelectorAll('#fillableDataTableEditor input.data-table-input').length;
      log('net', `弹窗 DOM 里的 input 数：${domInputs}（未打开弹窗时为 0，正常）`);
      const domDefs = scanDomInputs();
      if (domDefs) log('net', 'DOM 兜底表结构：' + domDefs.map((d) => `id=${d.id} ${d.expectedRows}×${d.headers.length}`).join('，'));
      const p = buildExportView();
      log('ok', `导出视图：${p.tables.length} 张表 / ${p.tables.reduce((a, t) => a + t.rows.length * t.headers.length, 0)} 格`);
      p.warnings.forEach((w) => log('warn', w));
      const baks = backupHistory().filter((b) => b.expId === String(S.expId));
      log('net', `本地备份：${baks.length} 份${baks.length ? '，最近 ' + new Date(baks[0].at).toLocaleString() : ''}`);
    } catch (e) {
      log('err', '诊断异常：' + e.message);
    } finally {
      log('ok', '================ 诊断结束 ================');
      setBusy(false);
    }
  }

  // ============================================================
  // 14. 入口
  // ============================================================
  async function main() {
    S.expId = getUrlParam('id');
    buildPanel();
    log('ok', `表格同步器 v${scriptVersion()} 已加载，实验 id=${S.expId}`);
    if (!/^experiment\.html/i.test(location.pathname)) {
      log('warn', '当前不是 experiment.html 页面，部分功能可能无效。');
    }
    if (typeof GM_registerMenuCommand === 'function') {
      GM_registerMenuCommand('导出 JSON', exportJSON);
      GM_registerMenuCommand('导出 XLSX', exportXLSX);
      GM_registerMenuCommand('诊断', runDiagnostics);
    }
    // 静默预读一次，便于面板直接显示尺寸
    try {
      await refreshAll();
      renderSummary(buildPayload());
      log('ok', '首次读取完成，可以导出/导入了。');
    } catch (e) {
      log('warn', '首次预读失败（可能未登录）：' + e.message + '　→ 可点「诊断」排查。');
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', main);
  else main();
})();
