/**
 * SheetJS 多来源加载器自测（不发真实网络请求：用假的 <script> 注入）。
 * 用法： node tests/test-xlsx-loader.js
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { parseHTML } = require('linkedom');
const REAL_XLSX = require('xlsx');

const SRC = path.join(__dirname, '..', 'webpage-table-sync.user.js');
const src = fs.readFileSync(SRC, 'utf8');
const marker = "if (document.readyState === 'loading')";
const body = src.slice(0, src.indexOf(marker)) +
  '\nreturn { ensureXLSX, getXLSX, XLSX_CDN_CANDIDATES };\n})();';

let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ✓ ' + name); } else { fail++; console.log('  ✗ ' + name + (extra ? '  → ' + extra : '')); }
}
function eq(name, a, b) { ok(name, JSON.stringify(a) === JSON.stringify(b), `actual=${JSON.stringify(a)} expect=${JSON.stringify(b)}`); }

/** 造一个沙箱：可选预置 XLSX；注入行为可配置 */
function makeSandbox(opts) {
  const o = opts || {};
  const injected = [];
  const { window: lw } = parseHTML('<!DOCTYPE html><html><head></head><body></body></html>');

  const head = {
    children: [],
    appendChild(el) {
      this.children.push(el);
      injected.push(el.src);
      const behave = (o.behavior || (() => 'fail'))(el.src, injected.length - 1);
      setTimeout(() => {
        if (behave === 'ok') {
          sandbox.XLSX = REAL_XLSX;            // 模拟库加载后挂到全局
          if (el.onload) el.onload();
        } else if (behave === 'loaded-no-xlsx') {
          if (el.onload) el.onload();
        } else {
          if (el.onerror) el.onerror();
        }
      }, 0);
    },
    removeChild(el) { this.children = this.children.filter((x) => x !== el); },
  };

  // 自己造一个极简 document（linkedom 的 head 是只读 getter，不能覆盖）
  const document = {
    head, documentElement: head, readyState: 'complete',
    createElement(tag) {
      if (tag !== 'script') return { style: {}, appendChild() {}, remove() {}, setAttribute() {}, addEventListener() {} };
      const el = { tagName: 'SCRIPT', parentNode: head, onload: null, onerror: null, _src: '' };
      Object.defineProperty(el, 'src', { get() { return el._src; }, set(v) { el._src = v; } });
      return el;
    },
    getElementById() { return null; },
    querySelector() { return null; },
    querySelectorAll() { return []; },
    addEventListener() {},
  };
  const window = { location: { href: 'http://t/public/experiment.html?id=28', search: '?id=28', pathname: '/public/experiment.html', host: 't' }, DOMParser: lw.DOMParser, addEventListener() {} };

  const sandbox = {
    console, document, window, location: window.location,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    URL, setTimeout, clearTimeout, Date, Math, JSON, Number, String, Array, Object, RegExp, Map, Set,
    isNaN, parseInt, parseFloat,
    alert() {}, confirm() { return true; }, prompt() { return '1'; },
    fetch: async () => { throw new Error('no fetch'); },
    Blob: function () {}, DOMParser: lw.DOMParser,
    xlsxLogs: [],
  };
  sandbox.globalThis = sandbox;
  sandbox.window.localStorage = sandbox.localStorage;
  if (o.withXLSX) sandbox.XLSX = REAL_XLSX;
  const fns = vm.runInNewContext(body, sandbox, { filename: 'u.js' });
  return { fns, sandbox, injected };
}

(async () => {
  console.log('=== A. 已经加载好：不应再注入任何脚本 ===');
  {
    const { fns, injected } = makeSandbox({ withXLSX: true });
    const X = await fns.ensureXLSX();
    ok('直接返回已加载的 SheetJS', X && X.utils, '未返回');
    eq('没有发起注入', injected.length, 0);
  }

  console.log('\n=== B. 第一个来源成功 ===');
  {
    const { fns, injected } = makeSandbox({ behavior: (url, i) => (i === 0 ? 'ok' : 'fail') });
    const X = await fns.ensureXLSX();
    ok('返回 SheetJS', X && X.utils);
    eq('只试了第一个来源', injected.length, 1);
    ok('用的是 jsDelivr', injected[0].includes('cdn.jsdelivr.net'), injected[0]);
  }

  console.log('\n=== C. jsDelivr 失败 → 自动落到下一个来源（含 npmmirror）===');
  {
    // 让前两个失败，第三个成功，验证按顺序回退
    const { fns, injected } = makeSandbox({ behavior: (url, i) => (i === 2 ? 'ok' : 'fail') });
    const X = await fns.ensureXLSX();
    ok('最终仍然加载成功', X && X.utils);
    eq('试了 3 个来源', injected.length, 3);
    ok('第 2 个是 npmmirror（国内可访问）', injected[1].includes('npmmirror'), injected[1]);
  }

  console.log('\n=== D. 全部失败 → 明确返回 null（不抛错，让上层给友好提示）===');
  {
    const { fns, injected } = makeSandbox({ behavior: () => 'fail' });
    const X = await fns.ensureXLSX();
    eq('返回 null', X, null);
    eq('把所有候选都试过一遍', injected.length, fns.XLSX_CDN_CANDIDATES.length);
  }

  console.log('\n=== E. "脚本加载了但没暴露 XLSX" 也要继续尝试下一个 ===');
  {
    const calls = [];
    const { fns, injected } = makeSandbox({
      behavior: (url, i) => { calls.push(i); return i === 0 ? 'loaded-no-xlsx' : 'ok'; },
    });
    const X = await fns.ensureXLSX();
    ok('最终成功', X && X.utils);
    eq('第一个来源失败后继续', calls, [0, 1]);
    eq('注入次数', injected.length, 2);
  }

  console.log('\n=== F. 并发调用只加载一次 ===');
  {
    const { fns, injected } = makeSandbox({ behavior: (url, i) => (i === 0 ? 'ok' : 'fail') });
    const [a, b, c] = await Promise.all([fns.ensureXLSX(), fns.ensureXLSX(), fns.ensureXLSX()]);
    ok('三个调用都拿到同一库', a === b && b === c && !!a.utils, '不同实例');
    eq('只注入了一次', injected.length, 1);
  }

  console.log('\n=== G. 候选来源顺序合理（国内可达的排前面）===');
  {
    const { fns } = makeSandbox({ withXLSX: true });
    const list = fns.XLSX_CDN_CANDIDATES;
    ok('至少 4 个来源', list.length >= 4, list.length);
    ok('包含 npmmirror（国内镜像）', list.some((u) => u.includes('npmmirror')), list.join(','));
    ok('全部是 https', list.every((u) => u.startsWith('https://')), list.join(','));
    ok('版本固定 0.18.5（与 @require 一致）', list.every((u) => u.includes('0.18.5')), list.join(','));
  }

  console.log(`\n结果：${pass} 通过 / ${fail} 失败\n`);
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error('测试自身崩了：', e); process.exit(1); });
