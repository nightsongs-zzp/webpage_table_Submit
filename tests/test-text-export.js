/**
 * 网页正文导出（Markdown / HTML / 纯文本）自测。
 * 用法： node tests/test-text-export.js
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { parseHTML } = require('linkedom');

const SRC = path.join(__dirname, '..', 'webpage-table-sync.user.js');
const src = fs.readFileSync(SRC, 'utf8');
const cut = src.indexOf("if (document.readyState === 'loading')");
const body = src.slice(0, cut) +
  '\nreturn { contentToMarkdown, htmlToMarkdown, markdownToPlain, markdownToReadableHtml, buildTextDoc, S };\n})();';

const { document, window } = parseHTML('<!DOCTYPE html><html><head></head><body></body></html>');
window.location = { href: 'http://test-host/public/experiment.html?id=22', search: '?id=22', pathname: '/public/experiment.html', host: 'test-host' };

const sandbox = {
  console, document, window, XLSX: null,
  location: window.location,
  localStorage: { _d: {}, getItem(k) { return this._d[k] === undefined ? null : this._d[k]; }, setItem(k, v) { this._d[k] = String(v); }, removeItem(k) { delete this._d[k]; } },
  URL: { createObjectURL: () => 'blob:x', revokeObjectURL() {} }, Blob: function () {},
  alert: () => {}, confirm: () => true, prompt: () => '1',
  setTimeout, clearTimeout, Date, Math, JSON, Number, String, Array, Object, RegExp, Map, Set, isNaN, parseInt, parseFloat,
  DOMParser: window.DOMParser,
  fetch: async () => { throw new Error('测试环境不应发起 fetch'); },
};
sandbox.globalThis = sandbox;
sandbox.window.localStorage = sandbox.localStorage;
const fns = vm.runInNewContext(body, sandbox, { filename: 'userscript.js' });

let pass = 0, fail = 0;
function eq(name, actual, expect) {
  const a = JSON.stringify(actual), e = JSON.stringify(expect);
  if (a === e) { pass++; console.log('  ✓ ' + name); } else { fail++; console.log(`  ✗ ${name}\n      actual=${a}\n      expect=${e}`); }
}
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  ✓ ' + name); } else { fail++; console.log('  ✗ ' + name + (extra ? '  → ' + extra : '')); }
}
function has(name, hay, needle) {
  ok(name, String(hay).includes(needle), '未包含 ' + JSON.stringify(needle) + '；实际片段=' + JSON.stringify(String(hay).slice(0, 200)));
}

console.log('=== A. HTML → Markdown ===');
{
  const html = [
    '<h3>实验目的</h3>',
    '<p>理解<strong>超导转变</strong>与<em>临界温度</em>的关系。</p>',
    '<ul><li>测量 R-T 曲线</li><li>求 <code>T_C</code><ul><li>用 $R_n/2$ 法</li></ul></li></ul>',
    '<ol><li>第一步</li><li>第二步</li></ol>',
    '<p>公式：$V_{Pt}=I R_{Pt}$　行间：</p>',
    '<p>$$R = \\frac{U}{I}$$</p>',
    '<p><a href="/uploads/a.pdf">讲义</a> <img src="/uploads/b.png" alt="装置图"></p>',
    '<blockquote>注意液氮安全</blockquote>',
    '<table><tr><th>器材</th><th>数量</th></tr><tr><td>杜瓦瓶</td><td>1</td></tr></table>',
    '<p>结束。</p>',
  ].join('');
  const md = fns.contentToMarkdown(html);
  has('标题转成 ## 级', md, '## 实验目的');
  has('加粗', md, '**超导转变**');
  has('斜体', md, '*临界温度*');
  has('无序列表', md, '- 测量 R-T 曲线');
  has('嵌套列表缩进', md, '  - 用 $R_n/2$ 法');
  has('有序列表', md, '1. 第一步');
  has('行内公式原样保留', md, '$V_{Pt}=I R_{Pt}$');
  has('行间公式原样保留', md, '$$R = \\frac{U}{I}$$');
  has('链接', md, '[讲义](/uploads/a.pdf)');
  has('图片', md, '![装置图](/uploads/b.png)');
  has('引用', md, '> 注意液氮安全');
  has('表格转 Markdown', md, '| 器材 | 数量 |');
  has('表格分隔行', md, '| --- | --- |');
  has('表格数据行', md, '| 杜瓦瓶 | 1 |');
  ok('没有残留 HTML 标签', !/<\/?(p|ul|li|h3|strong|table|tr|td)\b/i.test(md), md);
  ok('连续空行已压缩', !/\n{3,}/.test(md), JSON.stringify(md));
}

console.log('\n=== B. 本来就是 Markdown 的输入：原样保留 ===');
{
  const raw = '### 原理\n\n用 |A₁|=|A₂| 检验，公式 $I_0/T^2$。\n\n- 列表项';
  eq('不经过 HTML 转换', fns.contentToMarkdown(raw), raw);
  eq('空输入 → 空串', fns.contentToMarkdown(''), '');
}

console.log('\n=== C. Markdown → 纯文本 ===');
{
  const md = [
    '# 大标题',
    '## 小节',
    '这是**加粗**和*斜体*，还有 `代码`。',
    '公式 $T_C$ 与行间 $$R=\\frac{U}{I}$$ 也要能读。',
    '- 列表项 A',
    '1. 有序项',
    '| 器材 | 数量 |',
    '| --- | --- |',
    '| 杜瓦瓶 | 1 |',
    '![图](/a.png)',
    '[讲义](/a.pdf)',
  ].join('\n');
  const plain = fns.markdownToPlain(md);
  ok('没有 # 标题标记', !/^#/m.test(plain), plain);
  ok('没有 ** 加粗标记', !plain.includes('**'), plain);
  ok('没有反引号', !plain.includes('`'), plain);
  ok('没有图片语法', !plain.includes('!['), plain);
  has('行内公式去掉 $ 只留内容', plain, 'T_C');
  has('链接保留文字与地址', plain, '讲义 (/a.pdf)');
  has('表格分隔行被去掉', plain.split('\n').some((l) => /^\s*\|?[\s|:-]+\|?\s*$/.test(l) && l.includes('---')), false);
  has('表格内容仍可读', plain, '杜瓦瓶');
  has('列表项保留', plain, '列表项 A');
}

console.log('\n=== D. Markdown → 可读 HTML ===');
{
  const md = '# 标题\n\n段落 **粗** 和 $x^2$。\n\n$$\\int_0^1 x dx$$\n\n- a\n- b\n\n| h1 | h2 |\n| --- | --- |\n| 1 | 2 |\n\n> 引用\n\n---\n';
  const html = fns.markdownToReadableHtml(md, '测试文档');
  has('是完整 HTML', html, '<!DOCTYPE html>');
  has('标题正确', html, '<title>测试文档</title>');
  has('h1', html, '<h1>标题</h1>');
  has('加粗', html, '<strong>粗</strong>');
  has('列表', html, '<li>a</li>');
  has('表格', html, '<table>');
  has('分隔行没变成数据行', html.includes('| --- |'), false);
  has('引用', html, '<blockquote>引用</blockquote>');
  has('水平线', html, '<hr>');
  has('公式回填（无 katex 时退化为原文）', html, '$x^2$');
  ok('公式占位符已全部回填', !/[\uE000\uE001]/.test(html), '仍有私有字符残留');
  ok('HTML 转义生效（不会注入标签）', !html.includes('<script>'), '');
  const evil = fns.markdownToReadableHtml('<script>alert(1)</script>', 'x');
  has('尖括号被转义', evil, '&lt;script&gt;');
}

console.log('\n=== E. 组装整篇文档（含 15 个章节）===');
{
  fns.S.expId = '22';
  fns.S.expData = {
    name: '高温超导转变温度测量',
    category_name: '近代物理实验',
    difficulty: '★★★★',
    period: '4学时',
    virtualLabUrl: 'http://lab.example/virtual',
    core_knowledge: '<p>超导、迈斯纳效应</p>',
    pre_knowledge: '电磁学基础',
    target_content: '<h3>目的</h3><p>测量 $T_C$</p>',
    principle_content: '<p>四引线法测电阻：$R=U/I$</p>',
    equipment_table: '<table><tr><th>器材</th></tr><tr><td>杜瓦瓶</td></tr></table>',
    safety_content: '<ul><li>液氮防冻伤</li></ul>',
    operation_flow: '<ol><li>降温</li><li>记录</li></ol>',
    operation_tips: '<p>保持恒温</p>',
    error_analysis: '<p>接触电阻</p>',
    result_require: '<p>画出 R-T 曲线</p>',
    after_question: '<p>为什么用四引线？</p>',
    expand_content: '<p>BSCCO 应用</p>',
    test_bank: '<p>自测题若干</p>',
    data_table: '### 表1\n\n| 序号 | 电压 |\n| --- | --- |\n| 1 | |',
    data_calc: '<p>数据处理说明</p>',
    dataTableDefinition: '[{"id":1}]',
  };
  const doc = fns.buildTextDoc();
  ok('生成了文档', !!doc, 'buildTextDoc 返回空');
  const md = doc.markdown;
  has('一级标题含实验名', md, '# 高温超导转变温度测量（实验 id=22）');
  has('来源信息', md, '来源：test-host');
  has('说明不含表格数据', md, '只含**表格之外**的文字内容');
  has('目录', md, '## 目录');
  has('基本信息：分类', md, '所属分类');
  has('基本信息：难度', md, '★★★★');
  has('基本信息：虚拟实验', md, 'http://lab.example/virtual');
  ['核心知识点', '前置知识要求', '实验目的', '实验原理', '实验器材', '安全注意事项',
    '操作步骤 / 流程', '操作要点与注意事项', '误差分析', '结果要求', '课后思考题', '拓展应用', '自测题库']
    .forEach((t) => has('章节：' + t, md, '## ' + t));
  has('表格类内容单独成节并注明', md, '数据记录与数据处理（表格，见网页）');
  has('表格文字说明保留', md, '数据处理说明');
  has('内嵌表格转成 Markdown', md, '| 器材 |');
  has('LaTeX 未被破坏', md, '$R=U/I$');
  ok('没有把 dataTableDefinition 原始 JSON 塞进去', !md.includes('"id":1'), md.slice(0, 400));
  eq('章节计数合理（>=13）', doc.sectionCount >= 13, true);
}

console.log('\n=== F. 空字段处理 ===');
{
  fns.S.expData = { name: '空实验', core_knowledge: '', target_content: '   ', principle_content: '<p></p>' };
  const doc = fns.buildTextDoc();
  ok('仍能生成', !!doc);
  has('列出空章节提示', doc.markdown, '以下章节本次为空');
  has('空章节点名', doc.markdown, '实验目的');
  ok('空章节没有正文标题', !/^## 实验目的$/m.test(doc.markdown), doc.markdown);
}

console.log('\n=== G. 无数据时返回 null（由上层报错）===');
{
  fns.S.expData = null;
  eq('S.expData 为空 → null', fns.buildTextDoc(), null);
}

console.log(`\n结果：${pass} 通过 / ${fail} 失败\n`);
process.exit(fail ? 1 : 0);
