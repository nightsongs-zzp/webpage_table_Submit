/**
 * 隐私自检：扫描仓库里可能泄露个人信息 / 敏感凭据的内容。
 * 只扫「git 已跟踪」的文件（未跟踪的本地笔记本来就不会提交）。
 *
 * 用法： node tests/scan-personal.js [--history]
 *   --history  额外扫描提交历史里的作者信息（需要 force push 才能改写）
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const WITH_HISTORY = process.argv.includes('--history');

/**
 * 规则用「可读描述 + 正则」表达，**不在源码里写死任何真实姓名/账号**，
 * 否则扫描器自己就成了泄露源（这正是本文件要避免的事）。
 */
const RULES = [
  { label: '邮箱地址', re: /[A-Za-z0-9._%+-]+@(?:qq|163|126|gmail|outlook|foxmail|sina)\.com/i },
  { label: '会话 cookie', re: /connect\.sid\s*[=:]/ },
  { label: '手机号', re: /(?<![\d.])1[3-9]\d{9}(?![\d.])/ },
  { label: '疑似学号（独立出现的 8~12 位数字）', re: /(?<![\w.])\d{8,12}(?![\w.])/ },
  {
    label: '服务器地址',
    // 只认"像一个端点"的地址，且四段都必须是合法八位组（0~255）——
    // 否则 SVG 的 viewBox / path 数据（如 "3.235.577.075"）会被误判成 IP。
    test: (line) => {
      const m = line.match(/(?<![\w.$+\d])(\d{1,3}(?:\.\d{1,3}){3})(:\d{1,5})?(?![\w.])/);
      if (!m) return null;
      const octetsOk = m[1].split('.').every((o) => Number(o) <= 255);
      if (!octetsOk) return null;
      const hasPort = !!m[2];
      const inContext = /(https?:\/\/|@match|@connect|\bhost\b|\bOrigin\b|\bReferer\b)/i.test(line);
      return (hasPort || inContext) ? m[0] : null;
    },
  },
];
// 允许保留的"假"地址（示例/测试用），避免误报
const ALLOW_IP = /^(?:0\.0\.0\.0|127\.0\.0\.1|1\.2\.3\.4|x)$/;

const SKIP_DIRS = new Set(['node_modules', '.git', 'data', '测试']);
const TEXT_EXT = new Set(['.js', '.json', '.md', '.txt', '.html', '.css', '.yml', '.yaml', '.sh', '.bat', '.ps1']);
const SELF = path.join('tests', 'scan-personal.js');

function trackedFiles() {
  const out = execSync('git ls-files -z', { cwd: ROOT, encoding: 'utf8' })
    .split('\0').filter(Boolean);
  return out
    .map((p) => p.replace(/\//g, path.sep))
    .filter((p) => !SKIP_DIRS.has(p.split(path.sep)[0]))
    .filter((p) => TEXT_EXT.has(path.extname(p).toLowerCase()))
    .filter((p) => p !== SELF);
}

let hits = 0;
console.log('=== 1. 已跟踪文本文件扫描 ===');
for (const rel of trackedFiles()) {
  const abs = path.join(ROOT, rel);
  let text;
  try { text = fs.readFileSync(abs, 'utf8'); } catch (e) { continue; }
  text.split('\n').forEach((line, i) => {
    for (const r of RULES) {
      // 规则可以是正则(re)或自定义判定函数(test)
      let hit = null;
      if (typeof r.test === 'function') hit = r.test(line);
      else { const m = line.match(r.re); hit = m ? m[0] : null; }
      if (!hit) continue;
      if (r.label === '服务器地址' && ALLOW_IP.test(hit)) continue;
      if (r.label === '疑似学号（独立出现的 8~12 位数字）') {
        // 只对"上下文像学号"的行报警，避免 CSS z-index、时间戳之类的噪音
        if (!/(学号|姓名|学生|number|student)/i.test(line)) continue;
      }
      hits++;
      console.log(`  ⚠ [${r.label}] ${rel.replace(/\\/g, '/')}:${i + 1}`);
      console.log(`      ${line.trim().slice(0, 120)}`);
    }
  });
}
if (!hits) console.log('  ✓ 无命中');

console.log('\n=== 2. git 元数据（不在文件里，但会随提交泄露）===');
for (const [label, cmd] of [
  ['user.name', ['config', 'user.name']],
  ['user.email', ['config', 'user.email']],
  ['remote.origin.url', ['config', 'remote.origin.url']],
]) {
  let v = '';
  try { v = execSync(['git', ...cmd].join(' '), { cwd: ROOT, encoding: 'utf8' }).trim(); } catch (e) { v = '(未设置)'; }
  console.log(`  ${label} = ${v}`);
}

if (WITH_HISTORY) {
  console.log('\n=== 3. 提交历史作者 ===');
  try {
    const log = execSync('git log --all "--format=%an <%ae>"', { cwd: ROOT, encoding: 'utf8' })
      .trim().split('\n').filter(Boolean);
    [...new Set(log)].forEach((l) => console.log(`  ${l}`));
    console.log(`  （共 ${log.length} 条提交）`);
    console.log('  ⚠ 改写历史会改变所有 commit hash 并需要 force push — 执行前先确认远端状态。');
  } catch (e) { console.log('  （无提交）'); }
} else {
  console.log('\n（提交历史里的作者信息需要 --history 才扫，避免误触）');
}

console.log(`\n合计命中：${hits} 处`);
process.exit(hits ? 1 : 0);
