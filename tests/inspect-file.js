'use strict';
const fs = require('fs');
const p = process.argv[2];
const raw = fs.readFileSync(p, 'utf8');

// ---- .xlsx：走 SheetJS ----
if (/\.xlsx$/i.test(p) || raw.charCodeAt(0) === 0x50 /* 'P' of PK */) {
  const XLSX = require('xlsx');
  const wb = XLSX.read(fs.readFileSync(p), { type: 'buffer', cellDates: false });
  console.log('文件：' + p);
  console.log('工作表：' + JSON.stringify(wb.SheetNames));
  wb.SheetNames.forEach((sn) => {
    const ws = wb.Sheets[sn];
    const ref = ws['!ref'] || '(空)';
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true, defval: '' });
    console.log(`\n--- [${sn}] ref=${ref} 行数=${aoa.length} ---`);
    console.log('  第1行(表名)  : ' + JSON.stringify(aoa[0]));
    if (aoa[1]) console.log('  第2行(列名)  : ' + JSON.stringify(aoa[1]));
    if (aoa[2]) console.log('  第3行(数据1) : ' + JSON.stringify(aoa[2]));
    const body = aoa.slice(2).filter((r) => Array.isArray(r));
    const filled = body.reduce((a, r) => a + r.filter((c) => String(c).trim() !== '').length, 0);
    console.log(`  数据行=${body.length} 已填格=${filled}`);
  });
  process.exit(0);
}

let j;
try {
  j = JSON.parse(raw);
} catch (e) {
  console.log('\n!!! JSON.parse 失败:', e.message);
  const m = /position (\d+)/.exec(e.message);
  if (m) {
    const at = parseInt(m[1], 10);
    console.log('上下文:', JSON.stringify(raw.slice(Math.max(0, at - 160), at + 160)));
  }
  process.exit(1);
}

console.log('\n=== 顶层 ===');
console.log('file =', JSON.stringify(j.file));
console.log('tables =', j.tables.length, ' _baseline =', j._baseline.length);
console.log('_warnings =', JSON.stringify(j._warnings));

console.log('\n=== summary ===');
j.summary.forEach((s) => console.log(`  id=${s.id} rows=${s.rows} cols=${s.cols} filled=${s.filled} name=${JSON.stringify(s.name)}`));

console.log('\n=== tables 结构 ===');
j.tables.forEach((t) => {
  const r0 = t.rows && t.rows[0];
  console.log(`  id=${t.id} headers=${t.headers.length} rows=${t.rows.length} rowsIsArray=${Array.isArray(t.rows)}`);
  console.log(`     rows[0] = ${JSON.stringify(r0)}`);
  console.log(`     headers[0] = ${JSON.stringify(t.headers[0])}`);
  console.log(`     headers 末列 = ${JSON.stringify(t.headers[t.headers.length - 1])}`);
});

console.log('\n=== 表头自检：首字符/末字符不平衡（服务器脏数据）===');
let dirty = 0;
j.tables.forEach((t) => {
  t.headers.forEach((h, i) => {
    const s = String(h);
    const opens = (s.match(/“/g) || []).length, closes = (s.match(/”/g) || []).length;
    const asciiQ = (s.match(/"/g) || []).length;
    const startsQ = s.startsWith('"'), endsQ = s.endsWith('"');
    if (opens !== closes || asciiQ > 0 || startsQ || endsQ) {
      dirty++;
      console.log(`  ⚠ 表${t.id} 列${i}: opens=${opens} closes=${closes} ascii_quote=${asciiQ} 首字符引号=${startsQ} 末字符引号=${endsQ}`);
      console.log(`     值 = ${JSON.stringify(s)}`);
    }
  });
});
console.log(`\n可疑表头数：${dirty}`);
