// 사용법: node scripts/analyze.js <전달.xlsx> <이번달.xlsx>
// 브라우저 화면과 같은 로직(js/core.js)으로 결과를 터미널에 요약합니다.
global.XLSX = require('xlsx');
const fs = require('fs');
const OT = require('../js/core.js');
const files = process.argv.slice(2);
if (files.length !== 2) { console.error('사용법: node scripts/analyze.js <전달.xlsx> <이번달.xlsx>'); process.exit(1); }
const months = files.map(f => OT.parseWorkbook(XLSX.read(fs.readFileSync(f), { cellFormula: true }), require('path').basename(f)));
const r = OT.compare(months[0], months[1]);
const h = n => (Math.round(n * 10) / 10).toString();
console.log(`${r.label.p} → ${r.label.c}`);
console.log(`  평일 ${r.prev.weekdays} → ${r.cur.weekdays}일, 휴일 ${r.prev.holidays.length} → ${r.cur.holidays.length}일`);
console.log(`  평일 OT ${h(r.P.wdOt)}h → ${h(r.C.wdOt)}h ($${r.P.payWd.toFixed(2)} → $${r.C.payWd.toFixed(2)})`);
console.log(`  휴일 OT ${h(r.P.hdOt)}h → ${h(r.C.hdOt)}h ($${r.P.payHd.toFixed(2)} → $${r.C.payHd.toFixed(2)})`);
console.log(`  월급여 $${r.P.pay.toFixed(2)} → $${r.C.pay.toFixed(2)}`);
for (const x of r.rows) {
  console.log(`\n- ${x.name} [${x.tags.map(t => t.t).join(', ') || '특이사항 없음'}]`);
  if (x.status !== 'both') continue;
  console.log(`  평일 OT ${x.p.sheet.wdTot} → ${x.c.sheet.wdTot}h (일수 효과 ${h(x.wdDaysEffect)}, 패턴 효과 ${h(x.wdPatternEffect)})`);
  console.log(`  휴일 OT ${x.p.sheet.hdTot} → ${x.c.sheet.hdTot}h (일수 효과 ${h(x.hdDaysEffect)}, 패턴 효과 ${h(x.hdPatternEffect)})`);
  console.log(`  평일 주 시간대 ${x.wdShift.pMain && x.wdShift.pMain.key} → ${x.wdShift.cMain && x.wdShift.cMain.key}`);
  x.notes.concat(x.c.A.checks).forEach(n => console.log('  · ' + n));
}
