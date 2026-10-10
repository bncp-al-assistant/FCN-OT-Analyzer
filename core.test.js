// 예시 파일(samples/)로 파싱·비교 결과를 검증합니다.  npm test
global.XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const assert = require('assert');
const OT = require('../js/core.js');
const load = f => OT.parseWorkbook(XLSX.read(fs.readFileSync(path.join(__dirname, '..', 'samples', f)), { cellFormula: true }), f);

const jul = load('Sample_TCN_Zone_July.2026.xlsx');
const aug = load('Sample_TCN_Zone_August.2026.xlsx');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('월 인식', () => { assert.strictEqual(jul.month, 7); assert.strictEqual(aug.month, 8); assert.strictEqual(jul.year, 2026); });
t('휴일 인식(휴일 OT 수식)', () => {
  assert.deepStrictEqual(jul.holidays, [3, 10, 17, 24, 31]);
  assert.deepStrictEqual(aug.holidays, [7, 14, 21, 28]);
  assert.strictEqual(jul.weekdays, 26); assert.strictEqual(aug.weekdays, 27);
});
t('작업자 9명', () => { assert.strictEqual(jul.workers.length, 9); assert.strictEqual(aug.workers.length, 9); });

// 파일 순서를 거꾸로 넣어도 전달/이번달이 맞게 정렬되는지
const r = OT.compare(aug, jul);
t('전달·이번달 자동 정렬', () => { assert.strictEqual(r.prev.month, 7); assert.strictEqual(r.cur.month, 8); });
t('팀 합계 = 시트 요약값', () => {
  assert.strictEqual(r.P.wdOt, 761); assert.strictEqual(r.C.wdOt, 753);
  assert.strictEqual(r.P.hdOt, 507); assert.strictEqual(r.C.hdOt, 358);
  assert.ok(Math.abs(r.P.pay - 6010.945) < 0.01); assert.ok(Math.abs(r.C.pay - 5204.8776) < 0.01);
});
t('일별 재계산이 시트 요약과 일치(7월)', () => { jul.workers.forEach(w => assert.deepStrictEqual(w.A.checks, [], w.name)); });

const by = name => r.rows.find(x => x.name === name);
t('휴일 1일 감소 → 휴일 OT 감소는 전부 일수 효과', () => {
  const w = by('WORKER 01');
  assert.strictEqual(w.dHdOt, -14); assert.strictEqual(w.hdDaysEffect, -14); assert.strictEqual(w.hdPatternEffect, 0);
  assert.strictEqual(w.wdDaysEffect, 5); assert.strictEqual(w.wdPatternEffect, 0);
  assert.strictEqual(w.tags.length, 0);
});
t('평일 시간대 변경 감지', () => {
  const w = by('WORKER 03');
  assert.ok(w.wdShift.changed);
  assert.ok(w.tags.some(x => x.t === '평일 시간대 변경'));
  assert.ok(w.notes.some(x => /Departure/.test(x)));
});
t('평일 1일 평균 OT 변화 감지', () => { assert.ok(by('WORKER 06').tags.some(x => /평일 일평균 OT/.test(x.t))); });
t('B/W 누락 감지', () => { assert.deepStrictEqual(by('WORKER 06').c.A.bwMissing, [31]); });
t('기본급 포함 B/O 수동 조정 인식', () => {
  const w = by('WORKER 02').c.A;
  assert.ok(w.manualAdj); assert.strictEqual(w.builtinExpected, w.builtinSheet);
});
console.log(`\n${n} tests passed`);
