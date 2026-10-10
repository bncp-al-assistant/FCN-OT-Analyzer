/* ===== Timesheet OT core (parse + compare) — pure functions, XLSX global required ===== */
var OTCore = (function () {
  var MONTHS = ['january','february','march','april','may','june','july','august','september','october','november','december'];
  var MON3 = MONTHS.map(function (m) { return m.slice(0, 3); });
  var WD_KO = ['일','월','화','수','목','금','토'];

  function norm(v) { return v == null ? '' : String(v).replace(/\s+/g, '').toUpperCase(); }
  function num(v) { if (v == null || v === '') return 0; var n = typeof v === 'number' ? v : parseFloat(v); return isFinite(n) ? n : 0; }
  function round(n, d) { var p = Math.pow(10, d == null ? 2 : d); return Math.round(n * p) / p; }
  function colLetter(c) { return XLSX.utils.encode_col(c); }

  function cell(ws, r, c) { return ws[XLSX.utils.encode_cell({ r: r, c: c })]; }
  function val(ws, r, c) { var x = cell(ws, r, c); return x ? x.v : null; }

  // Excel time (fraction of day) or "HH:MM" -> minutes since midnight; null if not a time
  function toMin(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') { if (v >= 0 && v < 1.0001) return Math.round(v * 1440) % 1440; if (v >= 1 && v < 2) return Math.round((v - 1) * 1440); return null; }
    var m = String(v).trim().match(/^(\d{1,2})[:.](\d{2})/);
    return m ? (+m[1]) * 60 + (+m[2]) : null;
  }
  function fmtMin(m) { if (m == null) return '—'; var h = Math.floor(m / 60), mm = m % 60; return (h < 10 ? '0' : '') + h + ':' + (mm < 10 ? '0' : '') + mm; }

  function detectYearMonth(fileName, sheetName, monthLabel) {
    var s = (fileName + ' ' + sheetName).toLowerCase();
    var month = null, year = null;
    for (var i = 0; i < 12; i++) { if (s.indexOf(MONTHS[i]) >= 0) { month = i + 1; break; } }
    if (!month) for (i = 0; i < 12; i++) { if (new RegExp('(^|[^a-z])' + MON3[i] + '([^a-z]|$)').test(s)) { month = i + 1; break; } }
    if (!month && monthLabel) { var mm = String(monthLabel).match(/(\d{1,2})\s*월/); if (mm) month = +mm[1]; }
    var y4 = s.match(/20\d{2}/); if (y4) year = +y4[0];
    if (!year) { var y2 = s.match(/[._\-\s](\d{2})(?:\D|$)/); if (y2) year = 2000 + (+y2[1]); }
    return { year: year, month: month };
  }

  function parseWorkbook(wb, fileName) {
    var ws = wb.Sheets[wb.SheetNames[0]];
    var sheetName = wb.SheetNames[0];
    var range = XLSX.utils.decode_range(ws['!ref']);
    var maxR = range.e.r, maxC = range.e.c;
    var warnings = [];

    // header: row containing TIME / Name
    var hdrRow = -1, timeCol = -1, nameCol = -1, enCol = -1, rateCol = -1;
    for (var r = 0; r <= Math.min(12, maxR) && hdrRow < 0; r++) {
      for (var c = 0; c <= maxC; c++) {
        var t = norm(val(ws, r, c));
        if (t === 'TIME') { hdrRow = r; timeCol = c; }
      }
    }
    if (hdrRow < 0) throw new Error('TIME 열을 찾지 못했습니다. 타임시트 양식이 맞는지 확인해 주세요.');
    for (c = 0; c <= maxC; c++) {
      var h = norm(val(ws, hdrRow, c));
      if (h === 'NAME') nameCol = c;
      if (h === 'E/N') enCol = c;
      if (h.indexOf('SALARY') === 0 || h.indexOf('HOURLY') >= 0) rateCol = c;
    }
    if (nameCol < 0) nameCol = timeCol - 2;
    if (rateCol < 0) rateCol = timeCol - 1;

    // day row: first row below header with 1,2,3 in consecutive columns after TIME
    var dayRow = -1, dayCol0 = timeCol + 1, nDays = 0;
    for (r = hdrRow; r <= hdrRow + 3; r++) {
      if (num(val(ws, r, timeCol + 1)) === 1 && num(val(ws, r, timeCol + 2)) === 2) {
        dayRow = r; for (c = timeCol + 1; c <= maxC; c++) { if (num(val(ws, r, c)) === nDays + 1) nDays++; else break; }
        break;
      }
    }
    if (dayRow < 0) throw new Error('날짜(1~31) 행을 찾지 못했습니다.');
    var monthLabel = val(ws, hdrRow, dayCol0);

    // summary columns by label
    var S = {};
    for (c = dayCol0 + nDays; c <= maxC; c++) {
      var h2 = norm(val(ws, hdrRow, c));
      if (!h2) continue;
      if (h2.indexOf('기본근무') >= 0) { S.bw = c; S.bo = c + 1; }
      else if (h2.indexOf('평일근무') >= 0) { S.wdTot = c; S.wd15 = c + 1; S.wd2 = c + 2; }
      else if (h2.indexOf('휴일근무') >= 0) { S.hdCnt = c; S.hdTot = c + 1; }
      else if (h2.indexOf('월급여') >= 0) S.pay = c;
      else if (h2.indexOf('기본급') >= 0) S.payBase = c;
      else if (h2.indexOf('평일OT') >= 0) S.payWd = c;
      else if (h2.indexOf('휴일OT') >= 0) S.payHd = c;
      else if (h2.indexOf('보너스') >= 0) S.bonus = c;
    }
    ['bw', 'wdTot', 'hdTot', 'pay'].forEach(function (k) { if (S[k] == null) warnings.push('요약 열 "' + k + '"를 찾지 못했습니다.'); });

    // meta: 근무일 / 휴일 in top rows
    var metaDays = null, metaHol = null;
    for (r = 0; r < hdrRow; r++) for (c = 0; c <= maxC; c++) {
      var m1 = norm(val(ws, r, c));
      if (m1 === '근무일') metaDays = num(val(ws, r, c + 1));
      if (m1 === '휴일') metaHol = num(val(ws, r, c + 1));
    }

    var ym = detectYearMonth(fileName || '', sheetName, monthLabel);

    // workers
    var workers = [];
    var firstHolFormula = null;
    for (r = dayRow + 1; r <= maxR; r++) {
      if (norm(val(ws, r, timeCol)) !== 'IN') continue;
      var rows = { IN: r };
      for (var k = 1; k <= 6 && r + k <= maxR; k++) {
        var lab = norm(val(ws, r + k, timeCol));
        if (lab === 'IN') break;
        if (lab === 'OUT' || lab === 'B/W' || lab === 'B/O' || lab === 'E/O') rows[lab] = r + k;
      }
      var name = val(ws, r, nameCol);
      if (!name) continue;
      var w = {
        no: val(ws, r, 0), id: enCol >= 0 ? String(val(ws, r, enCol) || '') : '', name: String(name).trim(),
        rate: num(val(ws, r, rateCol)), days: [], sheet: {}, formulas: {}
      };
      for (var d = 0; d < nDays; d++) {
        var cc = dayCol0 + d;
        var inv = val(ws, r, cc), outv = rows.OUT != null ? val(ws, rows.OUT, cc) : null;
        var inM = toMin(inv), outM = toMin(outv);
        var note = (inM == null && inv != null && String(inv).trim()) ? String(inv).trim() : '';
        w.days.push({
          d: d + 1, inM: inM, outM: outM, note: note,
          bw: rows['B/W'] != null ? num(val(ws, rows['B/W'], cc)) : 0,
          bo: rows['B/O'] != null ? num(val(ws, rows['B/O'], cc)) : 0,
          eo: rows['E/O'] != null ? num(val(ws, rows['E/O'], cc)) : 0
        });
      }
      Object.keys(S).forEach(function (key) {
        var x = cell(ws, r, S[key]);
        w.sheet[key] = x ? num(x.v) : 0;
        if (x && x.f) w.formulas[key] = x.f;
      });
      if (!firstHolFormula && w.formulas.hdTot) firstHolFormula = w.formulas.hdTot;
      workers.push(w);
    }
    if (!workers.length) throw new Error('작업자(IN 행)를 찾지 못했습니다.');

    // holidays: from 휴일 OT formula columns, else from B/O-blank pattern, else Fridays
    var holSet = {}, holSource = '';
    if (firstHolFormula) {
      var refs = firstHolFormula.replace(/\$/g, '').match(/[A-Z]{1,3}\d+/g) || [];
      refs.forEach(function (ref) { var cI = XLSX.utils.decode_cell(ref).c; var di = cI - dayCol0 + 1; if (di >= 1 && di <= nDays) holSet[di] = true; });
      if (Object.keys(holSet).length) holSource = '휴일 OT 수식에 지정된 날짜';
    }
    // pattern-based detection (for cross-check)
    var patSet = {};
    for (d = 1; d <= nDays; d++) {
      var act = 0, noBo = 0;
      workers.forEach(function (w) { var x = w.days[d - 1]; if (x.bw > 0 || x.eo > 0) { act++; if (!x.bo && x.eo > 0) noBo++; } });
      if (act && noBo / act >= 0.6) patSet[d] = true;
    }
    if (!holSource) {
      if (Object.keys(patSet).length) { holSet = patSet; holSource = 'B/O가 비어 있는 날(자동 감지)'; }
      else if (ym.year && ym.month) { for (d = 1; d <= nDays; d++) if (new Date(ym.year, ym.month - 1, d).getDay() === 5) holSet[d] = true; holSource = '금요일(기본값)'; }
    } else {
      var a = Object.keys(holSet).sort().join(','), b = Object.keys(patSet).sort().join(',');
      if (a !== b) warnings.push('휴일 OT 수식의 휴일(' + Object.keys(holSet).join(', ') + '일)과 B/O가 비어 있는 날(' + (Object.keys(patSet).join(', ') || '없음') + '일)이 다릅니다.');
    }
    var holidays = Object.keys(holSet).map(Number).sort(function (x, y) { return x - y; });
    if (metaHol != null && metaHol !== holidays.length) warnings.push('시트 상단 휴일 수(' + metaHol + ')와 감지된 휴일 수(' + holidays.length + ')가 다릅니다.');

    var month = {
      fileName: fileName, sheetName: sheetName, year: ym.year, month: ym.month, nDays: nDays,
      metaDays: metaDays, metaHol: metaHol, holidays: holidays, holSource: holSource, warnings: warnings, workers: workers
    };
    month.dayInfo = [];
    for (d = 1; d <= nDays; d++) {
      var wd = (ym.year && ym.month) ? new Date(ym.year, ym.month - 1, d).getDay() : null;
      month.dayInfo.push({ d: d, wd: wd, wdKo: wd == null ? '' : WD_KO[wd], hol: !!holSet[d] });
    }
    month.weekdays = nDays - holidays.length;
    workers.forEach(function (w) { analyzeWorker(w, month); });
    return month;
  }

  function patternKey(x) { return fmtMin(x.inM) + '–' + fmtMin(x.outM); }

  function distribution(list) {
    var map = {}; list.forEach(function (x) { var k = patternKey(x); map[k] = (map[k] || 0) + 1; });
    var arr = Object.keys(map).map(function (k) { return { key: k, n: map[k] }; });
    arr.sort(function (a, b) { return b.n - a.n; });
    var tot = list.length || 1;
    arr.forEach(function (e) { e.share = e.n / tot; });
    return arr;
  }

  function analyzeWorker(w, month) {
    var hol = {}; month.holidays.forEach(function (d) { hol[d] = true; });
    var A = { wdDays: 0, hdDays: 0, offDays: 0, boSum: 0, eoWd: 0, eoHd: 0, boHd: 0, bwSum: 0, statusNotes: [], gaps: [], wdBoDays: 0, bwMissing: [] };
    var wdList = [], hdList = [];
    w.days.forEach(function (x) {
      x.hol = !!hol[x.d];
      // 휴일에는 B/W 8h가 유급휴일로 일괄 기재되므로 실제 근무는 E/O·B/O로 판단
      var worked = x.hol ? (x.bo > 0 || x.eo > 0) : (x.bw > 0 || x.bo > 0 || x.eo > 0);
      x.ot = x.bo + x.eo;
      x.worked = worked;
      A.bwSum += x.bw;
      if (x.note) A.statusNotes.push({ d: x.d, note: x.note });
      if (!worked) { A.offDays++; return; }
      if (x.hol) { A.hdDays++; A.eoHd += x.eo; A.boHd += x.bo; if (x.inM != null && x.outM != null) hdList.push(x); }
      else { A.wdDays++; if (x.bw > 0 && x.bo > 0) A.wdBoDays++; if (x.bo > 0 && !(x.bw > 0)) A.bwMissing.push(x.d); A.boSum += x.bo; A.eoWd += x.eo; if (x.inM != null && x.outM != null) wdList.push(x); }
      if (x.inM != null && x.outM != null) {
        var span = (x.outM - x.inM + 1440) % 1440 / 60;
        var rec = x.hol ? x.bo + x.eo : x.bw + x.bo + x.eo;
        x.gap = round(span - rec, 2);
        if (x.gap < 0 || x.gap > 3) A.gaps.push({ d: x.d, span: span, rec: rec, gap: x.gap, inM: x.inM, outM: x.outM });
      }
    });
    A.wdDist = distribution(wdList);
    A.hdDist = distribution(hdList);
    A.wdOtGrid = A.boSum + A.eoWd;               // 평일 B/O+E/O 전체 (기본 포함분 포함)
    A.builtinExpected = A.wdBoDays * 2;          // B/W와 B/O가 모두 있는 평일 × 2h (기본급에 포함되는 B/O)
    A.builtinSheet = w.sheet.bo;
    A.wdOtPaid = w.sheet.wdTot;                  // 시트 평일 OT 합계 (기본 포함분 제외)
    A.hdOtPaid = w.sheet.hdTot;
    A.wdOtPerDay = A.wdDays ? A.wdOtPaid / A.wdDays : 0;
    A.hdOtPerDay = A.hdDays ? A.eoHd / A.hdDays : 0;
    // recomputation check
    A.checks = [];
    var expWd15 = A.boSum - A.builtinSheet, expWd2 = A.eoWd;
    if (Math.abs(expWd15 - num(w.sheet.wd15)) > 0.01) A.checks.push('평일 OT 1.5배: 일별 합산 ' + round(expWd15) + 'h ≠ 시트 ' + round(w.sheet.wd15) + 'h');
    if (Math.abs(expWd2 - num(w.sheet.wd2)) > 0.01) A.checks.push('평일 OT 2배: 일별 합산 ' + round(expWd2) + 'h ≠ 시트 ' + round(w.sheet.wd2) + 'h');
    if (Math.abs(A.eoHd + A.boHd - num(w.sheet.hdTot)) > 0.01) A.checks.push('휴일 OT: 일별 합산 ' + round(A.eoHd + A.boHd) + 'h ≠ 시트 ' + round(w.sheet.hdTot) + 'h');
    if (Math.abs(A.builtinExpected - A.builtinSheet) > 0.01) A.checks.push('기본급 포함 B/O: 평일 ' + A.wdBoDays + '일×2h = ' + A.builtinExpected + 'h 이어야 하나 시트는 ' + A.builtinSheet + 'h');
    A.manualAdj = (w.formulas.bo && /\)\s*\*\s*2\s*[-+]\s*\d+/.test(w.formulas.bo)) ? w.formulas.bo : '';
    if (A.bwMissing.length) A.checks.push(A.bwMissing.join(', ') + '일: B/O는 있는데 B/W(기본 8h)가 비어 있음');
    w.A = A;
  }

  function monthLabel(m) { return m.month ? (m.year ? m.year + '년 ' : '') + m.month + '월' : m.fileName; }
  function sum(arr, f) { return arr.reduce(function (s, x) { return s + f(x); }, 0); }

  function teamTotals(m) {
    var ws = m.workers;
    return {
      head: ws.length,
      wdOt: sum(ws, function (w) { return w.sheet.wdTot; }),
      wd15: sum(ws, function (w) { return w.sheet.wd15; }),
      wd2: sum(ws, function (w) { return w.sheet.wd2; }),
      hdOt: sum(ws, function (w) { return w.sheet.hdTot; }),
      payWd: sum(ws, function (w) { return w.sheet.payWd; }),
      payHd: sum(ws, function (w) { return w.sheet.payHd; }),
      payBase: sum(ws, function (w) { return w.sheet.payBase; }),
      pay: sum(ws, function (w) { return w.sheet.pay; }),
      wdPersonDays: sum(ws, function (w) { return w.A.wdDays; }),
      hdPersonDays: sum(ws, function (w) { return w.A.hdDays; })
    };
  }

  function keyOf(w) { return w.id ? 'id:' + w.id : 'nm:' + w.name.toUpperCase(); }

  function compare(m1, m2) {
    var prev = m1, cur = m2;
    if (m1.year && m2.year && m1.month && m2.month && (m1.year * 12 + m1.month) > (m2.year * 12 + m2.month)) { prev = m2; cur = m1; }
    var P = teamTotals(prev), C = teamTotals(cur);
    var map = {};
    prev.workers.forEach(function (w) { map[keyOf(w)] = { p: w }; });
    cur.workers.forEach(function (w) { var k = keyOf(w); (map[k] = map[k] || {}).c = w; });
    var order = [];
    cur.workers.forEach(function (w) { order.push(keyOf(w)); });
    prev.workers.forEach(function (w) { if (order.indexOf(keyOf(w)) < 0) order.push(keyOf(w)); });

    var holDiff = cur.holidays.length - prev.holidays.length;
    var wdDiff = cur.weekdays - prev.weekdays;

    var rows = order.map(function (k) {
      var p = map[k].p, c = map[k].c, w = c || p;
      var row = { key: k, name: w.name, id: w.id, p: p, c: c, tags: [], notes: [] };
      if (!p) { row.status = 'new'; row.tags.push({ t: '신규', s: 'info' }); return row; }
      if (!c) { row.status = 'gone'; row.tags.push({ t: '이번 달 명단 없음', s: 'warn' }); return row; }
      row.status = 'both';
      var pa = p.A, ca = c.A;
      row.dWdOt = c.sheet.wdTot - p.sheet.wdTot;
      row.dHdOt = c.sheet.hdTot - p.sheet.hdTot;
      row.dPayWd = c.sheet.payWd - p.sheet.payWd;
      row.dPayHd = c.sheet.payHd - p.sheet.payHd;
      row.dPay = c.sheet.pay - p.sheet.pay;
      // decomposition: calendar/attendance effect vs pattern effect
      row.wdExpected = pa.wdOtPerDay * ca.wdDays;
      row.wdDaysEffect = row.wdExpected - p.sheet.wdTot;
      row.wdPatternEffect = c.sheet.wdTot - row.wdExpected;
      var pHdPer = pa.hdDays ? p.sheet.hdTot / pa.hdDays : 0;
      row.hdExpected = pHdPer * ca.hdDays;
      row.hdDaysEffect = row.hdExpected - p.sheet.hdTot;
      row.hdPatternEffect = c.sheet.hdTot - row.hdExpected;

      // shift change detection
      row.wdShift = shiftChange(pa.wdDist, ca.wdDist);
      row.hdShift = shiftChange(pa.hdDist, ca.hdDist);
      if (row.wdShift.changed) row.tags.push({ t: '평일 시간대 변경', s: 'alert' });
      if (row.hdShift.changed) row.tags.push({ t: '휴일 시간대 변경', s: 'alert' });
      var perDayDelta = ca.wdOtPerDay - pa.wdOtPerDay;
      row.wdPerDayDelta = perDayDelta;
      if (!row.wdShift.changed && Math.abs(perDayDelta) >= 0.25) row.tags.push({ t: '평일 일평균 OT ' + (perDayDelta > 0 ? '+' : '') + round(perDayDelta, 2) + 'h', s: 'alert' });

      var expectedDays = cur.nDays;
      if (ca.offDays > 0 && (ca.wdDays + ca.hdDays) < expectedDays) row.tags.push({ t: '근무 ' + (ca.wdDays + ca.hdDays) + '/' + expectedDays + '일', s: 'warn' });
      ca.statusNotes.forEach(function (n) { row.notes.push(n.d + '일 "' + n.note + '" 기재'); });
      if (Math.abs(c.rate - p.rate) > 1e-6) row.tags.push({ t: '시급 ' + round(p.rate, 3) + '→' + round(c.rate, 3), s: 'warn' });
      if (ca.checks.length) row.tags.push({ t: '시트 계산 확인 필요', s: 'warn' });
      if (ca.gaps.length) row.tags.push({ t: '출퇴근·시간 불일치 ' + ca.gaps.length + '일', s: 'warn' });
      return row;
    });

    // calendar effect summary for continuing workers
    return {
      prev: prev, cur: cur, P: P, C: C, rows: rows,
      holDiff: holDiff, wdDiff: wdDiff,
      label: { p: monthLabel(prev), c: monthLabel(cur) }
    };
  }

  function shiftChange(pd, cd) {
    var res = { changed: false, pMain: pd[0] || null, cMain: cd[0] || null, moves: [] };
    if (!pd.length || !cd.length) return res;
    var keys = {}; pd.forEach(function (e) { keys[e.key] = 1; }); cd.forEach(function (e) { keys[e.key] = 1; });
    Object.keys(keys).forEach(function (k) {
      var a = pd.filter(function (e) { return e.key === k; })[0], b = cd.filter(function (e) { return e.key === k; })[0];
      var sa = a ? a.share : 0, sb = b ? b.share : 0;
      res.moves.push({ key: k, pn: a ? a.n : 0, cn: b ? b.n : 0, ps: sa, cs: sb, d: sb - sa });
    });
    res.moves.sort(function (x, y) { return Math.abs(y.d) - Math.abs(x.d); });
    if (res.pMain.key !== res.cMain.key) res.changed = true;
    if (res.moves.length && Math.abs(res.moves[0].d) >= 0.15) res.changed = true;
    return res;
  }

  return { parseWorkbook: parseWorkbook, compare: compare, fmtMin: fmtMin, round: round, monthLabel: monthLabel, WD_KO: WD_KO };
})();
if (typeof module !== 'undefined') module.exports = OTCore;
