(function () {
  var R = OTCore.round, fm = OTCore.fmtMin;
  var state = { prev: null, cur: null, isSample: false, effTab: 'wd' };
  var $ = function (id) { return document.getElementById(id); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function h(n) { return (Math.round(n * 10) / 10).toLocaleString('en-US', { maximumFractionDigits: 1 }); }
  function usd(n) { return '$' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function sgn(n, f) { var v = Math.abs(n) < 0.005 ? 0 : n; return (v > 0 ? '+' : v < 0 ? '−' : '±') + f(Math.abs(v)); }
  function dcls(n) { return Math.abs(n) < 0.005 ? 'd-0' : n > 0 ? 'd-up' : 'd-down'; }
  function pct(a, b) { return a ? ((b - a) / a * 100) : 0; }

  /* ---------- file handling ---------- */
  function readFile(file) {
    return new Promise(function (res, rej) {
      var fr = new FileReader();
      fr.onload = function () {
        try { res(OTCore.parseWorkbook(XLSX.read(new Uint8Array(fr.result), { type: 'array', cellFormula: true }), file.name)); }
        catch (e) { rej(new Error(file.name + ': ' + e.message)); }
      };
      fr.onerror = function () { rej(new Error(file.name + ': 파일을 읽지 못했습니다.')); };
      fr.readAsArrayBuffer(file);
    });
  }
  function ymKey(m) { return m.year && m.month ? m.year * 12 + m.month : null; }
  function showErr(msg) { var e = $('err'); e.textContent = msg; e.hidden = !msg; }

  function setMonths(a, b, isSample) {
    if (a && b && ymKey(a) && ymKey(b) && ymKey(a) > ymKey(b)) { var t = a; a = b; b = t; }
    if (a) state.prev = a; if (b) state.cur = b;
    state.isSample = !!isSample;
    render();
  }
  function handleFiles(list, slot) {
    showErr('');
    var files = Array.prototype.slice.call(list || []).filter(function (f) { return /\.xls[xm]?$/i.test(f.name); });
    if (!files.length) { showErr('엑셀 파일(.xlsx)을 선택해 주세요.'); return; }
    Promise.all(files.slice(0, 2).map(readFile)).then(function (ms) {
      if (slot === 'prev') { state.prev = ms[0]; state.isSample = false; render(); return; }
      if (slot === 'cur') { state.cur = ms[0]; state.isSample = false; render(); return; }
      if (ms.length === 1) {
        // single file: place by month relative to whichever is loaded
        var m = ms[0];
        if (state.cur && ymKey(m) && ymKey(state.cur) && ymKey(m) < ymKey(state.cur)) state.prev = m; else { state.prev = state.cur; state.cur = m; }
        state.isSample = false; render(); return;
      }
      setMonths(ms[0], ms[1], false);
    }).catch(function (e) { showErr(e.message); });
  }
  $('fileBoth').addEventListener('change', function (e) { handleFiles(e.target.files); e.target.value = ''; });
  $('filePrev').addEventListener('change', function (e) { handleFiles(e.target.files, 'prev'); e.target.value = ''; });
  $('fileCur').addEventListener('change', function (e) { handleFiles(e.target.files, 'cur'); e.target.value = ''; });
  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (e) { handleFiles(e.dataTransfer.files); });

  /* ---------- tooltip ---------- */
  var tip = $('tip');
  document.addEventListener('mouseover', function (e) {
    var t = e.target.closest && e.target.closest('[data-tip]');
    if (!t) { tip.hidden = true; return; }
    tip.innerHTML = t.getAttribute('data-tip'); tip.hidden = false;
  });
  document.addEventListener('mousemove', function (e) {
    if (tip.hidden) return;
    var x = e.clientX + 14, y = e.clientY + 14, w = tip.offsetWidth, hh = tip.offsetHeight;
    if (x + w > window.innerWidth - 8) x = e.clientX - w - 10;
    if (y + hh > window.innerHeight - 8) y = e.clientY - hh - 10;
    tip.style.left = x + 'px'; tip.style.top = y + 'px';
  });

  /* ---------- render ---------- */
  function render() {
    var p = state.prev, c = state.cur;
    $('moPrev').textContent = p ? OTCore.monthLabel(p) : '—';
    $('fnPrev').textContent = p ? p.fileName : '파일 선택';
    $('moCur').textContent = c ? OTCore.monthLabel(c) : '—';
    $('fnCur').textContent = c ? c.fileName : '파일 선택';
    $('sampleFlag').hidden = !state.isSample;
    if (!p || !c) { $('out').innerHTML = '<div class="panel method"><p>두 달 치 타임시트가 모두 있어야 비교할 수 있습니다. 나머지 한 달 파일을 올려 주세요.</p></div>'; return; }
    var r = OTCore.compare(p, c);
    state.r = r;
    var zone = (c.sheetName || '').replace(/_.*$/, '');
    if (zone) $('zoneName').textContent = zone + ' · Time Sheet';
    $('out').innerHTML = [secCalendar(r), secTotals(r), secFindings(r), secEffects(r), secShift(r), secWorkers(r), secMethod()].join('');
    bindTabs();
  }

  function miniCal(m) {
    var out = ['일', '월', '화', '수', '목', '금', '토'].map(function (d) { return '<span class="h">' + d + '</span>'; });
    var first = m.dayInfo[0].wd;
    if (first == null) first = 0;
    for (var i = 0; i < first; i++) out.push('<span></span>');
    m.dayInfo.forEach(function (d) { out.push('<span class="' + (d.hol ? 'hd' : '') + '"' + (d.hol ? ' data-tip="휴일 ' + d.d + '일(' + d.wdKo + ')"' : '') + '>' + d.d + '</span>'); });
    return '<div class="mini-cal" aria-label="달력">' + out.join('') + '</div>';
  }
  function calMonth(m, other, isCur) {
    var hd = m.holidays.length, wd = m.weekdays;
    function chg(a, b) { if (!isCur) return ''; var d = a - b; return '<span class="chg ' + (d ? '' : 'd-0') + '">' + (d ? (d > 0 ? '+' : '−') + Math.abs(d) + '일' : '변동 없음') + '</span>'; }
    return '<div class="cal-month"><div class="eyebrow">' + (isCur ? '이번 달' : '전달') + '</div><h3>' + OTCore.monthLabel(m) + '</h3>' +
      '<div class="cal-stats">' +
      '<div><span class="muted" style="font-size:.74rem">전체</span><b>' + m.nDays + '</b>' + chg(m.nDays, other.nDays) + '</div>' +
      '<div><span class="muted" style="font-size:.74rem">평일</span><b>' + wd + '</b>' + chg(wd, other.weekdays) + '</div>' +
      '<div><span class="muted" style="font-size:.74rem">휴일</span><b style="color:var(--hol)">' + hd + '</b>' + chg(hd, other.holidays.length) + '</div>' +
      '<div><span class="muted" style="font-size:.74rem">기본급 포함 B/O</span><b>' + wd * 2 + 'h</b><span class="chg muted">평일×2h</span></div>' +
      '</div>' + miniCal(m) +
      '<p class="muted" style="font-size:.76rem">휴일: ' + m.holidays.map(function (d) { return d + '일(' + (m.dayInfo[d - 1].wdKo || '') + ')'; }).join(', ') + ' · 기준: ' + esc(m.holSource) + '</p></div>';
  }
  function secCalendar(r) {
    return '<section class="block"><div class="sec-head"><h2>달력 비교</h2><span class="muted">휴일 OT 수식에 들어간 날짜를 휴일로 읽습니다</span></div>' +
      '<div class="panel cal-grid">' + calMonth(r.prev, r.cur, false) + calMonth(r.cur, r.prev, true) + '</div></section>';
  }

  function secTotals(r) {
    var P = r.P, C = r.C;
    function row(cls, label, sub, a, b, f) {
      var d = b - a;
      return '<tr class="' + cls + '"><td>' + label + (sub ? '<span class="sub">' + sub + '</span>' : '') + '</td><td class="num">' + f(a) + '</td><td class="num">' + f(b) + '</td><td class="num ' + dcls(d) + '">' + sgn(d, f) + '</td><td class="num ' + dcls(d) + '">' + (a ? sgn(pct(a, b), function (x) { return x.toFixed(1) + '%'; }) : '—') + '</td></tr>';
    }
    var hrs = function (x) { return h(x) + 'h'; };
    var perWd = function (T) { return T.wdPersonDays ? T.wdOt / T.wdPersonDays : 0; };
    var perHd = function (T) { return T.hdPersonDays ? T.hdOt / T.hdPersonDays : 0; };
    return '<section class="block"><div class="sec-head"><h2>팀 합계</h2><span class="muted">시트의 요약 열(평일근무 OT·휴일근무 OT·급여) 기준</span></div>' +
      '<div class="panel ledger"><table><thead><tr><th>항목</th><th class="num">' + r.label.p + '</th><th class="num">' + r.label.c + '</th><th class="num">증감</th><th class="num">증감률</th></tr></thead><tbody>' +
      row('row-wd', '평일 OT 시간', '1.5배 ' + h(P.wd15) + '→' + h(C.wd15) + 'h · 2배 ' + h(P.wd2) + '→' + h(C.wd2) + 'h', P.wdOt, C.wdOt, hrs) +
      row('row-wd', '평일 OT 1인·1일 평균', '평일 근무 연인원 ' + P.wdPersonDays + '→' + C.wdPersonDays + '일', perWd(P), perWd(C), function (x) { return x.toFixed(2) + 'h'; }) +
      row('row-wd', '평일 OT 금액', '', P.payWd, C.payWd, usd) +
      row('row-hol', '휴일 OT 시간', '휴일 근무 연인원 ' + P.hdPersonDays + '→' + C.hdPersonDays + '일', P.hdOt, C.hdOt, hrs) +
      row('row-hol', '휴일 OT 1인·1회 평균', '', perHd(P), perHd(C), function (x) { return x.toFixed(2) + 'h'; }) +
      row('row-hol', '휴일 OT 금액', '', P.payHd, C.payHd, usd) +
      row('', '기본급', '', P.payBase, C.payBase, usd) +
      row('', '월급여 합계', '인원 ' + P.head + '→' + C.head + '명', P.pay, C.pay, usd) +
      '</tbody></table></div></section>';
  }

  function secFindings(r) {
    var F = [];
    var both = r.rows.filter(function (x) { return x.status === 'both'; });
    // calendar
    var hd = r.holDiff, wd = r.wdDiff;
    if (hd || wd) {
      var sumHdCal = both.reduce(function (s, x) { return s + x.hdDaysEffect; }, 0);
      var sumWdCal = both.reduce(function (s, x) { return s + x.wdDaysEffect; }, 0);
      F.push(['달력', '휴일이 <b>' + r.prev.holidays.length + '일 → ' + r.cur.holidays.length + '일</b>, 평일이 <b>' + r.prev.weekdays + '일 → ' + r.cur.weekdays + '일</b>로 바뀌었습니다. 근무일수 변화(달력 + 결근·출국)만으로 설명되는 OT 변화는 평일 <b>' + sgn(sumWdCal, h) + 'h</b>, 휴일 <b>' + sgn(sumHdCal, h) + 'h</b>입니다.']);
    } else F.push(['달력', '두 달의 평일·휴일 일수가 같습니다(평일 ' + r.cur.weekdays + '일, 휴일 ' + r.cur.holidays.length + '일). OT 차이는 근무 패턴이나 인원 변화에서 나옵니다.']);
    // totals
    F.push(['평일 OT', '팀 평일 OT는 <b>' + h(r.P.wdOt) + 'h → ' + h(r.C.wdOt) + 'h</b>(' + sgn(r.C.wdOt - r.P.wdOt, h) + 'h), 금액은 <b>' + sgn(r.C.payWd - r.P.payWd, usd) + '</b>입니다.' + patternLine(both, 'wd')]);
    F.push(['휴일 OT', '팀 휴일 OT는 <b>' + h(r.P.hdOt) + 'h → ' + h(r.C.hdOt) + 'h</b>(' + sgn(r.C.hdOt - r.P.hdOt, h) + 'h), 금액은 <b>' + sgn(r.C.payHd - r.P.payHd, usd) + '</b>입니다.' + patternLine(both, 'hd')]);
    // shifts
    var sh = both.filter(function (x) { return x.wdShift.changed || x.hdShift.changed || Math.abs(x.wdPerDayDelta) >= 0.25; });
    if (sh.length) F.push(['시간대 변경', sh.map(function (x) {
      var m = x.wdShift.moves.filter(function (mv) { return Math.abs(mv.d) >= 0.05; }).slice(0, 2).map(function (mv) { return mv.key + ' ' + mv.pn + '일→' + mv.cn + '일'; }).join(', ');
      return '<b>' + esc(x.name) + '</b>: 평일 ' + (m || '퇴근 시각 분포 변화') + ' (평일 1일 평균 OT ' + sgn(x.wdPerDayDelta, function (v) { return v.toFixed(2); }) + 'h)';
    }).join('<br>')]);
    else F.push(['시간대 변경', '평일·휴일 출퇴근 시간대가 바뀐 작업자는 없습니다.']);
    // attendance
    var att = both.filter(function (x) { return x.c.A.offDays > 0 && (x.c.A.wdDays + x.c.A.hdDays) < r.cur.nDays && x.notes.length; });
    var newG = r.rows.filter(function (x) { return x.status !== 'both'; });
    if (att.length || newG.length) F.push(['인원 변동', att.map(function (x) { return '<b>' + esc(x.name) + '</b>: ' + x.notes.join(', ') + ' → ' + (x.c.A.wdDays + x.c.A.hdDays) + '일 근무'; }).concat(newG.map(function (x) { return '<b>' + esc(x.name) + '</b>: ' + (x.status === 'new' ? '이번 달 신규' : '이번 달 명단에 없음'); })).join('<br>')]);
    // checks
    var chk = both.filter(function (x) { return x.c.A.checks.length || x.c.A.gaps.length; });
    var warns = r.prev.warnings.concat(r.cur.warnings);
    if (chk.length || warns.length) F.push(['데이터 점검', warns.map(esc).concat(chk.map(function (x) { return '<b>' + esc(x.name) + '</b>: ' + x.c.A.checks.concat(x.c.A.gaps.filter(function (g) { return x.c.A.bwMissing.indexOf(g.d) < 0; }).map(function (g) { return g.d + '일 출퇴근 ' + fm(g.inM) + '–' + fm(g.outM) + '(' + g.span + 'h)와 기록 ' + g.rec + 'h 차이'; })).map(esc).join(' · '); })).join('<br>')]);
    return '<section class="block"><div class="sec-head"><h2>주요 변화</h2></div><div class="panel findings">' +
      F.map(function (f) { return '<div class="finding"><span class="eyebrow" style="padding-top:2px">' + f[0] + '</span><p>' + f[1] + '</p></div>'; }).join('') + '</div></section>';
  }
  function patternLine(both, k) {
    var cal = 0, pat = 0;
    both.forEach(function (x) { cal += x[k + 'DaysEffect']; pat += x[k + 'PatternEffect']; });
    return ' 이 중 근무일수 효과 <b>' + sgn(cal, h) + 'h</b>, 근무 패턴(하루 OT 시간) 효과 <b>' + sgn(pat, h) + 'h</b>' + (both.length < (state.r ? state.r.rows.length : 0) ? ', 나머지는 신규·제외 인원' : '') + '.';
  }

  /* effect chart (diverging horizontal bars, one scale) */
  function effectSvg(r, k) {
    var rows = r.rows.filter(function (x) { return x.status === 'both'; });
    var vals = []; rows.forEach(function (x) { vals.push(x[k + 'DaysEffect'], x[k + 'PatternEffect']); });
    var maxAbs = Math.max.apply(null, vals.map(Math.abs).concat([1]));
    var step = niceStep(maxAbs);
    var lim = Math.ceil(maxAbs / step) * step;
    var W = 760, L = 190, Rm = 50, rowH = 34, top = 26, H = top + rows.length * rowH + 8;
    var x0 = L + (W - L - Rm) / 2, sc = (W - L - Rm) / 2 / lim;
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%" style="min-width:560px;display:block" role="img" aria-label="작업자별 OT 증감 요인">';
    for (var t = -lim; t <= lim + 1e-9; t += step) {
      var xx = x0 + t * sc;
      s += '<line x1="' + xx + '" x2="' + xx + '" y1="' + (top - 6) + '" y2="' + (H - 4) + '" stroke="var(--rule-2)" stroke-width="1"/>';
      s += '<text x="' + xx + '" y="' + (top - 10) + '" text-anchor="middle" font-size="10" fill="var(--ink-3)" style="font-family:var(--font-num)">' + (t > 0 ? '+' : t < 0 ? '−' : '') + Math.abs(R(t, 1)) + 'h</text>';
    }
    s += '<line x1="' + x0 + '" x2="' + x0 + '" y1="' + (top - 6) + '" y2="' + (H - 4) + '" stroke="var(--ink-3)" stroke-width="1"/>';
    rows.forEach(function (x, i) {
      var y = top + i * rowH;
      var a = x[k + 'DaysEffect'], b = x[k + 'PatternEffect'];
      s += '<text x="' + (L - 12) + '" y="' + (y + rowH / 2 + 4) + '" text-anchor="end" font-size="11.5" fill="var(--ink)">' + esc(short(x.name)) + '</text>';
      [[a, 'var(--eff-cal)', y + 6, '근무일수 효과'], [b, 'var(--eff-pat)', y + 18, '패턴 효과']].forEach(function (q) {
        var v = q[0], w = Math.abs(v) * sc, bx = v >= 0 ? x0 : x0 - w;
        var tipTxt = esc(x.name) + '<br>' + q[3] + ' <b>' + sgn(v, h) + 'h</b>';
        s += '<rect x="' + bx + '" y="' + q[2] + '" width="' + Math.max(w, 1) + '" height="10" rx="2" fill="' + q[1] + '" data-tip="' + tipTxt + '"/>';
        if (Math.abs(v) >= 0.05) s += '<text x="' + (v >= 0 ? bx + w + 4 : bx - 4) + '" y="' + (q[2] + 9) + '" text-anchor="' + (v >= 0 ? 'start' : 'end') + '" font-size="10" fill="var(--ink-2)" style="font-family:var(--font-num)">' + sgn(v, h) + '</text>';
      });
      s += '<rect x="0" y="' + y + '" width="' + W + '" height="' + rowH + '" fill="transparent" data-tip="' + esc(x.name) + '<br>근무일수 효과 <b>' + sgn(a, h) + 'h</b><br>패턴 효과 <b>' + sgn(b, h) + 'h</b><br>합계 <b>' + sgn(a + b, h) + 'h</b>" style="pointer-events:' + 'all' + ';opacity:0"/>';
    });
    return s + '</svg>';
  }
  function niceStep(m) { var raw = m / 3, p = Math.pow(10, Math.floor(Math.log10(raw))), n = raw / p; return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * p; }
  function short(n) { return n.length > 24 ? n.slice(0, 23) + '…' : n; }

  function secEffects(r) {
    return '<section class="block" id="effects"><div class="sec-head"><h2>OT 증감 요인</h2>' +
      '<div class="tabs" role="tablist"><button class="tab" role="tab" id="tabWd" data-k="wd" aria-selected="' + (state.effTab === 'wd') + '">평일 OT</button><button class="tab" role="tab" id="tabHd" data-k="hd" aria-selected="' + (state.effTab === 'hd') + '">휴일 OT</button></div></div>' +
      '<div class="panel chart-box"><div class="legend"><span><i style="background:var(--eff-cal)"></i>근무일수 효과: 전달 1일 평균 OT × 이번 달 근무일수 − 전달 OT</span><span><i style="background:var(--eff-pat)"></i>패턴 효과: 하루 OT 시간이 달라져서 생긴 차이</span></div>' +
      '<div class="chart-scroll" id="effChart">' + effectSvg(r, state.effTab) + '</div></div></section>';
  }
  function bindTabs() {
    ['tabWd', 'tabHd'].forEach(function (id) {
      var b = $(id); if (!b) return;
      b.addEventListener('click', function () {
        state.effTab = b.getAttribute('data-k');
        $('tabWd').setAttribute('aria-selected', state.effTab === 'wd'); $('tabHd').setAttribute('aria-selected', state.effTab === 'hd');
        $('effChart').innerHTML = effectSvg(state.r, state.effTab);
      });
    });
  }

  /* shift table */
  function distTxt(d) { return d.slice(0, 3).map(function (e) { return '<span class="num">' + e.key + '</span> <span class="muted">' + e.n + '일</span>'; }).join('<br>') || '<span class="muted">—</span>'; }
  function secShift(r) {
    var rows = r.rows.filter(function (x) { return x.status === 'both'; });
    return '<section class="block"><div class="sec-head"><h2>근무 시간대 비교</h2><span class="muted">출근–퇴근 조합별 일수, 많은 순 3개</span></div>' +
      '<div class="panel ledger"><table style="min-width:760px"><thead><tr><th>작업자</th><th style="text-align:left">평일 · ' + r.label.p + '</th><th style="text-align:left">평일 · ' + r.label.c + '</th><th style="text-align:left">휴일 · ' + r.label.p + '</th><th style="text-align:left">휴일 · ' + r.label.c + '</th><th>판정</th></tr></thead><tbody>' +
      rows.map(function (x) {
        var tags = [];
        if (x.wdShift.changed) tags.push('<span class="chip alert">평일 변경</span>');
        else if (Math.abs(x.wdPerDayDelta) >= 0.25) tags.push('<span class="chip alert">평일 OT/일 ' + sgn(x.wdPerDayDelta, function (v) { return v.toFixed(2); }) + 'h</span>');
        if (x.hdShift.changed) tags.push('<span class="chip alert">휴일 변경</span>');
        if (!tags.length) tags.push('<span class="chip ok">동일</span>');
        return '<tr><td style="vertical-align:top"><b style="font-weight:600">' + esc(x.name) + '</b></td>' +
          [x.p.A.wdDist, x.c.A.wdDist, x.p.A.hdDist, x.c.A.hdDist].map(function (d) { return '<td style="text-align:left;vertical-align:top;font-size:.8rem;line-height:1.5">' + distTxt(d) + '</td>'; }).join('') +
          '<td style="vertical-align:top"><div class="chips" style="justify-content:flex-end">' + tags.join('') + '</div></td></tr>';
      }).join('') + '</tbody></table></div></section>';
  }

  /* worker detail */
  function seq(x) {
    var v = x.ot;
    if (!x.worked) return { cls: 'off', bg: '' };
    var ramp = x.hol ? ['--o1', '--o2', '--o3', '--o4'] : ['--b1', '--b2', '--b3', '--b4', '--b5'];
    var cuts = x.hol ? [4, 8, 11] : [2, 4, 5.5, 7];
    var i = 0; while (i < cuts.length && v > cuts[i]) i++;
    if (v <= 0) return { cls: x.hol ? 'hdc' : '', bg: '' };
    return { cls: (x.hol ? 'hdc ' : '') + (i >= (x.hol ? 2 : 3) ? 'dark' : ''), bg: 'background:var(' + ramp[i] + ')' };
  }
  function strip(m, w, label) {
    var main = w.A.wdDist[0] ? w.A.wdDist[0].key : null, hmain = w.A.hdDist[0] ? w.A.hdDist[0].key : null;
    var s = '<div class="lbl">' + label + '</div>';
    w.days.forEach(function (x) {
      var k = fm(x.inM) + '–' + fm(x.outM);
      var dev = x.worked && x.inM != null && ((x.hol && hmain && k !== hmain) || (!x.hol && main && k !== main));
      var st = seq(x);
      var di = m.dayInfo[x.d - 1];
      var tipTxt = '<b>' + x.d + '일(' + di.wdKo + ')</b> ' + (x.hol ? '휴일' : '평일') + '<br>' + (x.note ? esc(x.note) + '<br>' : '') + (x.inM != null ? '출퇴근 <b>' + k + '</b><br>' : '') + 'B/W ' + x.bw + ' · B/O ' + x.bo + ' · E/O ' + x.eo + (dev ? '<br>이 달 주 패턴(' + (x.hol ? hmain : main) + ')과 다름' : '');
      s += '<div class="cell ' + st.cls + (dev ? ' dev' : '') + '" style="' + st.bg + '" data-tip="' + tipTxt + '">' + (x.worked ? (x.ot ? h(x.ot) : '') : (x.note ? '×' : '')) + '</div>';
    });
    for (var i = w.days.length; i < 31; i++) s += '<div></div>';
    return s;
  }
  function stripHead(m) {
    var s = '<div></div>';
    m.dayInfo.forEach(function (d) { s += '<div class="dh' + (d.hol ? ' hd' : '') + '">' + d.d + '<br>' + d.wdKo + '</div>'; });
    for (var i = m.dayInfo.length; i < 31; i++) s += '<div></div>';
    return s;
  }
  function distBlock(title, pd, cd, lp, lc) {
    var keys = {}; pd.forEach(function (e) { keys[e.key] = 1; }); cd.forEach(function (e) { keys[e.key] = 1; });
    var all = Object.keys(keys).map(function (k) {
      var a = pd.filter(function (e) { return e.key === k; })[0], b = cd.filter(function (e) { return e.key === k; })[0];
      return { k: k, pn: a ? a.n : 0, cn: b ? b.n : 0, ps: a ? a.share : 0, cs: b ? b.share : 0 };
    }).sort(function (x, y) { return (y.pn + y.cn) - (x.pn + x.cn); });
    if (!all.length) return '<div class="dist"><h3>' + title + '</h3><p class="muted">근무 기록 없음</p></div>';
    return '<div class="dist"><h3>' + title + '</h3><div class="legend" style="font-size:.72rem"><span><i style="background:var(--eff-cal)"></i>' + lp + '</span><span><i style="background:var(--accent)"></i>' + lc + '</span></div>' +
      all.map(function (e) {
        return '<div class="dist-row" data-tip="' + e.k + '<br>' + lp + ' <b>' + e.pn + '일</b> · ' + lc + ' <b>' + e.cn + '일</b>"><span class="k">' + e.k + '</span><span class="dist-bars"><i class="p" style="width:' + (e.ps * 100) + '%"></i><i class="c" style="width:' + (e.cs * 100) + '%"></i></span><span class="v">' + e.pn + '일 → ' + e.cn + '일</span></div>';
      }).join('') + '</div>';
  }
  function workerRow(x, r) {
    var p = x.p, c = x.c;
    var tags = x.tags.map(function (t) { return '<span class="chip ' + t.s + '">' + esc(t.t) + '</span>'; }).join('') || '<span class="chip ok">특이사항 없음</span>';
    if (c && c.A.manualAdj) tags += '<span class="chip info" data-tip="기본급 포함 B/O 수식: =' + esc(c.A.manualAdj) + '">B/O 수식 조정</span>';
    function metric(lab, a, b, f) {
      if (a == null || b == null) return '<div class="w-metric"><span class="lab">' + lab + '</span>' + (b != null ? f(b) : a != null ? f(a) : '—') + '</div>';
      var d = b - a;
      return '<div class="w-metric"><span class="lab">' + lab + '</span><span>' + f(a) + ' → ' + f(b) + '</span><span class="' + dcls(d) + '" style="font-size:.78rem">' + sgn(d, f) + '</span></div>';
    }
    var hrs = function (v) { return h(v) + 'h'; };
    var sum = '<summary><div class="w-name"><span class="caret">›</span><span style="min-width:0">' + esc(x.name) + ' <small>' + esc(x.id) + '</small></span></div>' +
      metric('평일 OT', p && p.sheet.wdTot, c && c.sheet.wdTot, hrs) +
      metric('휴일 OT', p && p.sheet.hdTot, c && c.sheet.hdTot, hrs) +
      metric('급여', p && p.sheet.pay, c && c.sheet.pay, function (v) { return '$' + Math.round(v).toLocaleString('en-US'); }) +
      '<div class="chips">' + tags + '</div></summary>';
    var body = '<div class="w-body">';
    body += '<div class="strip-legend"><span>칸 안 숫자 = 그날 OT(B/O+E/O) 시간</span><span>평일<span class="ramp"><i style="background:var(--b1)"></i><i style="background:var(--b2)"></i><i style="background:var(--b3)"></i><i style="background:var(--b4)"></i><i style="background:var(--b5)"></i></span>적음→많음</span><span>휴일<span class="ramp"><i style="background:var(--o1)"></i><i style="background:var(--o2)"></i><i style="background:var(--o3)"></i><i style="background:var(--o4)"></i></span></span><span><span class="ramp"><i style="outline:2px solid var(--eff-pat);outline-offset:-2px;background:var(--s0)"></i></span>그 달 주 패턴과 다른 시간대</span><span><span class="ramp"><i style="background:repeating-linear-gradient(135deg,var(--s0) 0 3px,var(--rule) 3px 4px)"></i></span>근무 없음</span></div>';
    body += '<div class="strip-wrap"><div class="strip">' + (p ? stripHead(r.prev) + strip(r.prev, p, r.label.p) : '') + '</div></div>';
    body += '<div class="strip-wrap"><div class="strip">' + (c ? stripHead(r.cur) + strip(r.cur, c, r.label.c) : '') + '</div></div>';
    if (p && c) {
      body += '<div class="two-col">' + distBlock('평일 출퇴근 시간대', p.A.wdDist, c.A.wdDist, r.label.p, r.label.c) + distBlock('휴일 출퇴근 시간대', p.A.hdDist, c.A.hdDist, r.label.p, r.label.c) + '</div>';
      body += '<div class="two-col"><div class="decomp"><h3 style="color:var(--ink)">평일 OT 변화 분해</h3>' +
        '<span>근무일 ' + p.A.wdDays + '일 → ' + c.A.wdDays + '일 · 1일 평균 ' + p.A.wdOtPerDay.toFixed(2) + 'h → ' + c.A.wdOtPerDay.toFixed(2) + 'h</span>' +
        '<span>근무일수 효과 <b>' + sgn(x.wdDaysEffect, h) + 'h</b> · 패턴 효과 <b>' + sgn(x.wdPatternEffect, h) + 'h</b> · 합계 <b>' + sgn(x.dWdOt, h) + 'h</b> (' + sgn(x.dPayWd, usd) + ')</span></div>' +
        '<div class="decomp"><h3 style="color:var(--ink)">휴일 OT 변화 분해</h3>' +
        '<span>휴일 근무 ' + p.A.hdDays + '회 → ' + c.A.hdDays + '회 · 1회 평균 ' + p.A.hdOtPerDay.toFixed(2) + 'h → ' + c.A.hdOtPerDay.toFixed(2) + 'h</span>' +
        '<span>근무일수 효과 <b>' + sgn(x.hdDaysEffect, h) + 'h</b> · 패턴 효과 <b>' + sgn(x.hdPatternEffect, h) + 'h</b> · 합계 <b>' + sgn(x.dHdOt, h) + 'h</b> (' + sgn(x.dPayHd, usd) + ')</span></div></div>';
      var items = [];
      x.notes.forEach(function (n) { items.push(esc(n)); });
      c.A.checks.forEach(function (n) { items.push(esc(n)); });
      c.A.gaps.forEach(function (g) { if (c.A.bwMissing.indexOf(g.d) < 0) items.push(g.d + '일: 출퇴근 ' + fm(g.inM) + '–' + fm(g.outM) + '(' + g.span + 'h)인데 기록 합계 ' + g.rec + 'h'); });
      if (c.A.manualAdj) items.push('기본급 포함 B/O 수식이 수동 조정됨: =' + esc(c.A.manualAdj) + ' → ' + c.A.builtinSheet + 'h (B/W·B/O가 모두 있는 평일 ' + c.A.wdBoDays + '일×2h = ' + c.A.builtinExpected + 'h' + (c.A.builtinExpected === c.A.builtinSheet ? ', 일치' : ', 불일치') + ')');
      if (Math.abs(c.rate - p.rate) > 1e-6) items.push('시급 ' + R(p.rate, 4) + ' → ' + R(c.rate, 4));
      if (items.length) body += '<div><h3>확인할 점</h3><ul class="checks">' + items.map(function (i) { return '<li>' + i + '</li>'; }).join('') + '</ul></div>';
    }
    return '<details class="w" id="w-' + esc(x.key).replace(/[^a-z0-9]/gi, '') + '">' + sum + body + '</div></details>';
  }
  function secWorkers(r) {
    return '<section class="block"><div class="sec-head"><h2>작업자별 상세</h2><span class="muted">줄을 누르면 일별 근무표와 시간대 분포가 열립니다</span></div>' +
      '<div class="panel"><div class="w-head"><span>작업자</span><span>평일 OT</span><span>휴일 OT</span><span>월급여</span><span>표시</span></div><div class="workers">' +
      r.rows.map(function (x) { return workerRow(x, r); }).join('') + '</div></div></section>';
  }
  function secMethod() {
    return '<section class="block"><div class="sec-head"><h2>계산 기준</h2></div><div class="panel method"><ul>' +
      '<li>평일 OT = 시트의 "평일근무 OT 합계"(B/O − 기본급 포함분 + 평일 E/O). 휴일 OT = "휴일근무 OT 합계"(휴일 E/O).</li>' +
      '<li>휴일은 휴일 OT 수식(예: =L8+S8+Z8+AG8)에 들어간 날짜로 판단하고, B/O가 비어 있는 날과 다르면 경고합니다.</li>' +
      '<li>휴일에는 B/W 8h가 일괄 기재되므로, 휴일 근무 여부는 E/O 유무로 셉니다.</li>' +
      '<li>근무일수 효과 = 전달 1일 평균 OT × 이번 달 근무일수 − 전달 OT. 패턴 효과 = 나머지(하루 OT 시간 변화).</li>' +
      '<li>시간대 변경 판정: 가장 많은 출퇴근 조합이 바뀌었거나, 한 조합의 비중이 15%p 이상 바뀐 경우. 평일 1일 평균 OT가 0.25h 이상 바뀌면 따로 표시합니다.</li>' +
      '<li>파일은 이 브라우저 안에서만 읽고, 어디에도 업로드하지 않습니다.</li></ul></div></section>';
  }

  /* sample data (js/sample-data.js, 익명화된 예시) */
  function loadSample() {
    showErr('');
    if (typeof SAMPLE_FILES === 'undefined' || !SAMPLE_FILES.length) { showErr('예시 데이터가 없습니다. 타임시트 파일을 올려 주세요.'); return; }
    try {
      var ms = SAMPLE_FILES.map(function (f) {
        var bin = atob(f.b64), u = new Uint8Array(bin.length);
        for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
        return OTCore.parseWorkbook(XLSX.read(u, { type: 'array', cellFormula: true }), f.name);
      });
      setMonths(ms[0], ms[1], true);
    } catch (e) { showErr('예시 데이터를 읽지 못했습니다: ' + e.message); }
  }
  $('loadSample').addEventListener('click', loadSample);
  if (typeof XLSX === 'undefined') { showErr('엑셀 읽기 라이브러리(SheetJS)를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.'); return; }
  loadSample();
})();
