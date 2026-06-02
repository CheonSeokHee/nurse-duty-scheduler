/**
 * 간호사 3교대 듀티표 자동 생성기 (Google Apps Script)
 * ---------------------------------------------------------------
 * - 시트: "설정" + "듀티표"
 * - 메뉴: 듀티표 > [① 시트 세팅] [② 자동 배정] [③ 규칙 검사] [표 비우기]
 *
 * 규칙 요약 (설정 시트에서 변경 가능):
 *  - 근무: Day / Evening / Night / Off(O)
 *  - 하루 필요인원: D=3, E=3, N=2  (나머지는 Off)
 *  - 나이트는 3연속 고정 (N-N-N)
 *  - 나이트 시작 전날 = 오프
 *  - 나이트 끝나고 = 2오프
 *  - 연속근무 최대 4일 (주간근무 기준)
 *  - 금지 패턴: N 다음날 D/E,  E 다음날 D
 *  - 모든 근무에 차지(Charge) 1명 이상 (액팅만 근무 불가)
 *  - 월 오프 1인당 10~11개
 */

/* ===================== 상수 ===================== */
var SETTINGS_SHEET = '설정';
var DUTY_SHEET = '듀티표';

var SHIFT = { D: 'D', E: 'E', N: 'N', O: 'O' };
var WORK_SHIFTS = ['D', 'E', 'N'];

var COLORS = {
  D: '#cfe2f3', // 연파랑
  E: '#fce5cd', // 연주황
  N: '#434b66', // 남색 (글자 흰색)
  O: '#efefef', // 회색
  HEADER: '#1f3864',
  WEEKEND: '#fff2cc',
  VIOLATION: '#f4cccc' // 위반 빨강
};

var NURSE_START_ROW = 19; // 설정 시트에서 간호사 목록 시작 행
var DUTY_DATA_START_ROW = 4; // 듀티표에서 간호사 데이터 시작 행 (1:제목,2:일,3:요일)

/* ===================== 메뉴 ===================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🩺 듀티표')
    .addItem('① 시트 세팅(최초 1회)', 'setupSheets')
    .addSeparator()
    .addItem('② 자동 배정 실행', 'generateDuty')
    .addItem('③ 규칙 검사 / 위반 표시', 'checkRules')
    .addItem('칸 넓히기 / 보기 좋게', 'formatLayout')
    .addSeparator()
    .addItem('표 내용만 비우기', 'clearDutyValues')
    .addToUi();
}

/* ===================== 시트 세팅 ===================== */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 설정 시트 ---
  var s = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);
  s.clear();
  s.getRange('A1').setValue('■ 기본 설정').setFontWeight('bold');
  var rows = [
    ['연도', new Date().getFullYear()],
    ['월', new Date().getMonth() + 1],
    ['', ''],
    ['Day 필요인원', 3],
    ['Evening 필요인원', 3],
    ['Night 필요인원', 2],
    ['', ''],
    ['최대 연속근무(일)', 4],
    ['나이트 연속(일)', 3],
    ['나이트 전 오프(3연속시,일)', 1],
    ['나이트 후 오프(일)', 1],
    ['1인 월 최대 나이트(블록)', 2],
    ['나이트 분배(1인원/2균형/3균등)', 3],
    ['월 오프 최소', 10],
    ['월 오프 최대', 11],
    ['자동배정 시도횟수', 300]
  ];
  s.getRange(2, 1, rows.length, 2).setValues(rows);
  s.getRange(2, 1, rows.length, 1).setFontWeight('bold');

  // 간호사 목록 헤더
  s.getRange(NURSE_START_ROW - 1, 1).setValue('■ 간호사 목록 (역할: 차지 / 액팅)').setFontWeight('bold');
  var headers = [['이름', '역할', '요청오프(예: 3,10,21)', '요청근무(예: D:5 / N:12)', '전월 나이트(일)']];
  s.getRange(NURSE_START_ROW, 1, 1, 5).setValues(headers)
    .setFontWeight('bold').setBackground(COLORS.HEADER).setFontColor('#ffffff');

  // 기본 12명 (차지 6 / 액팅 6) 예시
  var nurses = [
    ['간호사1', '차지', '', ''],
    ['간호사2', '차지', '', ''],
    ['간호사3', '차지', '', ''],
    ['간호사4', '차지', '', ''],
    ['간호사5', '차지', '', ''],
    ['간호사6', '차지', '', ''],
    ['간호사7', '액팅', '', ''],
    ['간호사8', '액팅', '', ''],
    ['간호사9', '액팅', '', ''],
    ['간호사10', '액팅', '', ''],
    ['간호사11', '액팅', '', ''],
    ['간호사12', '액팅', '', '']
  ];
  s.getRange(NURSE_START_ROW + 1, 1, nurses.length, 4).setValues(nurses);

  // 역할 드롭다운
  var roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['차지', '액팅'], true).build();
  s.getRange(NURSE_START_ROW + 1, 2, 50, 1).setDataValidation(roleRule);

  s.setColumnWidth(1, 130);
  s.setColumnWidth(2, 80);
  s.setColumnWidth(3, 180);
  s.setColumnWidth(4, 200);
  s.setColumnWidth(5, 110);

  // --- 듀티표 시트 ---
  var d = ss.getSheetByName(DUTY_SHEET) || ss.insertSheet(DUTY_SHEET);
  drawDutyTemplate();

  ss.setActiveSheet(s);
  SpreadsheetApp.getUi().alert(
    '세팅 완료!\n\n' +
    '1) [설정] 시트에서 연/월, 인원, 간호사 이름·역할을 입력하세요.\n' +
    '2) 메뉴 [🩺 듀티표 > ② 자동 배정 실행] 을 누르면 표가 생성됩니다.\n' +
    '3) 칸을 직접 고친 뒤 [③ 규칙 검사]로 위반을 확인하세요.'
  );
}

/* ===================== 설정 읽기 ===================== */
function readSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SETTINGS_SHEET);
  if (!s) throw new Error('"설정" 시트가 없습니다. 먼저 [① 시트 세팅]을 실행하세요.');

  function num(row) { return Number(s.getRange(row, 2).getValue()); }

  var cfg = {
    year: num(2),
    month: num(3),
    need: { D: num(5), E: num(6), N: num(7) },
    maxConsec: num(9),
    nightLen: num(10),
    offBeforeNight: num(11),
    offAfterNight: num(12),
    maxNightBlocks: num(13) || 2,
    nightMode: num(14) || 2, // 1=인원우선 2=균형 3=나이트균등
    offMin: num(15),
    offMax: num(16),
    attempts: num(17) || 300
  };

  // 간호사 목록 읽기
  var last = s.getLastRow();
  var nurses = [];
  var vals = s.getRange(NURSE_START_ROW + 1, 1, last - NURSE_START_ROW, 5).getValues();
  for (var i = 0; i < vals.length; i++) {
    var name = (vals[i][0] || '').toString().trim();
    if (!name) continue;
    var role = (vals[i][1] || '액팅').toString().trim();
    var prevNightDays = Number(vals[i][4]) || 0;
    nurses.push({
      name: name,
      charge: role === '차지',
      reqOff: parseReqOff(vals[i][2]),
      reqWork: parseReqWork(vals[i][3]),
      prevNightDays: prevNightDays,
      prevBlocks: Math.round(prevNightDays / (cfg.nightLen || 3))
    });
  }
  cfg.nurses = nurses;
  cfg.numDays = new Date(cfg.year, cfg.month, 0).getDate();
  return cfg;
}

function parseReqOff(v) {
  if (!v) return {};
  var out = {};
  ('' + v).split(',').forEach(function (x) {
    var n = parseInt(('' + x).trim(), 10);
    if (n >= 1) out[n] = true;
  });
  return out;
}

function parseReqWork(v) {
  // "D:5 / N:12" -> { 5:'D', 12:'N' }
  if (!v) return {};
  var out = {};
  ('' + v).split(/[\/,]/).forEach(function (tok) {
    var m = ('' + tok).trim().match(/^([DENO]):(\d+)$/i);
    if (m) out[parseInt(m[2], 10)] = m[1].toUpperCase();
  });
  return out;
}

/* ===================== 듀티표 템플릿 그리기 ===================== */
function drawDutyTemplate() {
  var cfg = readSettings();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d = ss.getSheetByName(DUTY_SHEET) || ss.insertSheet(DUTY_SHEET);
  d.clear();
  d.clearConditionalFormatRules();

  var nd = cfg.numDays;
  var nN = cfg.nurses.length;
  var firstDayCol = 2;          // B열부터 1일
  var sumStartCol = firstDayCol + nd; // 합계 컬럼 시작

  // 제목
  d.getRange(1, 1).setValue(cfg.year + '년 ' + cfg.month + '월 듀티표')
    .setFontWeight('bold').setFontSize(13);

  // 일/요일 헤더
  var weekKr = ['일', '월', '화', '수', '목', '금', '토'];
  d.getRange(2, 1).setValue('이름 \\ 일').setFontWeight('bold');
  for (var day = 1; day <= nd; day++) {
    var col = firstDayCol + day - 1;
    d.getRange(2, col).setValue(day);
    var dow = new Date(cfg.year, cfg.month - 1, day).getDay();
    d.getRange(3, col).setValue(weekKr[dow]);
    if (dow === 0 || dow === 6) {
      d.getRange(2, col, 1 + nN + 1, 1).setBackground(COLORS.WEEKEND);
      d.getRange(3, col).setFontColor(dow === 0 ? '#cc0000' : '#1155cc');
    }
  }
  d.getRange(2, 1, 2, sumStartCol + 3).setFontWeight('bold');

  // 합계 헤더
  var sumHeaders = ['D계', 'E계', 'N계', 'OFF'];
  d.getRange(2, sumStartCol, 1, 4).setValues([sumHeaders])
    .setBackground(COLORS.HEADER).setFontColor('#ffffff').setFontWeight('bold');

  // 간호사 행
  for (var i = 0; i < nN; i++) {
    var row = DUTY_DATA_START_ROW + i;
    var nurse = cfg.nurses[i];
    var label = nurse.name + (nurse.charge ? ' (차지)' : '');
    d.getRange(row, 1).setValue(label);
    if (nurse.charge) d.getRange(row, 1).setFontWeight('bold');
    // 합계 수식
    var dataA1 = d.getRange(row, firstDayCol, 1, nd).getA1Notation();
    d.getRange(row, sumStartCol + 0).setFormula('=COUNTIF(' + dataA1 + ',"D")');
    d.getRange(row, sumStartCol + 1).setFormula('=COUNTIF(' + dataA1 + ',"E")');
    d.getRange(row, sumStartCol + 2).setFormula('=COUNTIF(' + dataA1 + ',"N")');
    d.getRange(row, sumStartCol + 3).setFormula('=COUNTIF(' + dataA1 + ',"O")');
  }

  // 하단 일별 집계 행
  var base = DUTY_DATA_START_ROW + nN;
  var footLabels = ['Day 합', 'Evening 합', 'Night 합', '차지 부족?'];
  for (var f = 0; f < footLabels.length; f++) {
    d.getRange(base + f, 1).setValue(footLabels[f]).setFontWeight('bold');
  }
  for (var day2 = 1; day2 <= nd; day2++) {
    var c = firstDayCol + day2 - 1;
    var colData = d.getRange(DUTY_DATA_START_ROW, c, nN, 1).getA1Notation();
    d.getRange(base + 0, c).setFormula('=COUNTIF(' + colData + ',"D")');
    d.getRange(base + 1, c).setFormula('=COUNTIF(' + colData + ',"E")');
    d.getRange(base + 2, c).setFormula('=COUNTIF(' + colData + ',"N")');
  }

  // 입력 데이터 검증 (D/E/N/O)
  var dvRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['D', 'E', 'N', 'O'], true).build();
  d.getRange(DUTY_DATA_START_ROW, firstDayCol, nN, nd).setDataValidation(dvRule);

  // 색상 (조건부 서식)
  applyShiftColors(d, DUTY_DATA_START_ROW, firstDayCol, nN, nd);

  formatLayout();
}

/* 칸 너비/높이/정렬을 보기 좋게 (데이터 보존) */
function formatLayout() {
  var cfg = readSettings();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d = ss.getSheetByName(DUTY_SHEET);
  if (!d) { SpreadsheetApp.getUi().alert('듀티표 시트가 없습니다.'); return; }

  var nd = cfg.numDays, nN = cfg.nurses.length;
  var firstDayCol = 2;
  var sumStartCol = firstDayCol + nd;

  // 열 너비: 이름 넓게, 날짜칸 넉넉히, 합계칸 보통
  d.setColumnWidth(1, 150);
  for (var c = firstDayCol; c < firstDayCol + nd; c++) d.setColumnWidth(c, 46);
  for (var sc = sumStartCol; sc < sumStartCol + 4; sc++) d.setColumnWidth(sc, 60);

  // 행 높이 키우기
  var totalRows = DUTY_DATA_START_ROW + nN + 4;
  d.setRowHeights(1, totalRows, 30);

  // 데이터 영역: 가운데 정렬 + 큰 글씨
  d.getRange(DUTY_DATA_START_ROW, firstDayCol, nN, nd)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontSize(12).setFontWeight('bold');

  // 일/요일 헤더 가운데 정렬
  d.getRange(2, firstDayCol, 2, nd).setHorizontalAlignment('center');

  d.setFrozenRows(3);
  d.setFrozenColumns(1);
  ss.toast('보기 좋게 정리했어요 (칸 넓힘 + 가운데 정렬)', '듀티표', 4);
}

function applyShiftColors(sheet, r0, c0, nN, nd) {
  var rng = sheet.getRange(r0, c0, nN, nd);
  var rules = sheet.getConditionalFormatRules();
  function rule(text, bg, fg) {
    var b = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text).setBackground(bg);
    if (fg) b.setFontColor(fg);
    return b.setRanges([rng]).build();
  }
  rules.push(rule('D', COLORS.D));
  rules.push(rule('E', COLORS.E));
  rules.push(rule('N', COLORS.N, '#ffffff'));
  rules.push(rule('O', COLORS.O, '#999999'));
  sheet.setConditionalFormatRules(rules);
}

/* ===================== 자동 배정 ===================== */
function generateDuty() {
  var cfg = readSettings();
  if (cfg.nurses.length === 0) {
    SpreadsheetApp.getUi().alert('간호사 목록이 비어있습니다. [설정] 시트를 확인하세요.');
    return;
  }

  // ── 표에 직접 입력해둔 칸 읽기 (부분 입력=고정, 빈칸만 채움) ──
  // 일부만 채워져 있으면 "입력 유지 모드"(찍은 칸 고정), 비었거나 꽉 찼으면 "새 배정 모드"
  var preset = null, presetCount = 0;
  var ss0 = SpreadsheetApp.getActiveSpreadsheet();
  var d0 = ss0.getSheetByName(DUTY_SHEET);
  if (d0) {
    try {
      var cur = d0.getRange(DUTY_DATA_START_ROW, 2, cfg.nurses.length, cfg.numDays).getValues();
      var filled = 0, empty = 0, tmp = [];
      for (var pi = 0; pi < cfg.nurses.length; pi++) {
        tmp[pi] = {};
        for (var pj = 0; pj < cfg.numDays; pj++) {
          var v = (cur[pi][pj] || '').toString().trim().toUpperCase();
          if (v === 'D' || v === 'E' || v === 'N' || v === 'O') { tmp[pi][pj + 1] = v; filled++; }
          else empty++;
        }
      }
      if (filled > 0 && empty > 0) { preset = tmp; presetCount = filled; } // 부분 입력 → 고정
    } catch (e) { preset = null; }
  }
  cfg.preset = preset; // tryBuild / isLocked 가 참고

  drawDutyTemplate();

  // 실행할 때마다 다른 배치가 나오도록 랜덤 베이스 시드 (재굴리기 가능)
  var base = Math.floor(Math.random() * 2000000000);
  var best = null;
  for (var a = 0; a < cfg.attempts; a++) {
    var rng = makeRng(base + a * 2654435761 + 12345);
    var sched = tryBuild(cfg, rng);
    var score = evaluate(cfg, sched);
    if (!best || score.total < best.score.total) best = { sched: sched, score: score };
    if (score.total === 0) break;
  }

  writeSchedule(cfg, best.sched);
  checkRules(); // 생성 직후 자동 검사

  var sc = best.score;
  var presetMsg = preset ? ('입력 ' + presetCount + '칸 고정 + ') : '';
  SpreadsheetApp.getActiveSpreadsheet().toast(
    presetMsg + '자동 배정 완료 (미충족 ' + sc.unfilled + ' / 오프편차 ' + sc.offDev +
    ' / 나이트편차 ' + sc.nightDev + ' / 위반 ' + sc.hard + ')',
    '듀티표', 6);
}

/* 한 번의 배정 시도 */
function tryBuild(cfg, rng) {
  var nd = cfg.numDays;
  var N = cfg.nurses.length;
  // sched[n][day] (day: 1..nd), 0번 인덱스 미사용
  var sched = [];
  for (var i = 0; i < N; i++) {
    sched[i] = [];
    for (var day = 0; day <= nd + 3; day++) sched[i][day] = '';
  }

  // 0) 요청근무/요청오프 먼저 고정
  for (var i2 = 0; i2 < N; i2++) {
    var nu = cfg.nurses[i2];
    for (var dd in nu.reqWork) {
      var di = parseInt(dd, 10);
      if (di >= 1 && di <= nd) sched[i2][di] = nu.reqWork[dd];
    }
    for (var od in nu.reqOff) {
      var oi = parseInt(od, 10);
      if (oi >= 1 && oi <= nd && !sched[i2][oi]) sched[i2][oi] = SHIFT.O;
    }
    // 표에 직접 입력해둔 칸(preset)도 고정
    if (cfg.preset && cfg.preset[i2]) {
      for (var pd in cfg.preset[i2]) {
        var pdi = parseInt(pd, 10);
        if (pdi >= 1 && pdi <= nd) sched[i2][pdi] = cfg.preset[i2][pd];
      }
    }
  }

  // 1) 나이트 블록 배정
  assignNights(cfg, sched, rng);

  // 2) Day / Evening 배정 (차지 커버리지 우선)
  for (var day3 = 1; day3 <= nd; day3++) {
    fillDayEvening(cfg, sched, day3, rng);
  }

  // 2-b) 오프 상한 보정: 일을 덜 한 사람(오프 > offMax)에게 빈 날 D/E 추가 → 오프 10~11에 가둠
  topUpUnderworked(cfg, sched, rng);

  // 3) 남은 칸은 오프
  for (var i3 = 0; i3 < N; i3++)
    for (var day4 = 1; day4 <= nd; day4++)
      if (!sched[i3][day4]) sched[i3][day4] = SHIFT.O;

  // 3-b) 맞교환 보정: 오프 10~11 지키면서 부족한 D/E(특히 E)를 채움 (사람이 손으로 하는 swap)
  repairStaffing(cfg, sched);

  // 4) 연속 오프 최대 제한 (3일 이상 연속 오프 → 가운데를 근무로 전환)
  limitConsecutiveOff(cfg, sched);

  return sched;
}

/* 맞교환 보정: 인원이 부족한 칸을, 근무상한(오프=최소)에 걸린 P의 '다른 날 근무'를
   여유 있는(오프>최소) Q에게 넘기고 P가 그 칸을 채움. → 오프 10~11 유지하면서 빈칸 제거 */
function repairStaffing(cfg, sched) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  var maxWork = nd - cfg.offMin, maxE = cfg.maxEvening || 3;
  for (var iter = 0; iter < nd * N; iter++) {
    var fixed = false;
    for (var d = 1; d <= nd && !fixed; d++) {
      var shifts = ['E', 'D']; // 이브닝 먼저
      for (var si = 0; si < shifts.length && !fixed; si++) {
        var sh = shifts[si];
        if (countShift(sched, d, sh) >= cfg.need[sh]) continue;
        if (sh === 'E' && countShift(sched, d, 'E') >= maxE) continue;
        for (var P = 0; P < N && !fixed; P++) {
          if (sched[P][d] !== 'O' && sched[P][d] !== '') continue;
          if (isLocked(cfg, P, d)) continue; // 요청칸은 안 건드림
          if (fullWorkload(sched, P, nd) < maxWork) continue; // P는 상한(오프 최소)인 사람
          if (!canWork(cfg, sched, P, d, sh)) continue;
          for (var d2 = 1; d2 <= nd && !fixed; d2++) {
            if (d2 === d) continue;
            if (isLocked(cfg, P, d2)) continue; // 요청 근무는 안 옮김
            var sh2 = sched[P][d2];
            if (sh2 !== 'D' && sh2 !== 'E') continue;
            for (var Q = 0; Q < N && !fixed; Q++) {
              if (Q === P) continue;
              if (sched[Q][d2] !== 'O' && sched[Q][d2] !== '') continue;
              if (isLocked(cfg, Q, d2)) continue; // Q의 요청오프도 보호
              if (fullWorkload(sched, Q, nd) >= maxWork) continue; // Q는 여유(오프>최소)
              if (!canWork(cfg, sched, Q, d2, sh2)) continue;
              var save = sched[P][d2];
              sched[P][d2] = 'O';
              if (!canWork(cfg, sched, P, d, sh) || !canWork(cfg, sched, Q, d2, sh2)) { sched[P][d2] = save; continue; }
              sched[P][d] = sh; sched[Q][d2] = sh2; fixed = true;
            }
          }
        }
      }
    }
    if (!fixed) break;
  }
}

/* 연속 오프 한도 — 역할별 (차지 maxConsecOffCharge=3, 액팅 maxConsecOffActing=2).
   한도 초과 시: ① 오프 여유 있으면 근무로 전환  ② 없으면 맞교환으로 오프를 다른 날로 이동 */
function limitConsecutiveOff(cfg, sched) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  for (var i = 0; i < N; i++) {
    var maxC = cfg.nurses[i].charge ? (cfg.maxConsecOffCharge || 3) : (cfg.maxConsecOffActing || 2);
    for (var pass = 0; pass < nd; pass++) {
      var run = 0, fixed = false;
      for (var day = 1; day <= nd; day++) {
        if (sched[i][day] === 'O') {
          run++;
          if (run > maxC && !isLocked(cfg, i, day)) { // 요청오프는 그대로 둠
            // ① 오프가 최소치보다 많으면 그냥 근무로 전환
            if (countOffRow(sched, i, nd) > cfg.offMin) {
              var sh = chooseFillShift(cfg, sched, i, day);
              if (sh) { sched[i][day] = sh; run = 0; fixed = true; continue; }
            }
            // ② 여유 없으면 맞교환으로 오프를 다른 날로 옮김 (오프 개수 유지)
            if (swapOffOut(cfg, sched, i, day)) { run = 0; fixed = true; }
          }
        } else run = 0;
      }
      if (!fixed) break; // 더 고칠 게 없으면 종료
    }
  }
}

/* i가 day(오프)에 근무하고, 대신 i의 다른 근무일을 오프로 — 파트너 Q와 맞교환(오프수·인원수 유지) */
function swapOffOut(cfg, sched, i, day) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  var shifts = ['E', 'D'];
  for (var si = 0; si < shifts.length; si++) {
    var sh = shifts[si];
    for (var Q = 0; Q < N; Q++) {
      if (Q === i) continue;
      if (sched[Q][day] !== sh) continue;        // Q가 day에 sh 근무 → i가 대신
      if (isLocked(cfg, Q, day)) continue;        // Q의 요청근무는 안 옮김
      for (var dW = 1; dW <= nd; dW++) {
        var shW = sched[i][dW];
        if (shW !== 'D' && shW !== 'E') continue; // i의 근무일
        if (isLocked(cfg, i, dW)) continue;        // i의 요청근무는 안 옮김
        if (sched[Q][dW] !== 'O') continue;        // Q는 dW에 오프 → Q가 대신
        if (isLocked(cfg, Q, dW)) continue;        // Q의 요청오프는 보호
        var a1 = sched[i][day], a2 = sched[i][dW], q1 = sched[Q][day], q2 = sched[Q][dW];
        sched[i][day] = sh; sched[i][dW] = 'O';
        sched[Q][day] = 'O'; sched[Q][dW] = shW;
        if (canWork(cfg, sched, i, day, sh) && canWork(cfg, sched, Q, dW, shW)) return true;
        sched[i][day] = a1; sched[i][dW] = a2; sched[Q][day] = q1; sched[Q][dW] = q2; // 롤백
      }
    }
  }
  return false;
}

function countOffRow(sched, i, nd) {
  var c = 0;
  for (var d = 1; d <= nd; d++) if (sched[i][d] === 'O') c++;
  return c;
}

/* (i, day)가 사용자가 요청한 칸(요청오프/요청근무)인지 → 보정에서 건드리지 않도록 */
function isLocked(cfg, i, day) {
  var nu = cfg.nurses[i];
  if (nu.reqOff && nu.reqOff[day]) return true;
  if (nu.reqWork && nu.reqWork[day]) return true;
  if (cfg.preset && cfg.preset[i] && cfg.preset[i][day]) return true; // 표에 직접 입력한 칸
  return false;
}

/* 보정 패스에서 채울 근무 선택: E는 하루 maxEvening(기본 3)까지만, 넘으면 D */
function chooseFillShift(cfg, sched, i, day) {
  var maxE = cfg.maxEvening || 3;
  if (countShift(sched, day, 'E') < maxE && canWork(cfg, sched, i, day, 'E')) return 'E';
  if (canWork(cfg, sched, i, day, 'D')) return 'D';
  return '';
}

/* 보정: ① 부족한 D/E 근무를 여유 인원으로 메꾸고  ② 그래도 일 덜 한 사람은 한가한 날에 추가
   → 빈칸(특히 E) 최소화 + 오프 10~11 유지 */
function topUpUnderworked(cfg, sched, rng) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  var maxWork = nd - cfg.offMin;   // 이 이상 일하면 오프 < 최소 → 금지
  var minWork = nd - cfg.offMax;   // 이만큼은 일해야 오프 ≤ 최대
  var maxE = cfg.maxEvening || 3;

  // ① 부족한 근무(빈칸) 채우기 — 이브닝 우선, 단 오프 최소(상한)는 반드시 지킴
  //    (오프 10~11 보장이 최우선. 상한 때문에 못 채운 칸은 빨강으로 남겨 수기 처리)
  var caps = [true];
  for (var ci = 0; ci < caps.length; ci++) {
    var respectCap = caps[ci];
    for (var pass = 0; pass < nd * 2; pass++) {
      var changed = false;
      for (var day = 1; day <= nd; day++) {
        var shifts = ['E', 'D']; // 이브닝 먼저 채움
        for (var si = 0; si < shifts.length; si++) {
          var sh = shifts[si];
          while (countShift(sched, day, sh) < cfg.need[sh]) {
            if (sh === 'E' && countShift(sched, day, 'E') >= maxE) break;
            var best = -1, bestW = 1e9;
            for (var i = 0; i < N; i++) {
              if (sched[i][day] !== '') continue;
              if (respectCap && fullWorkload(sched, i, nd) >= maxWork) continue; // 1차만 상한 적용
              if (!canWork(cfg, sched, i, day, sh)) continue;
              var w = fullWorkload(sched, i, nd);
              if (w < bestW) { bestW = w; best = i; }
            }
            if (best < 0) break;
            sched[best][day] = sh; changed = true;
          }
        }
      }
      if (!changed) break;
    }
  }

  // ② 그래도 오프가 너무 많은 사람은 한가한 날에 D/E 추가
  for (var pass2 = 0; pass2 < nd * 2; pass2++) {
    var changed2 = false;
    for (var i2 = 0; i2 < N; i2++) {
      if (fullWorkload(sched, i2, nd) >= minWork) continue;
      var bestDay = -1, bestShift = '', bestLoad = 1e9;
      for (var d2 = 1; d2 <= nd; d2++) {
        if (sched[i2][d2] !== '') continue;
        var s2 = chooseFillShift(cfg, sched, i2, d2);
        if (!s2) continue;
        var load = countShift(sched, d2, 'D') + countShift(sched, d2, 'E');
        if (load < bestLoad) { bestLoad = load; bestDay = d2; bestShift = s2; }
      }
      if (bestDay > 0) { sched[i2][bestDay] = bestShift; changed2 = true; }
    }
    if (!changed2) break;
  }
}

/* ── 나이트 배정 디스패처 ──
   요청(요청오프/요청근무)이 없고 하루 2명 체제면 → "정확 구성"(전원 목표치 딱 맞춤)
   요청이 있으면 → 그리디(요청 유연 처리, 나이트 ±1 변동 허용) */
function assignNights(cfg, sched, rng) {
  var hasReq = false;
  for (var i = 0; i < cfg.nurses.length; i++) {
    var nu = cfg.nurses[i];
    if ((nu.reqOff && Object.keys(nu.reqOff).length) || (nu.reqWork && Object.keys(nu.reqWork).length)) { hasReq = true; break; }
  }
  if (cfg.preset) hasReq = true; // 표에 직접 입력한 칸이 있으면 유연 모드(고정 존중)
  if (!hasReq && cfg.need.N === 2 && constructNights(cfg, sched, rng)) return; // 정확 구성 성공
  assignNightsGreedy(cfg, sched, rng); // 폴백
}

/* 나이트 블록 길이별 종료 후 오프 수: 3일 블록 → 2오프, 1~2일 블록 → 1오프 */
function offAfterFor(cfg, blockLen) {
  return (blockLen >= cfg.nightLen) ? (cfg.offAfterNight3 || 2) : (cfg.offAfterNight || 1);
}
/* 나이트 목표 T를 1~3일 블록으로 분해 (2~3 위주, 가끔 1일 섞음) */
function splitNightBlocks(T, rng) {
  if (T === 0) return [];
  var b = [], r = T;
  while (r > 0) {
    var size;
    if (r === 1) size = 1;
    else if (r === 2) size = (rng && rng() < 0.15) ? 1 : 2;      // 2 남으면 가끔 1+1
    else { var x = rng ? rng() : 0.7; size = x < 0.15 ? 1 : (x < 0.5 ? 2 : 3); } // 가끔 1, 보통 2~3
    b.push(size); r -= size;
  }
  return b;
}
function shuffleArr(arr, rng) {
  for (var i = arr.length - 1; i > 0; i--) { var j = Math.floor(rng() * (i + 1)); var t = arr[i]; arr[i] = arr[j]; arr[j] = t; }
  return arr;
}
/* 한 역할(차지 또는 액팅)이 nd일을 1인 1명씩 덮도록 2~3블록 타일링 → owner[1..nd] (실패 시 null) */
function tileNightRole(nd, idxs, rng) {
  var count = idxs.length;
  if (count === 0) return null;
  var base = Math.floor(nd / count), extra = nd % count;
  var order = shuffleArr(idxs.slice(), rng);
  var queues = [];
  for (var k = 0; k < count; k++) {
    var sizes = splitNightBlocks(base + (k < extra ? 1 : 0), rng);
    if (sizes === null) return null;
    queues.push({ nurse: order[k], sizes: sizes });
  }
  var seq = [], last = -1, remain = 0, gi;
  for (gi = 0; gi < queues.length; gi++) remain += queues[gi].sizes.length;
  var guard = 0;
  while (remain > 0 && guard++ < 100000) {
    queues.sort(function (a, b) { return b.sizes.length - a.sizes.length; });
    var pick = null, qi;
    for (qi = 0; qi < queues.length; qi++) if (queues[qi].sizes.length > 0 && queues[qi].nurse !== last) { pick = queues[qi]; break; }
    if (!pick) for (qi = 0; qi < queues.length; qi++) if (queues[qi].sizes.length > 0) { pick = queues[qi]; break; }
    seq.push({ nurse: pick.nurse, size: pick.sizes.shift() }); last = pick.nurse; remain--;
  }
  for (var si = 1; si < seq.length; si++) if (seq[si].nurse === seq[si - 1].nurse) return null; // 같은 너스 연속 = 실패
  var owner = [], day = 1;
  for (var bi = 0; bi < seq.length; bi++) for (var t = 0; t < seq[bi].size; t++) { owner[day] = seq[bi].nurse; day++; }
  return (day - 1 === nd) ? owner : null;
}
/* 정확 구성: 차지/액팅 각각 타일링 → 매 밤 차지1+액팅1, 전원 정확히 목표 나이트. 전후 오프 부여 */
function constructNights(cfg, sched, rng) {
  var nd = cfg.numDays, N = cfg.nurses.length, charges = [], actings = [];
  for (var i = 0; i < N; i++) cfg.nurses[i].charge ? charges.push(i) : actings.push(i);
  if (!charges.length || !actings.length) return false;
  var co = tileNightRole(nd, charges, rng), ao = tileNightRole(nd, actings, rng);
  if (!co || !ao) return false;
  for (var d = 1; d <= nd; d++) { sched[co[d]][d] = SHIFT.N; sched[ao[d]][d] = SHIFT.N; }
  // 블록 전후 오프 (3연속만 앞 오프)
  for (var p = 0; p < N; p++) {
    for (var day = 1; day <= nd; day++) {
      if (sched[p][day] === 'N' && sched[p][day - 1] !== 'N') {
        var len = 0; while (sched[p][day + len] === 'N') len++;
        var end = day + len - 1;
        var obn = (len >= cfg.nightLen) ? cfg.offBeforeNight : 0;
        for (var b = 1; b <= obn; b++) { var bd = day - b; if (bd >= 1 && !sched[p][bd]) sched[p][bd] = SHIFT.O; }
        var oan = offAfterFor(cfg, len);
        for (var a = 1; a <= oan; a++) { var ad = end + a; if (ad <= nd && !sched[p][ad]) sched[p][ad] = SHIFT.O; }
      }
    }
  }
  return true;
}

/* 나이트 블록(그리디): 2~maxLen일 윈도우 타일링 + 차지 커버리지 (요청 있을 때 폴백) */
function assignNightsGreedy(cfg, sched, rng) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  var maxLen = cfg.nightLen || 3;             // 블록 최대 길이
  var minLen = maxLen >= 2 ? 2 : maxLen;      // 블록 최소 길이(2 허용)
  var perNurse = cfg.maxNightBlocks || 2;     // 윈도우 개수 계산용

  // ── 역할별 나이트 목표 계산 (액팅 6 고정, 차지 = 나머지) ──
  var chargeCount = 0, actingCount = 0;
  for (var ri = 0; ri < N; ri++) { if (cfg.nurses[ri].charge) chargeCount++; else actingCount++; }
  var totalSlots = cfg.need.N * nd;                 // 한 달 전체 나이트 자리
  var maxActingTotal = (cfg.need.N - 1) * nd;       // 매 밤 차지 1명 확보 후 액팅이 가질 수 있는 최대
  var idealActing = actingCount * (cfg.actingNightTarget || 6); // 액팅 1인 목표(기본 6)
  // 액팅이 최대한 가져가고(차지 최소화) → 안 나눠떨어지는 자투리(31일 등)는 액팅 쪽으로
  var actingTotal = Math.min(idealActing, maxActingTotal);
  var chargeTotal = totalSlots - actingTotal;       // 차지는 커버리지 최소만
  var actingTarget = actingCount > 0 ? actingTotal / actingCount : 0;
  var chargeTarget = chargeCount > 0 ? chargeTotal / chargeCount : 0;
  var nightTarget = [];
  for (var ti = 0; ti < N; ti++) nightTarget[ti] = cfg.nurses[ti].charge ? chargeTarget : actingTarget;

  // 윈도우 개수 W (블록 크기 결정) — 분배 방식 기준
  var Wmin = Math.ceil(nd / maxLen);
  var Wmax = Math.max(Wmin, Math.round(N * perNurse / cfg.need.N));
  var mode = cfg.nightMode || 2;
  var W = (mode <= 1) ? Wmin : (mode >= 3) ? Wmax : Math.round((Wmin + Wmax) / 2);
  W = Math.max(Math.ceil(nd / maxLen), Math.min(W, Math.floor(nd / minLen)));

  // 윈도우 크기 목록: 모두 minLen에서 시작 → 남는 일수만큼 maxLen 쪽으로 키움
  var sizes = []; for (var w = 0; w < W; w++) sizes.push(minLen);
  var remainder = nd - minLen * W;
  var step = Math.max(1, maxLen - minLen);
  var wi = 0;
  while (remainder >= step && wi < W) { sizes[wi] += step; remainder -= step; wi++; }
  if (remainder > 0) sizes[W - 1] += remainder; // 자투리는 마지막에 흡수(월 길이 보정)

  // 큰 블록/작은 블록이 번갈아 나오도록 재배치 (3,2,3,2 …)
  sizes.sort(function (a, b) { return b - a; });
  var mixed = [], lo = 0, hi = sizes.length - 1, take = true;
  while (lo <= hi) { take ? mixed.push(sizes[lo++]) : mixed.push(sizes[hi--]); take = !take; }
  sizes = mixed;

  var thisNights = [], sizeCount = [];
  for (var i = 0; i < N; i++) { thisNights[i] = 0; sizeCount[i] = {}; }

  var start = 1;
  for (var s = 0; s < sizes.length && start <= nd; s++) {
    var end = Math.min(start + sizes[s] - 1, nd);
    var winSize = end - start + 1;
    var assigned = 0;
    // 1) 차지 1명 우선 (모든 나이트 차지 커버). 목표 지키며 → 안 되면 초과 허용 폴백
    var c = pickWindowNurse(cfg, sched, start, end, winSize, thisNights, sizeCount, nightTarget, true, true, rng);
    if (c < 0) c = pickWindowNurse(cfg, sched, start, end, winSize, thisNights, sizeCount, nightTarget, true, false, rng);
    if (c >= 0) { placeNight(cfg, sched, c, start, end); thisNights[c] += winSize; sizeCount[c][winSize] = (sizeCount[c][winSize] || 0) + 1; assigned++; }
    // 2) 나머지 인원 (액팅 우선 → 액팅이 목표까지 흡수, 차지는 나머지만)
    while (assigned < cfg.need.N) {
      c = pickWindowNurse(cfg, sched, start, end, winSize, thisNights, sizeCount, nightTarget, false, true, rng);
      if (c < 0) c = pickWindowNurse(cfg, sched, start, end, winSize, thisNights, sizeCount, nightTarget, false, false, rng);
      if (c < 0) break;
      placeNight(cfg, sched, c, start, end); thisNights[c] += winSize; sizeCount[c][winSize] = (sizeCount[c][winSize] || 0) + 1; assigned++;
    }
    start = end + 1;
  }
}

/* 나이트 블록 1개 배정 + 전후 오프 */
function placeNight(cfg, sched, i, start, end) {
  for (var day = start; day <= end; day++) sched[i][day] = SHIFT.N;
  // 앞 오프: 3연속(=nightLen) 블록일 때만 적용. 짧은 블록은 앞에 근무 가능
  var obn = (end - start + 1 >= cfg.nightLen) ? cfg.offBeforeNight : 0;
  for (var b = 1; b <= obn; b++) {
    var bd = start - b;
    if (bd >= 1 && !sched[i][bd]) sched[i][bd] = SHIFT.O;
  }
  var oan = offAfterFor(cfg, end - start + 1);
  for (var a = 1; a <= oan; a++) {
    var ad = end + a;
    if (ad <= cfg.numDays && !sched[i][ad]) sched[i][ad] = SHIFT.O;
  }
}

/* 윈도우 [start..end]에 나이트 가능한 간호사 1명 선택 (목표 기반)
   - nightTarget[i]: 그 사람의 목표 나이트 수 (액팅 6 / 차지 나머지)
   - respectTarget: true면 목표 초과자 제외(폴백 시 false로 재호출)
   - 우선순위: 목표 대비 부족분 큰 사람 → (2번째 자리)액팅 우선 → 전월 적은 사람 */
function pickWindowNurse(cfg, sched, start, end, winSize, thisNights, sizeCount, nightTarget, requireCharge, respectTarget, rng) {
  var N = cfg.nurses.length, pool = [];
  for (var i = 0; i < N; i++) {
    if (requireCharge && !cfg.nurses[i].charge) continue;
    // 목표 초과 방지 (폴백 땐 무시)
    if (respectTarget && thisNights[i] + winSize > Math.ceil(nightTarget[i])) continue;
    var ok = true;
    // 블록 구간이 비어 있어야
    for (var day = start; day <= end; day++) { if (sched[i][day] !== '') { ok = false; break; } }
    if (!ok) continue;
    // 시작 전 오프 자리가 근무면 불가 (3연속 블록일 때만 앞 오프 필요)
    var obn = (end - start + 1 >= cfg.nightLen) ? cfg.offBeforeNight : 0;
    for (var b = 1; b <= obn; b++) {
      var bd = start - b;
      if (bd >= 1) { var pv = sched[i][bd]; if (pv === 'D' || pv === 'E' || pv === 'N') { ok = false; break; } }
    }
    if (!ok) continue;
    // 종료 후 오프 자리가 근무면 불가
    var oanW = offAfterFor(cfg, end - start + 1);
    for (var a = 1; a <= oanW; a++) {
      var ad = end + a;
      if (ad <= cfg.numDays) { var nv = sched[i][ad]; if (nv === 'D' || nv === 'E' || nv === 'N') { ok = false; break; } }
    }
    if (ok) pool.push(i);
  }
  if (!pool.length) return -1;
  pool.sort(function (a, b) {
    // 1) 목표 대비 부족분(deficit) 큰 사람 먼저 → 각자 목표치까지 채움
    var defa = nightTarget[a] - thisNights[a], defb = nightTarget[b] - thisNights[b];
    if (defa !== defb) return defb - defa;
    // 1-b) 같은 길이 블록을 두 번 안 하도록 → 2일+3일=5로 딱 떨어지게 (4·6 변동 방지)
    var sa = (sizeCount[a][winSize] || 0), sb = (sizeCount[b][winSize] || 0);
    if (sa !== sb) return sa - sb;
    // 2) 커버리지 외(2번째) 자리는 액팅 우선 → 액팅이 목표까지 흡수, 차지는 나머지만
    if (!requireCharge) {
      var ra = cfg.nurses[a].charge ? 1 : 0, rb = cfg.nurses[b].charge ? 1 : 0;
      if (ra !== rb) return ra - rb;
    }
    // 3) 전월 나이트 적은 사람 → 동률 랜덤 (캐리오버 공평)
    var na = cfg.nurses[a].prevNightDays || 0, nb = cfg.nurses[b].prevNightDays || 0;
    if (na !== nb) return na - nb;
    return rng() - 0.5;
  });
  return pool[0];
}

/* 하루의 Day+Evening 채우기 — 각 근무에 차지 1명씩 먼저 확보(차지 구하기 빡센 근무부터) */
function fillDayEvening(cfg, sched, day, rng) {
  // 1) 차지가 없는 D/E를 찾아, 차지 후보가 적은(빡센) 근무부터 차지 1명 확보
  var needCharge = [];
  ['D', 'E'].forEach(function (sh) {
    if (countShift(sched, day, sh) < cfg.need[sh] && !shiftHasCharge(cfg, sched, day, sh))
      needCharge.push(sh);
  });
  needCharge.sort(function (a, b) {
    return countEligible(cfg, sched, day, a, true) - countEligible(cfg, sched, day, b, true);
  });
  needCharge.forEach(function (sh) {
    var c = pickDayCandidate(cfg, sched, day, sh, true, rng);
    if (c >= 0) sched[c][day] = sh;
  });
  // 2) 남은 슬롯 채우기 (역할 무관)
  ['D', 'E'].forEach(function (sh) {
    while (countShift(sched, day, sh) < cfg.need[sh]) {
      var c = pickDayCandidate(cfg, sched, day, sh, false, rng);
      if (c < 0) break;
      sched[c][day] = sh;
    }
  });
}

/* day에 shift 근무 가능한 후보 수 (chargeOnly=true면 차지만) */
function countEligible(cfg, sched, day, shift, chargeOnly) {
  var n = 0;
  for (var i = 0; i < cfg.nurses.length; i++) {
    if (sched[i][day] !== '') continue;
    if (chargeOnly && !cfg.nurses[i].charge) continue;
    if (canWork(cfg, sched, i, day, shift)) n++;
  }
  return n;
}

/* Day 또는 Evening 채우기 (단일 근무) */
function fillDayShift(cfg, sched, day, shift, rng) {
  var need = cfg.need[shift];
  // 이미 (요청근무 등으로) 배정된 인원 고려
  var have = countShift(sched, day, shift);
  if (have >= need) return;

  // 차지 1명 우선 확보
  var hasCharge = shiftHasCharge(cfg, sched, day, shift);
  while (countShift(sched, day, shift) < need) {
    var cand = pickDayCandidate(cfg, sched, day, shift, !hasCharge, rng);
    if (cand < 0) {
      // 차지 강제 실패 시 아무나라도
      cand = pickDayCandidate(cfg, sched, day, shift, false, rng);
      if (cand < 0) break;
    }
    sched[cand][day] = shift;
    if (cfg.nurses[cand].charge) hasCharge = true;
  }
}

function pickDayCandidate(cfg, sched, day, shift, requireCharge, rng) {
  var N = cfg.nurses.length;
  var pool = [];
  var maxWork = cfg.numDays - cfg.offMin; // 오프 최소치 보장 → 이만큼 일하면 더 안 줌
  for (var i = 0; i < N; i++) {
    if (sched[i][day] !== '') continue;             // 이미 배정/오프
    if (requireCharge && !cfg.nurses[i].charge) continue;
    if (fullWorkload(sched, i, cfg.numDays) >= maxWork) continue; // 근무 상한(오프<최소 방지)
    if (!canWork(cfg, sched, i, day, shift)) continue;
    pool.push(i);
  }
  if (!pool.length) return -1;
  // 전체(한 달) 근무량이 적은 사람 우선 → 총 근무가 고르게 = 오프 개수가 고르게
  pool.sort(function (a, b) {
    var diff = fullWorkload(sched, a, cfg.numDays) - fullWorkload(sched, b, cfg.numDays);
    return diff !== 0 ? diff : (rng() - 0.5);
  });
  return pool[0];
}

/* 한 달 전체 근무(D/E/N) 일수 — 나이트는 이미 배정돼 있으므로 실시간 총 근무량 */
function fullWorkload(sched, i, nd) {
  var c = 0;
  for (var d = 1; d <= nd; d++) {
    var v = sched[i][d];
    if (v === 'D' || v === 'E' || v === 'N') c++;
  }
  return c;
}

/* 해당 칸에 shift 근무 가능한지 (하드 규칙) */
function canWork(cfg, sched, i, day, shift) {
  var prev = day - 1 >= 1 ? sched[i][day - 1] : '';
  // N 다음날 D/E 금지
  if (prev === 'N') return false;
  // E 다음날 D 금지
  if (shift === 'D' && prev === 'E') return false;
  // 연속근무 최대 제한 (앞뒤 양방향 — 나이트 전에 근무가 붙는 경우까지 차단)
  var back = 0;
  for (var d = day - 1; d >= 1; d--) {
    var v = sched[i][d];
    if (v === 'D' || v === 'E' || v === 'N') back++; else break;
  }
  var fwd = 0;
  for (var f = day + 1; f <= cfg.numDays; f++) {
    var vf = sched[i][f];
    if (vf === 'D' || vf === 'E' || vf === 'N') fwd++; else break;
  }
  if (back + 1 + fwd > cfg.maxConsec) return false;
  return true;
}

/* ===================== 유틸 ===================== */
function countShift(sched, day, shift) {
  var c = 0;
  for (var i = 0; i < sched.length; i++) if (sched[i][day] === shift) c++;
  return c;
}
function shiftHasCharge(cfg, sched, day, shift) {
  for (var i = 0; i < sched.length; i++)
    if (sched[i][day] === shift && cfg.nurses[i].charge) return true;
  return false;
}
function workloadUpTo(sched, i, day) {
  var c = 0;
  for (var d = 1; d < day; d++) {
    var v = sched[i][d];
    if (v === 'D' || v === 'E' || v === 'N') c++;
  }
  return c;
}
function makeRng(seed) {
  var s = seed % 2147483647;
  if (s <= 0) s += 2147483646;
  return function () { s = (s * 16807) % 2147483647; return s / 2147483647; };
}

/* 점수 (낮을수록 좋음) */
function evaluate(cfg, sched) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  var unfilled = 0, hard = 0, offDev = 0;

  for (var day = 1; day <= nd; day++) {
    WORK_SHIFTS.forEach(function (sh) {
      var diff = cfg.need[sh] - countShift(sched, day, sh);
      if (diff > 0) unfilled += diff;
      // 차지 없는 근무
      if (countShift(sched, day, sh) > 0 && !shiftHasCharge(cfg, sched, day, sh)) hard++;
    });
  }
  var nightDev = 0;
  var avgNight = (cfg.need.N * nd) / N; // 1인 평균 나이트 일수
  var maxChargeN = 0, minActingN = 1e9, hasActing = false;
  for (var i = 0; i < N; i++) {
    var off = 0, nights = 0;
    for (var d = 1; d <= nd; d++) {
      if (sched[i][d] === 'O' || sched[i][d] === '') off++;
      else if (sched[i][d] === 'N') nights++;
    }
    if (off < cfg.offMin) offDev += (cfg.offMin - off);
    if (off > cfg.offMax) offDev += (off - cfg.offMax);
    nightDev += Math.abs(nights - avgNight); // 나이트 균등도 (평균에서 벗어난 정도)
    if (cfg.nurses[i].charge) { if (nights > maxChargeN) maxChargeN = nights; }
    else { hasActing = true; if (nights < minActingN) minActingN = nights; }
  }
  // 액팅이 차지보다 나이트가 적으면(차지 최다 > 액팅 최소) 벌점 → 재굴리기 시 회피
  var roleViol = hasActing ? Math.max(0, maxChargeN - minActingN) : 0;
  return {
    unfilled: unfilled, hard: hard, offDev: offDev, nightDev: Math.round(nightDev), roleViol: roleViol,
    total: unfilled * 10 + hard * 8 + offDev * 2 + nightDev * 6 + roleViol * 10
  };
}

/* ===================== 시트에 기록 ===================== */
function writeSchedule(cfg, sched) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d = ss.getSheetByName(DUTY_SHEET);
  var nd = cfg.numDays, N = cfg.nurses.length;
  var out = [];
  for (var i = 0; i < N; i++) {
    var row = [];
    for (var day = 1; day <= nd; day++) row.push(sched[i][day] || 'O');
    out.push(row);
  }
  d.getRange(DUTY_DATA_START_ROW, 2, N, nd).setValues(out);
}

/* ===================== 규칙 검사 ===================== */
function checkRules() {
  var cfg = readSettings();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d = ss.getSheetByName(DUTY_SHEET);
  if (!d) { SpreadsheetApp.getUi().alert('듀티표 시트가 없습니다.'); return; }

  var nd = cfg.numDays, N = cfg.nurses.length;
  var c0 = 2;
  var data = d.getRange(DUTY_DATA_START_ROW, c0, N, nd).getValues();

  // 위반 배경색 초기화 (조건부서식과 별개의 실배경)
  var rng = d.getRange(DUTY_DATA_START_ROW, c0, N, nd);
  rng.setBackground(null);

  var msgs = [];
  var bg = [];
  for (var i = 0; i < N; i++) { bg[i] = []; for (var j = 0; j < nd; j++) bg[i][j] = null; }

  function get(i, day) { return (day >= 1 && day <= nd) ? data[i][day - 1] : ''; }
  function flag(i, day, why) {
    if (day >= 1 && day <= nd) bg[i][day - 1] = COLORS.VIOLATION;
    msgs.push(cfg.nurses[i].name + ' / ' + day + '일: ' + why);
  }

  for (var ii = 0; ii < N; ii++) {
    // 연속근무 / 금지패턴 / 나이트 블록 검사
    var streak = 0;
    for (var day = 1; day <= nd; day++) {
      var cur = get(ii, day), prev = get(ii, day - 1);
      var isWork = (cur === 'D' || cur === 'E' || cur === 'N');
      streak = isWork ? streak + 1 : 0;
      if (streak > cfg.maxConsec) flag(ii, day, '연속근무 ' + streak + '일 (최대 ' + cfg.maxConsec + ')');
      if (cur === 'D' && prev === 'N') flag(ii, day, 'N 다음날 D 금지');
      if (cur === 'E' && prev === 'N') flag(ii, day, 'N 다음날 E 금지');
      if (cur === 'D' && prev === 'E') flag(ii, day, 'E 다음날 D 금지');
    }
    // 나이트 블록 길이 & 전후 오프
    for (var day2 = 1; day2 <= nd; day2++) {
      if (get(ii, day2) === 'N' && get(ii, day2 - 1) !== 'N') {
        // 블록 시작
        var len = 0; while (get(ii, day2 + len) === 'N') len++;
        var endDay = day2 + len - 1;
        var nMax = cfg.nightLen; // 1~maxLen 허용 (1일 나이트 OK)
        if (len > nMax && endDay < nd) // 월말 잘림은 예외
          flag(ii, day2, '나이트 연속 ' + len + '일 (최대 ' + nMax + ')');
        // 시작 전 오프 (3연속 블록일 때만)
        var obn2 = (len >= cfg.nightLen) ? cfg.offBeforeNight : 0;
        for (var b = 1; b <= obn2; b++) {
          if (day2 - b >= 1 && get(ii, day2 - b) !== 'O')
            flag(ii, day2 - b, '3연속 나이트 시작 전 ' + obn2 + '오프 필요');
        }
        // 종료 후 오프 (3일 블록 2오프 / 1~2일 블록 1오프)
        var oanC = offAfterFor(cfg, len);
        for (var a = 1; a <= oanC; a++) {
          if (endDay + a <= nd && get(ii, endDay + a) !== 'O')
            flag(ii, endDay + a, '나이트(' + len + '일) 후 ' + oanC + '오프 필요');
        }
      }
    }
    // 월 오프 수
    var off = 0;
    for (var day3 = 1; day3 <= nd; day3++) if (get(ii, day3) === 'O') off++;
    if (off < cfg.offMin || off > cfg.offMax)
      msgs.push('⚠ ' + cfg.nurses[ii].name + ': 오프 ' + off + '개 (목표 ' + cfg.offMin + '~' + cfg.offMax + ')');
    // 연속 오프 최대 (차지 3 / 액팅 2)
    var maxC = cfg.nurses[ii].charge ? (cfg.maxConsecOffCharge || 3) : (cfg.maxConsecOffActing || 2);
    var orun = 0;
    for (var day5 = 1; day5 <= nd; day5++) {
      if (get(ii, day5) === 'O') {
        orun++;
        if (orun > maxC) flag(ii, day5, '오프 ' + orun + '일 연속 (' + (cfg.nurses[ii].charge ? '차지' : '액팅') + ' 최대 ' + maxC + ')');
      } else orun = 0;
    }
  }

  // 일별 인원 / 차지 검사
  for (var day4 = 1; day4 <= nd; day4++) {
    WORK_SHIFTS.forEach(function (sh) {
      var cnt = 0, charge = 0;
      for (var i = 0; i < N; i++) {
        if (data[i][day4 - 1] === sh) { cnt++; if (cfg.nurses[i].charge) charge++; }
      }
      if (cnt !== cfg.need[sh])
        msgs.push('⚠ ' + day4 + '일 ' + sh + ' 인원 ' + cnt + '명 (필요 ' + cfg.need[sh] + ')');
      if (cnt > 0 && charge === 0)
        msgs.push('⚠ ' + day4 + '일 ' + sh + ': 차지 없음 (액팅만 근무)');
    });
  }

  // 배경색 일괄 적용
  rng.setBackgrounds(bg);

  if (msgs.length === 0) {
    SpreadsheetApp.getUi().alert('✅ 규칙 위반이 없습니다!');
  } else {
    var head = '발견된 항목 ' + msgs.length + '건 (빨간 칸 = 하드규칙 위반):\n\n';
    SpreadsheetApp.getUi().alert(head + msgs.slice(0, 60).join('\n') +
      (msgs.length > 60 ? '\n... 외 ' + (msgs.length - 60) + '건' : ''));
  }
}

/* ===================== 내용 비우기 ===================== */
function clearDutyValues() {
  var cfg = readSettings();
  var d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DUTY_SHEET);
  if (!d) return;
  d.getRange(DUTY_DATA_START_ROW, 2, cfg.nurses.length, cfg.numDays)
    .clearContent().setBackground(null);
}
