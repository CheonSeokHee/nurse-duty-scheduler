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
 *  - 선호 듀티(소프트): 간호사별로 Day/Evening/Night 선호 + 강도(약간/보통/강하게)를
 *    넣으면 되도록 그쪽으로 배정 (하드 규칙 아님 — 하루 필요인원 맞추는 선에서 최대 반영)
 *  - 첫 요청오프 전날 = 가능하면 Day로 고정 (그날 D 초과/하드규칙 위반이면 건너뜀)
 *  - 첫 요청오프 다음날 = 무조건 나이트 (D-O-N 패턴; 그 칸에 직접 입력이 있으면 입력 존중)
 */

/* ===================== 상수 ===================== */
var SETTINGS_SHEET = '설정';
var DUTY_SHEET = '듀티표';

var SHIFT = { D: 'D', E: 'E', N: 'N', O: 'O' };
var WORK_SHIFTS = ['D', 'E', 'N'];

// 선호 듀티 반영(소프트). 커버리지·오프를 먼저 지키는 선에서 선호 쪽으로 몰아준다.
// 강도(약간/보통/강하게)는 간호사별로 설정 시트에서 지정 → 아래 맵으로 환산.
var PREF_DAY_SORT = true;                              // Day/Evening 채울 때 선호자 살짝 우선
var NIGHT_STRENGTH_MAP = { 1: 1.25, 2: 1.5, 3: 1.9 };  // 약간/보통/강하게 → 나이트 목표 배율
var PREFMISS_WEIGHT_MAP = { 1: 2, 2: 3, 3: 5 };        // 약간/보통/강하게 → 선호 반영 점수 가중
// 선호 근무가 전체 근무일에서 차지하는 "목표 비율" — 100%로 극단화되지 않게 상한 역할.
// 강하게여도 ~75%까지만 선호 근무, 나머지는 다른 듀티도 섞임 (예: 20근무 → 15일 선호 + 5일 기타)
var PREF_TARGET_RATIO = { 1: 0.55, 2: 0.65, 3: 0.75 };

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
    .addItem('수기 입력 잠금 해제(전체 새로 짜기)', 'resetPresetLock')
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
  var headers = [['이름', '역할', '요청오프(예: 3,10,21)', '요청근무(예: D:5 / N:12)', '전월 나이트(일)', '선호 듀티', '강도']];
  s.getRange(NURSE_START_ROW, 1, 1, 7).setValues(headers)
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

  // 선호 듀티 드롭다운 (소프트 — 되도록 반영)
  var prefRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['상관없음', 'Day', 'Evening', 'Night'], true).build();
  s.getRange(NURSE_START_ROW + 1, 6, 50, 1).setDataValidation(prefRule);

  // 강도 드롭다운 (선호를 얼마나 세게 반영할지 — 비우면 '보통')
  var strRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['약간', '보통', '강하게'], true).build();
  s.getRange(NURSE_START_ROW + 1, 7, 50, 1).setDataValidation(strRule);

  s.setColumnWidth(1, 130);
  s.setColumnWidth(2, 80);
  s.setColumnWidth(3, 180);
  s.setColumnWidth(4, 200);
  s.setColumnWidth(5, 110);
  s.setColumnWidth(6, 100);
  s.setColumnWidth(7, 70);

  // --- 듀티표 시트 ---
  var d = ss.getSheetByName(DUTY_SHEET) || ss.insertSheet(DUTY_SHEET);
  drawDutyTemplate();

  ss.setActiveSheet(s);
  SpreadsheetApp.getUi().alert(
    '세팅 완료!\n\n' +
    '1) [설정] 시트에서 연/월, 인원, 간호사 이름·역할을 입력하세요.\n' +
    '   - "선호 듀티" 열에서 사람마다 Day/Evening/Night 선호를 고르면\n' +
    '     자동배정이 되도록 그쪽으로 몰아줍니다(필요인원 맞추는 선에서).\n' +
    '   - "강도"(약간/보통/강하게)로 얼마나 세게 몰아줄지 조절(비우면 보통).\n' +
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
  var vals = s.getRange(NURSE_START_ROW + 1, 1, last - NURSE_START_ROW, 7).getValues();
  for (var i = 0; i < vals.length; i++) {
    var name = (vals[i][0] || '').toString().trim();
    if (!name) continue;
    var role = (vals[i][1] || '액팅').toString().trim();
    var prevNightDays = Number(vals[i][4]) || 0;
    var reqOff = parseReqOff(vals[i][2]);
    nurses.push({
      name: name,
      charge: role === '차지',
      reqOff: reqOff,
      reqWork: parseReqWork(vals[i][3]),
      prevNightDays: prevNightDays,
      prevBlocks: Math.round(prevNightDays / (cfg.nightLen || 3)),
      prefShift: parsePref(vals[i][5]), // 선호 듀티 (소프트): '' / 'D' / 'E' / 'N'
      prefStrength: parsePrefStrength(vals[i][6]) // 강도: 1(약간)/2(보통)/3(강하게)
      // 첫 요청오프 전날 Day 고정은 tryBuild 0-b에서 요청오프+수기입력(O)을 합쳐 계산
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

/* {3:true, 10:true, ...} 같은 날짜맵에서 가장 빠른 날 반환 (없으면 0) */
function minDayKey(obj) {
  var min = 0;
  for (var k in obj) {
    var n = parseInt(k, 10);
    if (n >= 1 && (min === 0 || n < min)) min = n;
  }
  return min;
}

/* 선호 듀티 콤보값 -> 'D' / 'E' / 'N' / ''(상관없음) */
function parsePref(v) {
  var t = ('' + (v || '')).trim().toUpperCase();
  if (t === 'DAY' || t === 'D') return 'D';
  if (t === 'EVENING' || t === 'E') return 'E';
  if (t === 'NIGHT' || t === 'N') return 'N';
  return ''; // 상관없음 / 빈칸
}

/* 선호 우선순위(작을수록 먼저 뽑힘) — Day/Evening 칸 채울 때.
   '약간'(강도1)은 선택 단계에서 밀어주지 않음(점수 prefMiss로만 살짝 반영) → 진짜 약한 쏠림.
   '보통/강하게'(강도2~3)는 선택 단계에서도 우선 → 강한 쏠림. */
function dayPrefRank(nurse, shift) {
  if (!nurse.prefShift || (nurse.prefStrength || 2) < 2) return 1; // 상관없음 / 약간 → 중립
  if (nurse.prefShift === shift) return 0;   // 이 근무를 선호 → 우선
  return 2;                                  // 다른 근무 선호 → 후순위
}

/* 강도 콤보값 -> 1(약간) / 2(보통) / 3(강하게). 빈칸/기타는 보통(2) */
function parsePrefStrength(v) {
  var t = ('' + (v || '')).trim();
  if (t.charAt(0) === '약') return 1; // 약간
  if (t.charAt(0) === '강') return 3; // 강하게
  return 2;                           // 보통 / 빈칸
}
/* 나이트 목표 배율
   - N선호: 강도만큼 나이트 ↑ (약간/보통/강하게 = 1.25/1.5/1.9배)
   - D/E선호: 나이트는 "강도와 무관하게" 살짝만 줄임(고정 1/1.4). 강도는 D/E 비중에만 반영.
     (강도를 나이트 회피에도 걸면 나이트 배치가 바뀌어 D/E 비중이 되레 줄 수 있어 분리함)
   - 상관없음: 1 */
function nightWeightOf(nurse) {
  if (!nurse.prefShift) return 1;
  if (nurse.prefShift === 'N') return NIGHT_STRENGTH_MAP[nurse.prefStrength] || 1.5;
  return 1 / 1.4; // D/E 선호자는 나이트 약간만 회피(고정)
}
/* 선호 반영 점수(prefMiss) 1인당 가중 — 강할수록 best 선택에서 선호를 더 챙김 */
function prefMissWeightOf(nurse) {
  return PREFMISS_WEIGHT_MAP[nurse.prefStrength] || 3;
}

/* 선호 근무 목표 일수 = 예상 근무일(nd-offMax) × 강도별 목표 비율
   → 이 이상은 밀어붙이지 않음 (강하게여도 다른 듀티를 섞기 위한 상한) */
function prefTargetCount(cfg, nurse) {
  var ratio = PREF_TARGET_RATIO[nurse.prefStrength] || 0.65;
  return Math.round((cfg.numDays - cfg.offMax) * ratio);
}

/* i가 선호 근무(D/E)를 이미 목표만큼 받았는지 → 도달했으면 선택에서 보통 사람 취급 */
function prefSatisfied(cfg, sched, i) {
  var nu = cfg.nurses[i];
  if (nu.prefShift !== 'D' && nu.prefShift !== 'E') return false;
  var c = 0;
  for (var d = 1; d <= cfg.numDays; d++) if (sched[i][d] === nu.prefShift) c++;
  return c >= prefTargetCount(cfg, nu);
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

  // ── 표에 직접 입력해둔 칸 읽기 + 영속 잠금 ──
  // 수기 입력 칸은 문서 속성에 저장돼 재실행해도 계속 고정된다.
  //  · 첫 실행(직전 생성본 없음): 부분 입력이면 채워진 칸 전부 고정 (기존 동작)
  //  · 재실행: 저장된 잠금 + (직전 생성본과 달라진 칸 = 새 수기 입력)을 고정
  //  · 잠긴 칸을 지우고 실행하면 그 칸은 잠금 해제
  var preset = null, presetCount = 0;
  var ss0 = SpreadsheetApp.getActiveSpreadsheet();
  var d0 = ss0.getSheetByName(DUTY_SHEET);
  if (d0) {
    try {
      var cur = d0.getRange(DUTY_DATA_START_ROW, 2, cfg.nurses.length, cfg.numDays).getValues();
      var state = loadPresetState(cfg); // {preset, lastGen}
      preset = computePreset(cur, state.preset, state.lastGen, cfg.nurses.length, cfg.numDays);
      presetCount = countPresetCells(preset);
      savePresetState(cfg, preset); // 다음 실행을 위해 잠금 저장(없으면 삭제)
    } catch (e) { preset = null; }
  }
  cfg.preset = preset; // tryBuild / isLocked 가 참고

  drawDutyTemplate();

  // 실행할 때마다 다른 배치가 나오도록 랜덤 베이스 시드 (재굴리기 가능)
  var base = Math.floor(Math.random() * 2000000000);
  var best = null;
  for (var a = 0; a < cfg.attempts; a++) {
    // 포트폴리오 탐색: 절반은 선호 넛지 ON(선호 잘 반영), 절반은 OFF(커버리지 우선)
    // → 점수(evaluate)가 둘 중 "빈칸 없고 선호도 챙긴" 최선을 고름
    cfg.prefNudge = (a % 2 === 0);
    var rng = makeRng(base + a * 2654435761 + 12345);
    var sched = tryBuild(cfg, rng);
    var score = evaluate(cfg, sched);
    if (!best || score.total < best.score.total) best = { sched: sched, score: score };
    if (score.total === 0) break;
  }

  writeSchedule(cfg, best.sched);
  saveLastGenerated(cfg, best.sched); // 다음 실행 때 "수기로 고친 칸"을 구분하는 기준
  checkRules(); // 생성 직후 자동 검사

  var sc = best.score;
  var presetMsg = preset ? ('입력 ' + presetCount + '칸 고정 + ') : '';
  SpreadsheetApp.getActiveSpreadsheet().toast(
    presetMsg + '자동 배정 완료 (미충족 ' + sc.unfilled + ' / 오프편차 ' + sc.offDev +
    ' / 나이트편차 ' + sc.nightDev + ' / 위반 ' + sc.hard + ')',
    '듀티표', 6);
}

/* ===================== 수기 입력 영속 잠금 ===================== */
/* 수기 입력(고정) 칸 계산 — 순수 함수.
   cur: 표의 현재 값(2D), savedPreset: 저장된 잠금({i:{day:'D'}}), lastGen: 직전 생성본(행 문자열 배열)
   · lastGen 있음(재실행): 저장된 잠금 칸 유지 + 직전 생성본과 달라진 칸을 새 잠금으로 추가.
     잠긴 칸이 지워져 있으면 잠금 해제.
   · lastGen 없음(첫 실행): 부분 입력일 때만 채워진 칸 전부 잠금 (꽉 참/빈 표 = 새 배정) */
function computePreset(cur, savedPreset, lastGen, N, nd) {
  function norm(v) {
    v = ('' + (v == null ? '' : v)).trim().toUpperCase();
    return (v === 'D' || v === 'E' || v === 'N' || v === 'O') ? v : '';
  }
  var lg = (lastGen && lastGen.length === N) ? lastGen : null;
  if (lg) { for (var li = 0; li < N; li++) if (('' + lastGen[li]).length !== nd) { lg = null; break; } }

  var out = {}, has = false;
  if (lg) {
    for (var i = 0; i < N; i++) {
      for (var j = 0; j < nd; j++) {
        var v = norm(cur[i][j]);
        if (!v) continue; // 빈칸 → (잠겨 있었어도) 해제
        var wasLocked = savedPreset && savedPreset[i] && savedPreset[i][j + 1];
        var genVal = lg[i].charAt(j);
        if (wasLocked || v !== genVal) { // 원래 잠금 or 생성본과 다름(=수기 수정)
          if (!out[i]) out[i] = {};
          out[i][j + 1] = v; has = true;
        }
      }
    }
  } else {
    var filled = 0, empty = 0, tmp = {};
    for (var i2 = 0; i2 < N; i2++) {
      for (var j2 = 0; j2 < nd; j2++) {
        var v2 = norm(cur[i2][j2]);
        if (v2) { if (!tmp[i2]) tmp[i2] = {}; tmp[i2][j2 + 1] = v2; filled++; }
        else empty++;
      }
    }
    if (filled > 0 && empty > 0) { out = tmp; has = true; } // 부분 입력 → 고정
  }
  return has ? out : null;
}

function countPresetCells(preset) {
  if (!preset) return 0;
  var c = 0;
  for (var i in preset) for (var d in preset[i]) c++;
  return c;
}

function presetKeys(cfg) {
  var suffix = cfg.year + '_' + cfg.month;
  return { preset: 'duty_preset_' + suffix, lastGen: 'duty_lastgen_' + suffix };
}

/* 저장된 잠금/직전 생성본 로드 (실패 시 빈 상태) */
function loadPresetState(cfg) {
  try {
    var p = PropertiesService.getDocumentProperties();
    var k = presetKeys(cfg);
    return {
      preset: JSON.parse(p.getProperty(k.preset) || 'null'),
      lastGen: JSON.parse(p.getProperty(k.lastGen) || 'null')
    };
  } catch (e) { return { preset: null, lastGen: null }; }
}

function savePresetState(cfg, preset) {
  try {
    var p = PropertiesService.getDocumentProperties();
    var k = presetKeys(cfg);
    if (preset) p.setProperty(k.preset, JSON.stringify(preset));
    else p.deleteProperty(k.preset);
  } catch (e) { /* 저장 실패해도 이번 실행엔 지장 없음 */ }
}

/* 생성 결과를 행 문자열 배열로 저장 → 다음 실행 때 수기 수정 칸 감지 기준 */
function saveLastGenerated(cfg, sched) {
  try {
    var rows = [];
    for (var i = 0; i < cfg.nurses.length; i++) {
      var s = '';
      for (var d = 1; d <= cfg.numDays; d++) s += (sched[i][d] || 'O');
      rows.push(s);
    }
    PropertiesService.getDocumentProperties()
      .setProperty(presetKeys(cfg).lastGen, JSON.stringify(rows));
  } catch (e) { /* 무시 */ }
}

/* 메뉴: 저장된 수기 입력 잠금을 전부 해제 (표 내용은 그대로) */
function resetPresetLock() {
  var cfg = readSettings();
  try {
    var p = PropertiesService.getDocumentProperties();
    var k = presetKeys(cfg);
    p.deleteProperty(k.preset);
    p.deleteProperty(k.lastGen);
  } catch (e) {}
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '수기 입력 잠금을 모두 해제했어요. 다음 [② 자동 배정]은 표 전체를 새로 짭니다.\n' +
    '(표가 부분 입력 상태면 그 칸들은 다시 고정됩니다)', '듀티표', 6);
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

  // 나이트 전후 필수오프 잠금맵 초기화 (placeNight/constructNights가 채움 → 보정이 못 건드림)
  cfg.nightOffLock = [];

  // 0-b) 첫 리퀘스트 오프 전날 = Day(가능할 때만) / 다음날 = 나이트(무조건) → D-O-N 패턴
  //      "첫 리퀘스트 오프" = 설정 시트의 요청오프 + 표에 수기로 입력한 O 중 가장 빠른 날.
  //      · 전날 Day: 빈칸 + 하드규칙 OK + 그날 Day 필요인원 미만일 때만 (충돌 시 건너뜀)
  //      · 다음날 N: 빈칸이면 무조건 고정 (사용자가 직접 친 칸만 존중).
  //        정확 타일링은 이 N과 일치하는 배치만 채택하므로 자연스럽게 나이트 블록이 이어짐
  cfg.forcedDay = [];
  cfg.forcedNight = [];
  cfg.anchorNext = []; // 첫 리퀘스트 오프 "다음날" (나이트 앵커)
  for (var fi = 0; fi < N; fi++) {
    var firstOff = minDayKey(cfg.nurses[fi].reqOff);
    if (cfg.preset && cfg.preset[fi]) { // 표에 직접 친 O(수기 리퀘스트)도 포함
      for (var pok in cfg.preset[fi]) {
        if (cfg.preset[fi][pok] !== 'O') continue;
        var pod = parseInt(pok, 10);
        if (pod >= 1 && (firstOff === 0 || pod < firstOff)) firstOff = pod;
      }
    }
    // 전날 Day (가능할 때만)
    var fb = (firstOff > 1) ? firstOff - 1 : 0;
    if (fb >= 1 && fb <= nd && !sched[fi][fb] &&
        countShift(sched, fb, 'D') < cfg.need.D &&
        canWork(cfg, sched, fi, fb, 'D')) {
      sched[fi][fb] = SHIFT.D;
      cfg.forcedDay[fi] = fb; // isLocked가 보호 → 보정 단계에서 안 바뀜
    }
    // 다음날 나이트는 여기서 미리 박지 않는다 — 미리 박으면 정확 타일링이 거의 다 기각돼
    // 표 품질이 무너짐. 대신 "앵커"로 등록해두고, 나이트 배정 후 같은 역할끼리
    // 블록 스왑으로 그 자리에 맞춘다(enforceFirstOffNight). 점수에도 미준수 벌점.
    // 오프를 연속으로 신청한 경우(예: 7,8) 묶음이 끝난 다음날(9일)이 앵커.
    if (firstOff >= 1) {
      var runEnd = firstOff;
      while (runEnd + 1 <= nd && sched[fi][runEnd + 1] === 'O') runEnd++; // 이 시점 O는 전부 요청/수기 오프
      var na = runEnd + 1;
      if (na <= nd && !sched[fi][na]) cfg.anchorNext[fi] = na;
    }
  }

  // 1) 나이트 블록 배정
  assignNights(cfg, sched, rng);

  // 1-b) 첫 리퀘스트 오프 다음날 = 나이트 강제 (같은 역할 블록 스왑 → 구성·인원 유지)
  enforceFirstOffNight(cfg, sched);

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

  // 3-c) 미달 돌려막기: E 미달인데 자유인이 다 E 불가면, D 근무자를 E로 회전 + 자유인이 D 채움
  rotateShiftDeadlocks(cfg, sched);

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

/* 나이트 전후 필수오프 칸을 잠금 (보정 패스가 이 오프를 근무로 바꾸지 못하게) */
function lockNightOff(cfg, i, day) {
  if (!cfg.nightOffLock) cfg.nightOffLock = [];
  if (!cfg.nightOffLock[i]) cfg.nightOffLock[i] = {};
  cfg.nightOffLock[i][day] = true;
}

/* i의 나이트 블록 목록 [{s,e,len}] */
function nightBlocksOf(cfg, sched, i) {
  var nd = cfg.numDays, out = [];
  for (var d = 1; d <= nd; d++) {
    if (sched[i][d] === 'N' && sched[i][d - 1] !== 'N') {
      var len = 0; while (sched[i][d + len] === 'N') len++;
      out.push({ s: d, e: d + len - 1, len: len });
    }
  }
  return out;
}

/* i의 자동 나이트오프(잠금)를 걷어내고 현재 블록 기준으로 다시 깐다 (블록 이동 후 정리용) */
function rebuildNightOffs(cfg, sched, i) {
  var nd = cfg.numDays, nu = cfg.nurses[i];
  if (cfg.nightOffLock && cfg.nightOffLock[i]) {
    for (var k in cfg.nightOffLock[i]) {
      var d0 = parseInt(k, 10);
      var userOwn = (nu.reqOff && nu.reqOff[d0]) || (cfg.preset && cfg.preset[i] && cfg.preset[i][d0]);
      if (sched[i][d0] === 'O' && !userOwn) sched[i][d0] = '';
    }
    cfg.nightOffLock[i] = {};
  }
  var bs = nightBlocksOf(cfg, sched, i);
  for (var b = 0; b < bs.length; b++) {
    var obn = (bs[b].len >= cfg.nightLen) ? cfg.offBeforeNight : 0;
    for (var x = 1; x <= obn; x++) {
      var bd = bs[b].s - x;
      if (bd >= 1 && !sched[i][bd]) sched[i][bd] = SHIFT.O;
      if (bd >= 1 && sched[i][bd] === SHIFT.O) lockNightOff(cfg, i, bd);
    }
    var oan = offAfterFor(cfg, bs[b].len);
    for (var y = 1; y <= oan; y++) {
      var ad = bs[b].e + y;
      if (ad <= nd && !sched[i][ad]) sched[i][ad] = SHIFT.O;
      if (ad <= nd && sched[i][ad] === SHIFT.O) lockNightOff(cfg, i, ad);
    }
  }
}

/* i가 [s,e] 나이트 블록을 새로 가질 수 있는가 (칸 비었나 + 전후 오프 가능 + 기존 블록과 간격) */
function canHostNightBlock(cfg, sched, i, s, e) {
  var nd = cfg.numDays, len = e - s + 1;
  if (s < 1 || e > nd || len < 1 || len > cfg.nightLen) return false;
  for (var d = s; d <= e; d++) if (sched[i][d] !== '') return false;
  var obn = (len >= cfg.nightLen) ? cfg.offBeforeNight : 0;
  for (var b = 1; b <= obn; b++) {
    var bd = s - b;
    if (bd >= 1) { var v = sched[i][bd]; if (v === 'D' || v === 'E' || v === 'N') return false; }
  }
  var oan = offAfterFor(cfg, len);
  for (var a = 1; a <= oan; a++) {
    var ad = e + a;
    if (ad <= nd) { var v2 = sched[i][ad]; if (v2 === 'D' || v2 === 'E' || v2 === 'N') return false; }
  }
  // 기존 블록들과의 간격 (붙으면 병합돼 길이 초과 위험 → 금지)
  var bs = nightBlocksOf(cfg, sched, i);
  for (var k = 0; k < bs.length; k++) {
    var ob = bs[k];
    if (ob.e >= s - 1 && ob.s <= e + 1) return false; // 겹침/인접
    var gap, need;
    if (ob.e < s) { gap = s - ob.e - 1; need = Math.max(offAfterFor(cfg, ob.len), obn); }
    else { gap = ob.s - e - 1; need = Math.max(oan, (ob.len >= cfg.nightLen) ? cfg.offBeforeNight : 0); }
    if (gap < need) return false;
  }
  // 연속근무: 앞 오프가 없는 짧은 블록은 앞 근무 연속과 합산 검사
  if (obn === 0) {
    var back = 0;
    for (var d2 = s - 1; d2 >= 1; d2--) { var v3 = sched[i][d2]; if (v3 === 'D' || v3 === 'E') back++; else break; }
    if (back + len > cfg.maxConsec) return false;
  }
  return true;
}

/* 첫 리퀘스트 오프 다음날(앵커)에 나이트가 오도록, 같은 역할의 그날 나이트 블록을 스왑.
   X(앵커 주인)가 그날 N 소유자 Y의 블록 꼬리[a..e1]를 가져오고(역할 동일 → 매일 구성 유지),
   대신 X의 다른 블록 하나를 Y에게 넘겨 개수 균형을 맞춘다(없으면 불균형 허용 → 점수가 조정). */
function enforceFirstOffNight(cfg, sched) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  if (!cfg.anchorNext) return;
  for (var X = 0; X < N; X++) {
    var a = cfg.anchorNext[X];
    if (!a || a > nd) continue;
    if (sched[X][a] === 'N') continue;  // 이미 만족
    if (sched[X][a] !== '') continue;   // 칸이 차 있으면(요청/오프 등) 존중
    var role = cfg.nurses[X].charge;
    // 그날 같은 역할의 N 소유자 Y 찾기
    var Y = -1;
    for (var i = 0; i < N; i++)
      if (i !== X && cfg.nurses[i].charge === role && sched[i][a] === 'N') { Y = i; break; }
    if (Y < 0) {
      // 그날 같은 역할 나이트 없음(미달) → 인원 여유 있으면 X가 a부터 새 블록 생성
      if (countShift(sched, a, 'N') < cfg.need.N) {
        var lens = [3, 2, 1];
        for (var li = 0; li < lens.length; li++) {
          if (a + lens[li] - 1 <= nd && canHostNightBlock(cfg, sched, X, a, a + lens[li] - 1)) {
            for (var dn = a; dn <= a + lens[li] - 1; dn++) sched[X][dn] = SHIFT.N;
            rebuildNightOffs(cfg, sched, X);
            cfg.forcedNight[X] = a;
            break;
          }
        }
      }
      continue;
    }
    // Y의 블록(a 포함) — 사용자가 직접 친 N이면 못 건드림
    var s1 = a; while (s1 > 1 && sched[Y][s1 - 1] === 'N') s1--;
    var e1 = a; while (e1 < nd && sched[Y][e1 + 1] === 'N') e1++;
    var locked = false;
    for (var dch = s1; dch <= e1; dch++) if (isLocked(cfg, Y, dch)) { locked = true; break; }
    if (locked) continue;
    // X가 가져갈 부분 = [a..e1] (a 앞은 X의 요청오프라 못 가짐). Y는 [s1..a-1]을 유지.
    // ① Y의 꼬리를 비우고 X가 가질 수 있는지 확인
    for (var dc = a; dc <= e1; dc++) sched[Y][dc] = '';
    rebuildNightOffs(cfg, sched, Y); // Y의 낡은 자동오프 정리(머리 블록 기준 재구성)
    if (!canHostNightBlock(cfg, sched, X, a, e1)) {
      // 롤백
      for (var dr = a; dr <= e1; dr++) sched[Y][dr] = SHIFT.N;
      rebuildNightOffs(cfg, sched, Y);
      continue;
    }
    for (var dx = a; dx <= e1; dx++) sched[X][dx] = SHIFT.N;
    rebuildNightOffs(cfg, sched, X);
    cfg.forcedNight[X] = a;
    // ② 개수 균형: X의 다른 블록 하나(앵커 블록 제외)를 Y에게 넘겨봄 (안 되면 불균형 허용)
    var xb = nightBlocksOf(cfg, sched, X);
    for (var bi = 0; bi < xb.length; bi++) {
      var B2 = xb[bi];
      if (B2.s <= a && a <= B2.e) continue; // 방금 만든 앵커 블록은 제외
      var lockedX = false;
      for (var dlx = B2.s; dlx <= B2.e; dlx++) if (isLocked(cfg, X, dlx)) { lockedX = true; break; }
      if (lockedX) continue;
      for (var dox = B2.s; dox <= B2.e; dox++) sched[X][dox] = '';
      rebuildNightOffs(cfg, sched, X);
      if (canHostNightBlock(cfg, sched, Y, B2.s, B2.e)) {
        for (var dy = B2.s; dy <= B2.e; dy++) sched[Y][dy] = SHIFT.N;
        rebuildNightOffs(cfg, sched, Y);
        break;
      }
      // 못 넘기면 X에게 복구
      for (var dxr = B2.s; dxr <= B2.e; dxr++) sched[X][dxr] = SHIFT.N;
      rebuildNightOffs(cfg, sched, X);
    }
  }
}

/* (i, day)가 사용자가 요청한 칸(요청오프/요청근무)인지 → 보정에서 건드리지 않도록 */
function isLocked(cfg, i, day) {
  var nu = cfg.nurses[i];
  if (nu.reqOff && nu.reqOff[day]) return true;
  if (nu.reqWork && nu.reqWork[day]) return true;
  if (cfg.preset && cfg.preset[i] && cfg.preset[i][day]) return true; // 표에 직접 입력한 칸
  if (cfg.forcedDay && cfg.forcedDay[i] === day) return true;          // 첫 요청오프 전날 Day 고정
  if (cfg.forcedNight && cfg.forcedNight[i] === day) return true;      // 첫 요청오프 다음날 N 고정
  if (cfg.nightOffLock && cfg.nightOffLock[i] && cfg.nightOffLock[i][day]) return true; // 나이트 전후 필수오프
  return false;
}

/* 보정 패스에서 채울 근무 선택: E는 하루 maxEvening(기본 3)까지만, 넘으면 D.
   단, 그 사람이 Day를 선호하면 D를 먼저 시도 */
function chooseFillShift(cfg, sched, i, day) {
  var maxE = cfg.maxEvening || 3;
  if (cfg.nurses[i].prefShift === 'D' && !prefSatisfied(cfg, sched, i) &&
      canWork(cfg, sched, i, day, 'D')) return 'D';
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

  // ② 오프 과다자(U) ↔ 과소자(V) 맞교환 — 인원수는 그대로 두고 오프만 재분배.
  //    (예전엔 빈날에 D를 '추가'해서 오프를 줄였지만, 그러면 그날 인원이 초과(Day 4명)됨.
  //     대신 U가 V의 근무를 가져오고 V는 오프 → 일별 인원 불변, 오프만 균형)
  for (var pass2 = 0; pass2 < nd * 2; pass2++) {
    var changed2 = false;
    for (var U = 0; U < N; U++) {
      if (fullWorkload(sched, U, nd) >= minWork) continue; // U: 오프 과다(일 부족)
      var done = false;
      for (var d2 = 1; d2 <= nd && !done; d2++) {
        // U가 그날 비어있어야(오프/빈칸) 새 근무를 받을 수 있음
        var uc = sched[U][d2];
        if (uc !== '' && uc !== 'O') continue;
        if (isLocked(cfg, U, d2)) continue;
        var shifts = ['E', 'D'];
        for (var si = 0; si < shifts.length && !done; si++) {
          var sh = shifts[si];
          // 그날 V가 sh 근무 중이고, V는 오프 여유(일 과다)면 V를 빼고 U를 넣음
          for (var V = 0; V < N && !done; V++) {
            if (V === U) continue;
            if (sched[V][d2] !== sh) continue;
            if (isLocked(cfg, V, d2)) continue;
            if (fullWorkload(sched, V, nd) <= minWork) continue; // V도 부족하면 안 뺌
            // V가 sh의 유일한 차지면 빼지 않음(차지 커버리지 유지)
            if (cfg.nurses[V].charge && !cfg.nurses[U].charge &&
                countChargeOnShift(cfg, sched, d2, sh) <= 1) continue;
            var saveU = sched[U][d2], saveV = sched[V][d2];
            sched[V][d2] = 'O'; sched[U][d2] = sh;
            // U의 근무 가능성 + V를 뺀 뒤 V의 연속오프 한도 확인
            if (canWork(cfg, sched, U, d2, sh) && !makesLongOff(cfg, sched, V)) { changed2 = true; done = true; }
            else { sched[U][d2] = saveU; sched[V][d2] = saveV; } // 롤백
          }
        }
      }
    }
    if (!changed2) break;
  }
}

/* 미달 근무 돌려막기(데드락 해소): sh가 미달인 날, 같은 날 other 근무자 X(sh 가능)를
   sh로 회전시키고, 빈 other 자리는 자유인 Y(other만 가능한 사람 포함)로 채움.
   예) E 미달인데 자유인은 다 E 불가 → D 근무자 한 명을 E로, 자유인이 그 D를 채움 */
function rotateShiftDeadlocks(cfg, sched) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  var maxWork = nd - cfg.offMin;
  for (var iter = 0; iter < nd * 2; iter++) {
    var fixed = false;
    for (var d = 1; d <= nd && !fixed; d++) {
      var pairs = [['E', 'D'], ['D', 'E']]; // [미달 근무, 회전해올 근무]
      for (var pi = 0; pi < pairs.length && !fixed; pi++) {
        var sh = pairs[pi][0], other = pairs[pi][1];
        if (countShift(sched, d, sh) >= cfg.need[sh]) continue;
        for (var X = 0; X < N && !fixed; X++) {
          if (sched[X][d] !== other) continue;
          if (isLocked(cfg, X, d)) continue;
          var saveX = sched[X][d];
          sched[X][d] = '';
          if (!canWork(cfg, sched, X, d, sh)) { sched[X][d] = saveX; continue; }
          sched[X][d] = sh; // X: other → sh 회전
          // 빈 other 자리를 자유인 Y로 채움
          for (var Y = 0; Y < N; Y++) {
            if (Y === X) continue;
            var saveY = sched[Y][d];
            if (saveY !== '' && saveY !== 'O') continue;
            if (isLocked(cfg, Y, d)) continue;
            if (fullWorkload(sched, Y, nd) >= maxWork) continue;
            if (!canWork(cfg, sched, Y, d, other)) continue;
            sched[Y][d] = other;
            // 두 근무 모두 차지 1명 이상 유지되는지 확인 (아니면 이 Y 롤백 후 다음 Y)
            if (shiftHasCharge(cfg, sched, d, sh) && shiftHasCharge(cfg, sched, d, other)) {
              fixed = true; break;
            }
            sched[Y][d] = saveY;
          }
          if (!fixed) sched[X][d] = saveX; // 채울 Y가 없으면 회전 롤백
        }
      }
    }
    if (!fixed) break;
  }
}

/* day의 shift 근무자 중 차지 수 */
function countChargeOnShift(cfg, sched, day, shift) {
  var c = 0;
  for (var i = 0; i < cfg.nurses.length; i++)
    if (sched[i][day] === shift && cfg.nurses[i].charge) c++;
  return c;
}
/* i의 스케줄에 역할별 연속오프 한도를 넘는 구간이 있는지 */
function makesLongOff(cfg, sched, i) {
  var maxC = cfg.nurses[i].charge ? (cfg.maxConsecOffCharge || 3) : (cfg.maxConsecOffActing || 2);
  var run = 0;
  for (var d = 1; d <= cfg.numDays; d++) {
    if (sched[i][d] === 'O' || sched[i][d] === '') { run++; if (run > maxC) return true; }
    else run = 0;
  }
  return false;
}

/* ── 나이트 배정 디스패처 ──
   하루 2명 체제면 우선 "정확 구성"(전원 목표치 딱 맞춤) 시도.
   요청/수기입력이 있어도 시도한다 — constructNights가 충돌을 스스로 검증해서
   충돌 없는 타일링이면 채택(고품질), 충돌하면 false → 그리디 폴백.
   (300회 attempt마다 타일링이 달라지므로, 요청이 적당하면 무충돌 타일링을 찾게 됨) */
function assignNights(cfg, sched, rng) {
  if (cfg.need.N === 2 && constructNights(cfg, sched, rng)) return; // 정확 구성 성공
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
/* 가중치로 total을 정수 분배(합=total, 최대잉여법). 선호 나이트 목표 계산용 */
function allocByWeight(weights, total) {
  var n = weights.length, wsum = 0, i;
  for (i = 0; i < n; i++) wsum += weights[i];
  if (wsum <= 0) { var ev = []; for (i = 0; i < n; i++) ev[i] = 0; return ev; }
  var raw = [], floors = [], used = 0;
  for (i = 0; i < n; i++) { raw[i] = total * weights[i] / wsum; floors[i] = Math.floor(raw[i]); used += floors[i]; }
  var rem = total - used; // 남은 칸을 잉여(소수부) 큰 순으로 +1
  var order = [];
  for (i = 0; i < n; i++) order.push(i);
  order.sort(function (a, b) { return (raw[b] - floors[b]) - (raw[a] - floors[a]); });
  for (i = 0; i < rem; i++) floors[order[i]]++;
  return floors;
}

/* 한 역할(차지 또는 액팅)이 nd일을 1인 1명씩 덮도록 2~3블록 타일링 → owner[1..nd] (실패 시 null)
   counts: idxs와 같은 길이의 1인당 나이트 수(합=nd). 없으면 균등 분배. */
function tileNightRole(nd, idxs, rng, counts) {
  var count = idxs.length;
  if (count === 0) return null;
  var base = Math.floor(nd / count), extra = nd % count;
  var order = shuffleArr(idxs.slice(), rng);
  var queues = [];
  for (var k = 0; k < count; k++) {
    var nightsK = counts ? counts[idxs.indexOf(order[k])] : (base + (k < extra ? 1 : 0));
    var sizes = splitNightBlocks(nightsK, rng);
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
  // 선호 듀티+강도 → 1인당 나이트 목표를 가중치로 분배(역할 그룹 합은 nd 유지 → 매 밤 커버리지 그대로)
  function nightWeights(idxs) {
    var w = [];
    for (var j = 0; j < idxs.length; j++) w[j] = nightWeightOf(cfg.nurses[idxs[j]]);
    return w;
  }
  var cCounts = allocByWeight(nightWeights(charges), nd);
  var aCounts = allocByWeight(nightWeights(actings), nd);

  // ── 사전 충돌 검증: 요청오프/요청근무/수기입력/전날Day와 부딪히면 그 타일링은 기각 ──
  function tilingValid(co, ao) {
    // ① 타일링이 N을 놓을 칸이 이미 다른 값으로 차 있으면 충돌
    for (var dv = 1; dv <= nd; dv++) {
      var v1 = sched[co[dv]][dv], v2 = sched[ao[dv]][dv];
      if ((v1 !== '' && v1 !== 'N') || (v2 !== '' && v2 !== 'N')) return false;
    }
    // ② 기존 N칸(요청근무 N / 수기 N)이 타일링 소유자와 다르면 그날 N 인원 초과 → 충돌
    for (var iv = 0; iv < N; iv++)
      for (var dv2 = 1; dv2 <= nd; dv2++)
        if (sched[iv][dv2] === 'N' && co[dv2] !== iv && ao[dv2] !== iv) return false;
    // ③ 블록 전후 필수오프 자리가 이미 '근무'(D/E)면 충돌 — 가상 블록으로 미리 검사
    var nightDaysOf = {};
    for (var dv3 = 1; dv3 <= nd; dv3++) {
      (nightDaysOf[co[dv3]] = nightDaysOf[co[dv3]] || {})[dv3] = true;
      (nightDaysOf[ao[dv3]] = nightDaysOf[ao[dv3]] || {})[dv3] = true;
    }
    for (var pk in nightDaysOf) {
      var pi = parseInt(pk, 10), days = nightDaysOf[pk];
      var blocks = [];
      for (var sd = 1; sd <= nd; sd++) {
        if (!days[sd] || days[sd - 1]) continue; // 블록 시작만
        var bl = 0; while (days[sd + bl]) bl++;
        var be = sd + bl - 1;
        blocks.push({ s: sd, e: be, len: bl });
        var ob = (bl >= cfg.nightLen) ? cfg.offBeforeNight : 0;
        for (var bb = 1; bb <= ob; bb++) {
          var bd0 = sd - bb;
          if (bd0 >= 1 && (sched[pi][bd0] === 'D' || sched[pi][bd0] === 'E')) return false;
        }
        var oa = offAfterFor(cfg, bl);
        for (var aa = 1; aa <= oa; aa++) {
          var ad0 = be + aa;
          if (ad0 <= nd && (sched[pi][ad0] === 'D' || sched[pi][ad0] === 'E')) return false;
        }
      }
      // ④ 같은 사람의 블록 사이 간격: 앞 블록 종료후 오프 + 뒷 블록(3연속) 시작전 오프 자리 필요
      for (var bi2 = 1; bi2 < blocks.length; bi2++) {
        var gap = blocks[bi2].s - blocks[bi2 - 1].e - 1;
        var needGap = Math.max(
          offAfterFor(cfg, blocks[bi2 - 1].len),
          (blocks[bi2].len >= cfg.nightLen) ? cfg.offBeforeNight : 0
        );
        if (gap < needGap) return false;
      }
    }
    return true;
  }

  // 요청/수기입력이 많은 달은 무충돌 타일링이 드물다 → attempt 안에서 여러 번 재추첨해
  // 정확 구성(차지1+액팅1, 위반 0) 성공률을 끌어올린다. (요청 없으면 보통 1회에 통과)
  var co = null, ao = null, found = false;
  for (var retry = 0; retry < 30 && !found; retry++) {
    co = tileNightRole(nd, charges, rng, cCounts);
    ao = tileNightRole(nd, actings, rng, aCounts);
    if (co && ao && tilingValid(co, ao)) found = true;
  }
  if (!found) return false;

  for (var d = 1; d <= nd; d++) { sched[co[d]][d] = SHIFT.N; sched[ao[d]][d] = SHIFT.N; }
  // 블록 전후 오프 (3연속만 앞 오프)
  for (var p = 0; p < N; p++) {
    for (var day = 1; day <= nd; day++) {
      if (sched[p][day] === 'N' && sched[p][day - 1] !== 'N') {
        var len = 0; while (sched[p][day + len] === 'N') len++;
        var end = day + len - 1;
        var obn = (len >= cfg.nightLen) ? cfg.offBeforeNight : 0;
        for (var b = 1; b <= obn; b++) { var bd = day - b; if (bd >= 1 && !sched[p][bd]) sched[p][bd] = SHIFT.O; if (bd >= 1 && sched[p][bd] === SHIFT.O) lockNightOff(cfg, p, bd); }
        var oan = offAfterFor(cfg, len);
        for (var a = 1; a <= oan; a++) { var ad = end + a; if (ad <= nd && !sched[p][ad]) sched[p][ad] = SHIFT.O; if (ad <= nd && sched[p][ad] === SHIFT.O) lockNightOff(cfg, p, ad); }
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

  // ── 선호 듀티 반영: 역할 그룹 내에서 나이트 목표를 가중치로 재분배(그룹 합은 유지) ──
  //    N선호 → 목표↑(가중치 2),  상관없음 → 1,  D/E선호 → 목표↓(가중치 0.5)
  //    그룹 합을 그대로 두므로 매 밤 인원/차지 커버리지는 변하지 않음(빈칸 안 생김).
  function prefWeight(nu) { return nightWeightOf(nu); }
  [true, false].forEach(function (isCharge) {
    var grp = [], sum = 0;
    for (var gi = 0; gi < N; gi++) if (cfg.nurses[gi].charge === isCharge) { grp.push(gi); sum += nightTarget[gi]; }
    if (!grp.length || sum <= 0) return;
    var wsum = 0; for (var wi2 = 0; wi2 < grp.length; wi2++) wsum += prefWeight(cfg.nurses[grp[wi2]]);
    if (wsum <= 0) return;
    for (var ki = 0; ki < grp.length; ki++)
      nightTarget[grp[ki]] = sum * prefWeight(cfg.nurses[grp[ki]]) / wsum;
  });

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
    if (bd >= 1 && sched[i][bd] === SHIFT.O) lockNightOff(cfg, i, bd);
  }
  var oan = offAfterFor(cfg, end - start + 1);
  for (var a = 1; a <= oan; a++) {
    var ad = end + a;
    if (ad <= cfg.numDays && !sched[i][ad]) sched[i][ad] = SHIFT.O;
    if (ad <= cfg.numDays && sched[i][ad] === SHIFT.O) lockNightOff(cfg, i, ad);
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
    // 0) 두 번째 자리(커버리지 외)는 "액팅 절대 우선" → 매일 밤 차지1+액팅1 구성 유지.
    //    (이걸 부족분보다 아래에 두면, 선호 재분배로 차지 목표가 액팅보다 커질 때
    //     차지가 또 뽑혀 차지2+액팅0 밤이 생긴다 — 실사용에서 발견된 버그)
    if (!requireCharge) {
      var ra = cfg.nurses[a].charge ? 1 : 0, rb = cfg.nurses[b].charge ? 1 : 0;
      if (ra !== rb) return ra - rb;
    }
    // 1) 목표 대비 부족분(deficit) 큰 사람 먼저 → 각자 목표치까지 채움
    //    (선호는 nightTarget 자체에 이미 반영됨 — N선호=목표↑, D/E선호=목표↓)
    var defa = nightTarget[a] - thisNights[a], defb = nightTarget[b] - thisNights[b];
    if (defa !== defb) return defb - defa;
    // 1-b) 같은 길이 블록을 두 번 안 하도록 → 2일+3일=5로 딱 떨어지게 (4·6 변동 방지)
    var sa = (sizeCount[a][winSize] || 0), sb = (sizeCount[b][winSize] || 0);
    if (sa !== sb) return sa - sb;
    // 2) 전월 나이트 적은 사람 → 동률 랜덤 (캐리오버 공평)
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
  // 2) 남은 슬롯 채우기 — 후보가 적은(빡센) 근무부터 채움 (보통 E가 제약이 많아 먼저)
  var order2 = ['D', 'E'].sort(function (a, b) {
    return countEligible(cfg, sched, day, a, false) - countEligible(cfg, sched, day, b, false);
  });
  order2.forEach(function (sh) {
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
  // 1순위: 전체 근무량 적은 사람(총 근무 고르게 = 오프 고르게, 커버리지 안정) →
  // 2순위: 선호 듀티(같은 근무량이면 그 근무 선호자 먼저) → 3순위: 랜덤.
  // 선호를 커버리지보다 위에 두면 빈칸(미충족)이 생기므로 일부러 2순위로 둠.
  // 이미 선호 목표(prefTargetCount)에 도달한 사람은 보통 사람 취급 → 다른 듀티도 섞임.
  var sat = {};
  if (PREF_DAY_SORT && cfg.prefNudge !== false) {
    for (var pi = 0; pi < pool.length; pi++)
      if (prefSatisfied(cfg, sched, pool[pi])) sat[pool[pi]] = true;
  }
  pool.sort(function (a, b) {
    var diff = fullWorkload(sched, a, cfg.numDays) - fullWorkload(sched, b, cfg.numDays);
    if (diff !== 0) return diff;
    if (PREF_DAY_SORT && cfg.prefNudge !== false) {
      var pa = sat[a] ? 1 : dayPrefRank(cfg.nurses[a], shift);
      var pb = sat[b] ? 1 : dayPrefRank(cfg.nurses[b], shift);
      if (pa !== pb) return pa - pb;
    }
    return rng() - 0.5;
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

/* 해당 칸에 shift 근무 가능한지 (하드 규칙) — 앞날(prev)·뒷날(next) 양방향 모두 검사.
   ※ 뒷날 검사가 없으면, 보정/추가 패스가 빈칸을 채울 때 이미 놓인 다음날과 E→D·N→D/E
      금지패턴을 만들어도 못 걸러서 하드 위반이 생긴다. */
function canWork(cfg, sched, i, day, shift) {
  var prev = day - 1 >= 1 ? sched[i][day - 1] : '';
  var next = day + 1 <= cfg.numDays ? sched[i][day + 1] : '';
  // N 다음날 D/E 금지 (이 칸이 N 다음날인 경우)
  if (prev === 'N') return false;
  // E 다음날 D 금지 (이 칸이 E 다음날인 경우)
  if (shift === 'D' && prev === 'E') return false;
  // 뒷날과의 금지패턴: 이 칸 때문에 다음날이 위반이 되는 경우
  if (shift === 'N' && (next === 'D' || next === 'E')) return false; // N 다음날 D/E
  if (shift === 'E' && next === 'D') return false;                   // E 다음날 D
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
  var unfilled = 0, hard = 0, offDev = 0, overStaff = 0;

  for (var day = 1; day <= nd; day++) {
    WORK_SHIFTS.forEach(function (sh) {
      var diff = cfg.need[sh] - countShift(sched, day, sh);
      if (diff > 0) unfilled += diff;
      if (diff < 0) overStaff += (-diff); // 인원 초과(예: Day 4명)도 위반 → 검색이 회피하도록
      // 차지 없는 근무
      if (countShift(sched, day, sh) > 0 && !shiftHasCharge(cfg, sched, day, sh)) hard++;
    });
  }
  var nightDev = 0;
  var avgNight = (cfg.need.N * nd) / N; // 1인 평균 나이트 일수
  var maxChargeN = 0, minActingN = 1e9, hasActing = false;
  var prefMiss = 0; // 선호 듀티 미반영도(선호 근무를 적게 할수록 ↑) — 적을수록 선호 잘 반영
  var consecOffViol = 0; // 연속오프 한도(차지3/액팅2) 초과 — 탐색이 이런 배치를 피하게 함
  for (var i = 0; i < N; i++) {
    var off = 0, nights = 0, work = 0, matchPref = 0;
    var pref = cfg.nurses[i].prefShift;
    var maxCO = cfg.nurses[i].charge ? (cfg.maxConsecOffCharge || 3) : (cfg.maxConsecOffActing || 2);
    var offRun = 0;
    for (var d = 1; d <= nd; d++) {
      var v = sched[i][d];
      if (v === 'O' || v === '') {
        off++;
        offRun++; if (offRun > maxCO) consecOffViol++;
      }
      else { work++; offRun = 0; if (v === 'N') nights++; if (pref && v === pref) matchPref++; }
    }
    if (off < cfg.offMin) offDev += (cfg.offMin - off);
    if (off > cfg.offMax) offDev += (off - cfg.offMax);
    // 선호가 있는 사람의 나이트 편차는 "의도된 편차"이므로 균등도/역할 벌점에서 제외
    if (!pref) {
      nightDev += Math.abs(nights - avgNight); // 나이트 균등도 (평균에서 벗어난 정도)
      if (cfg.nurses[i].charge) { if (nights > maxChargeN) maxChargeN = nights; }
      else { hasActing = true; if (nights < minActingN) minActingN = nights; }
    }
    // 선호 듀티(D/E): 목표 일수에 "맞추기" — 모자라도, 너무 넘쳐도(100% 쏠림) 벌점
    // → 강하게여도 목표(~75%)까지만 몰아주고 나머지는 다른 듀티를 섞음
    if (pref === 'D' || pref === 'E') {
      var tgt = Math.min(prefTargetCount(cfg, cfg.nurses[i]), work);
      prefMiss += Math.abs(tgt - matchPref) * prefMissWeightOf(cfg.nurses[i]);
    } else if (pref === 'N') {
      // 나이트 선호는 구조적 상한(타일링 목표)이 있어 극단화 위험 없음 → 많을수록 가점 유지
      prefMiss += (work - matchPref) * prefMissWeightOf(cfg.nurses[i]);
    }
  }
  // 액팅이 차지보다 나이트가 적으면(차지 최다 > 액팅 최소) 벌점 → 재굴리기 시 회피
  var roleViol = hasActing && minActingN < 1e9 ? Math.max(0, maxChargeN - minActingN) : 0;
  // 첫 리퀘스트 오프 다음날 나이트(D-O-N) 미준수 — "무조건" 규칙이라 강한 벌점
  var donMiss = 0;
  if (cfg.anchorNext) {
    for (var ai = 0; ai < N; ai++) {
      var aDay = cfg.anchorNext[ai];
      if (aDay && aDay <= nd && sched[ai][aDay] !== 'N') donMiss++;
    }
  }
  return {
    unfilled: unfilled, hard: hard, offDev: offDev, overStaff: overStaff, consecOffViol: consecOffViol,
    nightDev: Math.round(nightDev), roleViol: roleViol, prefMiss: prefMiss, donMiss: donMiss,
    // 커버리지(unfilled/overStaff/hard)·오프는 큰 가중으로 절대 우선 → 선호(prefMiss)는 그 안에서만 best 선택을 좌우
    total: unfilled * 100 + overStaff * 60 + hard * 80 + offDev * 30 + consecOffViol * 25 + donMiss * 50 +
      nightDev * 6 + roleViol * 10 + prefMiss
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
  // 표를 비우면 수기 입력 잠금도 같이 초기화 (완전 새 시작)
  try {
    var p = PropertiesService.getDocumentProperties();
    var k = presetKeys(cfg);
    p.deleteProperty(k.preset);
    p.deleteProperty(k.lastGen);
  } catch (e) {}
}
