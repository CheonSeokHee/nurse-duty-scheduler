/*
 * 자동 생성 파일 — 직접 수정하지 마세요.
 * 원본: Code.gs (구글시트 듀티 배정 알고리즘) → 순수 함수만 추출.
 * 재생성: node core/extract-from-codegs.js
 *
 * 이 파일은 시트/메뉴 I/O가 없는 "스케줄 코어"입니다.
 * 호출 진입점은 core/scheduler.js (buildConfig / generateSchedule) 를 사용하세요.
 */
/* eslint-disable */

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

function setupDayPicker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SETTINGS_SHEET);
  if (!s) { SpreadsheetApp.getUi().alert('설정 시트가 없습니다. 먼저 [① 시트 세팅]을 실행하세요.'); return; }
  s.getRange(2, 3).setValue('← 일수 지정(자동/28~31)').setFontWeight('bold');
  var dayRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['자동', '28', '29', '30', '31'], true).build();
  s.getRange(3, 3).setDataValidation(dayRule);
  if (!s.getRange(3, 3).getValue()) s.getRange(3, 3).setValue('자동');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '설정 시트 C3에 일수 콤보를 추가했어요. 30/31을 직접 고른 뒤 [② 자동 배정]을 누르세요.', '듀티표', 7);
}

/* [진단] 31일 열이 안 보이는 원인 파악용 — 핵심 값을 알림창으로 표시 */

function diagnose31() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SETTINGS_SHEET);
  var rawYear = s ? s.getRange(2, 2).getValue() : '(설정시트없음)';
  var rawMonth = s ? s.getRange(3, 2).getValue() : '(설정시트없음)';
  var cfg = readSettings();
  var d = ss.getSheetByName(DUTY_SHEET);
  var lines = [];
  lines.push('연도셀(B2)=' + rawYear + ' / 월셀(B3)=' + rawMonth);
  lines.push('인식한 numDays = ' + cfg.numDays + ' (31이어야 정상)');
  if (!d) { lines.push('※ 듀티표 시트가 없습니다.'); }
  else {
    lines.push('듀티표 최대열 = ' + d.getMaxColumns() + ' (36 이상이어야 31일+합계 표시)');
    lines.push('듀티표 최대행 = ' + d.getMaxRows());
    var col31 = 2 + 31 - 1; // 31일 칸 열 = 32(AF)
    var hdr31 = d.getRange(2, col31).getValue();
    var first31 = d.getRange(DUTY_DATA_START_ROW, col31).getValue();
    lines.push('헤더 31일칸(2행,32열) = "' + hdr31 + '"');
    lines.push('첫 간호사 31일칸(4행,32열) = "' + first31 + '"');
    var hidden = '';
    try { hidden = d.isColumnHiddenByUser(col31) ? '예(숨김!)' : '아니오'; } catch (e) { hidden = '확인불가'; }
    lines.push('31일 열이 숨김 상태? = ' + hidden);
  }
  SpreadsheetApp.getUi().alert('🔧 31일 진단\n\n' + lines.join('\n'));
}

/* ===================== 시트 세팅 ===================== */

function parseReqOff(v) {
  if (!v) return {};
  var out = {};
  ('' + v).split(',').forEach(function (x) {
    var n = parseInt(('' + x).trim(), 10);
    if (n >= 1) out[n] = true;
  });
  return out;
}

/* 듀티 개수: "D:14/N:5/E:0" -> { D:14, N:5, E:0 } — 이번 달 듀티별 목표 개수.
   (특정 날짜 지정은 듀티표에 직접 입력하면 자동 고정되므로 이 열은 개수 전용)
   0도 의미 있음: E:0 = 이브닝 안 받음 */

function parseDutyCount(v) {
  if (!v) return null;
  var out = {}, has = false;
  ('' + v).split(/[\/,]/).forEach(function (tok) {
    var m = ('' + tok).trim().match(/^([DEN]):(\d+)$/i);
    if (m) { out[m[1].toUpperCase()] = parseInt(m[2], 10); has = true; }
  });
  return has ? out : null;
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

/* i의 한 달 중 특정 듀티 개수 */

function countNurseShift(sched, i, shift, nd) {
  var c = 0;
  for (var d = 1; d <= nd; d++) if (sched[i][d] === shift) c++;
  return c;
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
/* 시트가 최소 needRows행·needCols열을 갖도록 보장 (모자라면 추가). 기본 시트는 26열뿐이라 필수. */

function ensureGrid(sheet, needRows, needCols) {
  var maxC = sheet.getMaxColumns();
  if (maxC < needCols) sheet.insertColumnsAfter(maxC, needCols - maxC);
  var maxR = sheet.getMaxRows();
  if (maxR < needRows) sheet.insertRowsAfter(maxR, needRows - maxR);
}

function tryBuild(cfg, rng) {
  var nd = cfg.numDays;
  var N = cfg.nurses.length;
  // sched[n][day] (day: 1..nd), 0번 인덱스 미사용
  var sched = [];
  for (var i = 0; i < N; i++) {
    sched[i] = [];
    for (var day = 0; day <= nd + 3; day++) sched[i][day] = '';
  }

  // 0) 요청오프 먼저 고정 (듀티 개수 목표는 나이트 배정/점수에서 반영,
  //    특정 날짜 근무 지정은 듀티표에 직접 입력 → preset으로 고정됨)
  for (var i2 = 0; i2 < N; i2++) {
    var nu = cfg.nurses[i2];
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
  // 1-a) 나이트 개수 맞추기: 같은 역할 내 과다↔과소 스왑으로 dutyCount.N 정확히 (슬랙0 보정)
  fixNightCounts(cfg, sched);

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

  // 5) 최종 안전망: 남은 하드 패턴 위반(N다음D/E, E다음D, 연속근무, 나이트 전후오프)은
  //    충돌 근무 칸을 O로 바꿔 "불법(규칙위반)"을 "빈칸(인원미달 경고)"로 강등.
  //    드물게(빡빡한 달) 구성이 못 피하는 1칸을 합법 상태로 만든다.
  repairHardViolations(cfg, sched);

  // 5-b) 잔여 연속오프 강제 해소: 위 단계들이 못 깬 초과 오프런(오프최소에 걸려 못 줄였거나
  //      repairHardViolations가 근무를 O로 바꿔 새로 생긴 런)을 자유오프 1칸을 S로 돌려 끊는다.
  //      하드규칙(연속오프 한도) > 오프최소 목표.
  breakOffRuns(cfg, sched);

  // 6) 인원 과잉 모드: 남는 액팅을 S(보조근무)로 → 오프 10~11 맞춤.
  assignSupport(cfg, sched);
  // 6-b) D/E 초과(3번째)가 생겼으면 그 액팅을 S로 relabel → 데이/이브닝 정확히 2 유지.
  convertOverstaffToS(cfg, sched);

  return sched;
}

/* 인원 과잉 모드에서 D/E가 정규(need)를 넘으면, 초과한 액팅을 S로 바꿔 D/E를 need로 맞춤.
   (S는 별도 보조근무 — 사람은 그대로 근무, 라벨만 D/E→S. 차지·잠금·요청은 안 건드림) */

function convertOverstaffToS(cfg, sched) {
  if (!cfg.surplus) return;
  var nd = cfg.numDays, N = cfg.nurses.length;
  ['D', 'E'].forEach(function (sh) {
    for (var d = 1; d <= nd; d++) {
      var guard = 0;
      while (countShift(sched, d, sh) > cfg.need[sh] && guard++ < N) {
        var pick = -1, pickDuty = -1;
        for (var i = 0; i < N; i++) {
          if (sched[i][d] !== sh) continue;
          if (cfg.nurses[i].charge) continue;          // 액팅만 S
          if (isLocked(cfg, i, d)) continue;            // 요청/수기 보호
          var hasDuty = cfg.nurses[i].dutyCount && cfg.nurses[i].dutyCount[sh] != null;
          if (pick < 0 || (!hasDuty && pickDuty === 1)) { pick = i; pickDuty = hasDuty ? 1 : 0; if (!hasDuty) break; }
        }
        if (pick < 0) break;                            // 초과인데 다 차지/잠금이면 둠
        sched[pick][d] = 'S';
      }
    }
  });
}

/* 목표 인원(데이3/이브닝3)에서 모자란 "3번째"를 S(액팅 추가근무)로 채움.
   "오프 유지": 오프에 여유 있는 사람(오프 > 최소)만 불러서 오프가 범위(10~11) 밖으로 안 나감.
   - ① 그날 쉬면서 오프 여유 있는 액팅을 S로
   - ② 쉬는 여유 액팅이 없으면, 일하는 액팅 B를 S로 돌리고 빈자리는 쉬는 여유 차지 C가 메움
   필수(floor=2)는 정규 배정에서 이미 채워짐. 여유 없으면 2명으로 두고 빨강 안 만듦. */

function assignSupport(cfg, sched) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  function consecOK(i, d) {
    var back = 0; for (var b = d - 1; b >= 1; b--) { var pv = sched[i][b]; if (pv && pv !== 'O') back++; else break; }
    var fwd = 0; for (var f = d + 1; f <= nd; f++) { var nv = sched[i][f]; if (nv && nv !== 'O') fwd++; else break; }
    return back + 1 + fwd <= cfg.maxConsec;
  }
  // S 놓을 수 있는 칸? 그날 오프 + 미잠금 + N다음날 아님 + 연속근무 OK
  function canPlaceS(i, d) {
    return sched[i][d] === 'O' && !isLocked(cfg, i, d) &&
      !(d > 1 && sched[i][d - 1] === 'N') && consecOK(i, d);
  }
  // 액팅이 오프 과다(off > offMax)면, 쉬는 날을 S(보조근무)로 돌려 오프를 offMax까지 내림.
  //   놓는 날은 그날 인원(D+E+S)이 적은 날 우선 → 보조가 필요한 날에 S가 가도록.
  for (var i = 0; i < N; i++) {
    if (cfg.nurses[i].charge) continue;            // S는 액팅만
    var guard = 0;
    while (countOffRow(sched, i, nd) > cfg.offMax && guard++ < nd) {
      var bestD = -1, bestLoad = 1e9;
      for (var d = 1; d <= nd; d++) {
        if (!canPlaceS(i, d)) continue;
        var load = countShift(sched, d, 'D') + countShift(sched, d, 'E') + countShift(sched, d, 'S');
        if (load < bestLoad) { bestLoad = load; bestD = d; }
      }
      if (bestD < 0) break;
      sched[i][bestD] = 'S';
    }
  }
}

/* 최종 하드 패턴 위반 제거 — 충돌 근무 칸을 O로 (잠긴/요청/수기 칸은 보존) */

function repairHardViolations(cfg, sched) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  for (var i = 0; i < N; i++) {
    for (var pass = 0; pass < 5; pass++) {
      var changed = false;
      // N 다음날 D/E, E 다음날 D → 뒷 칸(원인) 제거
      for (var d = 2; d <= nd; d++) {
        var cur = sched[i][d], pv = sched[i][d - 1];
        var bad = ((cur === 'D' || cur === 'E') && pv === 'N') || (cur === 'D' && pv === 'E');
        if (bad && cur !== 'O' && !isLocked(cfg, i, d)) { sched[i][d] = 'O'; changed = true; }
      }
      // 연속근무 초과 → 한도 넘는 칸 제거
      var run = 0;
      for (var d2 = 1; d2 <= nd; d2++) {
        var v = sched[i][d2];
        if (v && v !== 'O') { run++; if (run > cfg.maxConsec && !isLocked(cfg, i, d2)) { sched[i][d2] = 'O'; run = 0; changed = true; } }
        else run = 0;
      }
      // 나이트 전후 필수오프 자리에 근무가 있으면 제거 (나이트 자체는 보존)
      var blocks = nightBlocksOf(cfg, sched, i);
      for (var b = 0; b < blocks.length; b++) {
        var B = blocks[b];
        var obn = (B.len >= cfg.nightLen) ? cfg.offBeforeNight : 0;
        for (var x = 1; x <= obn; x++) { var bd = B.s - x; var bv = sched[i][bd]; if (bd >= 1 && bv && bv !== 'O' && bv !== 'N' && !isLocked(cfg, i, bd)) { sched[i][bd] = 'O'; changed = true; } }
        var oan = offAfterFor(cfg, B.len);
        for (var y = 1; y <= oan; y++) { var ad = B.e + y; var av = sched[i][ad]; if (ad <= nd && av && av !== 'O' && av !== 'N' && !isLocked(cfg, i, ad)) { sched[i][ad] = 'O'; changed = true; } }
      }
      if (!changed) break;
    }
  }
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

/* 잔여 연속오프 강제 해소(하드 보장): limitConsecutiveOff가 못 깬 초과 오프런을,
   런 안의 "자유오프"(요청오프·나이트 전후 필수오프가 아닌 칸) 1개를 S(보조근무)로 돌려 끊는다.
   오프최소(offMin) 밑으로 내려가더라도 하드규칙(연속오프 한도)을 우선한다.
   런이 전부 잠금(예: 두 3연속 사이 gap-3)이라 못 끊으면 그 런은 남긴다(예방은 pickWindowNurse가 담당). */

function breakOffRuns(cfg, sched) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  function consecOK(i, d) {
    var back = 0; for (var b = d - 1; b >= 1; b--) { var pv = sched[i][b]; if (pv && pv !== 'O') back++; else break; }
    var fwd = 0; for (var f = d + 1; f <= nd; f++) { var nv = sched[i][f]; if (nv && nv !== 'O') fwd++; else break; }
    return back + 1 + fwd <= cfg.maxConsec;
  }
  // d를 S(근무)로 바꿔도 되나? 자유오프 + N 다음날 아님 + 3연속N 앞 필수오프 아님 + 연속근무 OK
  function canWorkHere(i, d) {
    if (!(sched[i][d] === 'O' || sched[i][d] === '')) return false;
    if (isLocked(cfg, i, d)) return false;                       // 요청오프·나이트 필수오프 보호
    if (d > 1 && sched[i][d - 1] === 'N') return false;           // N 다음날 근무 금지
    if (d < nd && sched[i][d + 1] === 'N') {                      // 다음날이 3연속 나이트 시작이면 앞 필수오프
      var L = 0; while (sched[i][d + 1 + L] === 'N') L++;
      if (L >= cfg.nightLen) return false;
    }
    return consecOK(i, d);
  }
  for (var i = 0; i < N; i++) {
    var maxCO = cfg.nurses[i].charge ? (cfg.maxConsecOffCharge || 3) : (cfg.maxConsecOffActing || 2);
    for (var pass = 0; pass < nd; pass++) {
      var run = 0, rs = 0, hit = -1;
      for (var d = 1; d <= nd; d++) {
        if (sched[i][d] === 'O' || sched[i][d] === '') { if (run === 0) rs = d; run++; if (run > maxCO) { hit = rs; break; } }
        else run = 0;
      }
      if (hit < 0) break;                                         // 초과 런 없음
      var re = hit; while (re <= nd && (sched[i][re] === 'O' || sched[i][re] === '')) re++; re--;
      // 끊을 칸 후보 순서: 한도 지점(hit+maxCO)에 가까운 칸부터 → 런을 고르게 분할
      var target = hit + maxCO, placed = false;
      for (var span = 0; span <= (re - hit) && !placed; span++) {
        var cands = (span === 0) ? [target] : [target + span, target - span];
        for (var ci = 0; ci < cands.length && !placed; ci++) {
          var cd = cands[ci];
          if (cd >= hit && cd <= re && canWorkHere(i, cd)) { sched[i][cd] = 'S'; placed = true; }
        }
      }
      if (!placed) break;                                         // 전부 잠금 → 못 끊음(남김)
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

/* i의 나이트 배치가 규칙 OK인지 (블록 길이 2~사람별/전역 최대, 1박 금지, N 다음날 근무 없음) */

function nightArrangementOK(cfg, sched, i) {
  var nd = cfg.numDays, mx = cfg.nurses[i].nightMaxLen || cfg.nightLen;
  for (var d = 1; d <= nd; d++) {
    if (sched[i][d] === 'N' && sched[i][d - 1] !== 'N') {
      var len = 0; while (sched[i][d + len] === 'N') len++;
      if (len > mx || len > cfg.nightLen) return false;
      if (len === 1) return false;                            // 1박 금지(전원 — 편혜경·박수진 포함)
    }
    if ((sched[i][d] === 'D' || sched[i][d] === 'E' || sched[i][d] === 'S') && sched[i][d - 1] === 'N') return false;
  }
  return true;
}

/* 나이트 개수를 dutyCount.N에 정확히 맞춤 — 같은 역할 내에서 과다자의 나이트 1일을
   과소자에게 넘김(그날 N 인원수 불변: 같은 역할 1명 교체). 요청N/수기/잠금은 보호. */

function fixNightCounts(cfg, sched) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  function nightsOf(i) { var c = 0; for (var d = 1; d <= nd; d++) if (sched[i][d] === 'N') c++; return c; }
  function tgt(i) { var dc = cfg.nurses[i].dutyCount; return (dc && dc.N != null) ? dc.N : null; }
  [true, false].forEach(function (role) {
    var mem = [];
    for (var i = 0; i < N; i++) if (cfg.nurses[i].charge === role && tgt(i) != null) mem.push(i);
    if (!mem.length) return;
    for (var pass = 0; pass < nd * 2; pass++) {
      var over = -1, under = -1;
      for (var k = 0; k < mem.length; k++) {
        var m = mem[k], n = nightsOf(m);
        if (n > tgt(m) && over < 0) over = m;
        if (n < tgt(m) && under < 0) under = m;
      }
      if (over < 0 || under < 0) break;
      var moved = false;
      for (var d = 1; d <= nd && !moved; d++) {
        if (sched[over][d] !== 'N' || isLocked(cfg, over, d)) continue;
        var uv = sched[under][d];
        if ((uv !== '' && uv !== 'O') || isLocked(cfg, under, d)) continue;
        var so = sched[over][d], su = sched[under][d];
        sched[over][d] = ''; sched[under][d] = 'N';
        rebuildNightOffs(cfg, sched, over); rebuildNightOffs(cfg, sched, under);
        if (nightArrangementOK(cfg, sched, over) && nightArrangementOK(cfg, sched, under)) moved = true;
        else { sched[over][d] = so; sched[under][d] = su; rebuildNightOffs(cfg, sched, over); rebuildNightOffs(cfg, sched, under); }
      }
      if (!moved) break;
    }
  });
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
  var mxI = cfg.nurses[i].nightMaxLen || cfg.nightLen;
  if (len > mxI) return false;                               // 사람별 최대연속 존중
  if (len === 1) return false;                               // 1박 금지(전원 — 편혜경·박수진 포함)
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
    var Yclean = -1; // a가 블록 "시작"인 Y → 통째 스왑 가능(1박 안 생김). 우선 선택.
    for (var i = 0; i < N; i++) {
      if (i === X || cfg.nurses[i].charge !== role || sched[i][a] !== 'N') continue;
      if (sched[i][a - 1] !== 'N') { Yclean = i; break; }  // a부터 Y 블록 시작 → 클린
      if (Y < 0) Y = i;                                     // 아니면 분할 스왑 후보(1박 생길 수 있음)
    }
    if (Yclean >= 0) Y = Yclean;
    if (Y < 0) {
      // 그날 같은 역할 나이트 없음(미달) → 인원 여유 있으면 X가 a부터 새 블록 생성
      if (countShift(sched, a, 'N') < cfg.need.N) {
        var mxL = cfg.nurses[X].nightMaxLen || cfg.nightLen; // 사람별 나이트 최대연속 존중
        var lens = [3, 2].filter(function (L) { return L <= mxL; }); // 1박은 안 만듦(최소 2)
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
    // 1박 금지: 분할 후 Y 머리[s1..a-1]가 1박이거나 X가 [a..e1]에 1박만 가지면 스왑 포기.
    // (클린 스왑 = a가 Y 블록 시작인 경우만 통과 → 1박 안 생김. 그 외엔 앵커 미충족 허용.)
    if (cfg.nightLen >= 2 && ((a - s1) === 1 || (e1 - a + 1) === 1)) continue;
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
  var dc = cfg.nurses[i].dutyCount;
  // 듀티 개수 목표: D 미달이면 D 먼저, E 목표 도달(또는 E:0)이면 E 회피
  if (dc && dc.D != null && countNurseShift(sched, i, 'D', cfg.numDays) < dc.D &&
      canWork(cfg, sched, i, day, 'D')) return 'D';
  var eBlocked = dc && dc.E != null && countNurseShift(sched, i, 'E', cfg.numDays) >= dc.E;
  if (cfg.nurses[i].prefShift === 'D' && !prefSatisfied(cfg, sched, i) &&
      canWork(cfg, sched, i, day, 'D')) return 'D';
  if (!eBlocked && countShift(sched, day, 'E') < maxE && canWork(cfg, sched, i, day, 'E')) return 'E';
  if (canWork(cfg, sched, i, day, 'D')) return 'D';
  return '';
}

/* 보정: ① 부족한 D/E 근무를 여유 인원으로 메꾸고  ② 그래도 일 덜 한 사람은 한가한 날에 추가
   → 빈칸(특히 E) 최소화 + 오프 10~11 유지 */

function topUpUnderworked(cfg, sched, rng) {
  var nd = cfg.numDays, N = cfg.nurses.length;
  var maxWork = nd - cfg.offMin;   // 이 이상 일하면 오프 < 최소 → 금지
  // 1인 최소 근무: 오프최대 기준(nd-offMax)과 "커버리지 기준"(필요인원합×일수/N) 중 큰 값.
  // 인원이 빡빡(예: 12명·하루8명)하면 커버리지 기준이 더 커서 전원을 오프최소(10)로 밀어
  // 빈칸을 없앤다. (안 그러면 누군가 오프11에 머물러 마지막 칸이 미달됨)
  var coverageMin = Math.ceil((cfg.need.D + cfg.need.E + cfg.need.N) * nd / N);
  var minWork = Math.min(maxWork, Math.max(nd - cfg.offMax, coverageMin));
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
  return Math.max(2, cfg.offAfterNight || 2); // 모든 나이트 블록(2박·3박) 뒤 2오프 고정 (액팅·차지 공통)
}
/* 나이트 목표 T를 블록으로 분해.
   maxLen: 1블록 최대 길이(사람별, 기본 3 / 편혜경·박수진=2)
   prefer3: true면 3개씩 묶고(N-N-N) 자투리만 2로 — 액팅 전용
   · maxLen≤2(편혜경·박수진): "한 번에 1~2개" → 2씩 + 자투리(1 허용). 3은 절대 안 만듦.
   · maxLen≥3(그 외): 1박 블록 금지(최소 2, 2·3으로 분해). 자투리 1은 흡수. */

function splitNightBlocks(T, rng, maxLen, prefer3) {
  if (T <= 0) return [];
  var mx = maxLen || 3;
  if (mx <= 2) {                                             // 편혜경·박수진: 오직 2박만 (1박 금지)
    if (T % 2 === 1) return null;                            // 홀수는 2박으로 못 깜 → 폴백(나이트수는 짝수로 정규화됨)
    var b2 = []; for (var r2 = T; r2 > 0; r2 -= 2) b2.push(2); // 예: 6 → [2,2,2]
    return b2;
  }
  if (T === 1) return null;                                  // 1박뿐 = 1박 블록 → 금지(폴백 유도)
  var b = [], r = T;
  if (prefer3) {                                             // 액팅: 3연속 위주, 1박 금지
    while (r >= 3) { b.push(3); r -= 3; }
    if (r === 2) b.push(2);
    else if (r === 1) { b.pop(); b.push(2, 2); }             // …,3,1 → …,2,2 (1박 제거)
    return b;
  }
  while (r > 0) {                                            // 차지(mx=3): 2·3만, 1박 금지
    var size;
    if (r <= mx) size = r;
    else { var x = rng ? rng() : 0.6; size = x < 0.5 ? 2 : 3; }
    b.push(size); r -= size;
  }
  if (b.length >= 2 && b[b.length - 1] === 1) {              // 마지막 1박 제거
    var prev = b[b.length - 2];
    if (prev <= 2) { b[b.length - 2] = prev + 1; b.pop(); }  // 2,1 → 3
    else { b[b.length - 2] = 2; b[b.length - 1] = 2; }       // 3,1 → 2,2
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

/* 역할 그룹(idxs)의 1인당 나이트 수 배분 — 듀티 개수(N:n) 지정자는 그 수를 "고정"하고
   나머지 인원이 잔여분을 선호 가중치로 나눠 갖는다. (지정 합이 총량 초과면 비례 축소) */

function roleNightCounts(cfg, idxs, total) {
  var out = [], fixedSum = 0, freeIdx = [], freeW = [], j;
  for (j = 0; j < idxs.length; j++) {
    var nu = cfg.nurses[idxs[j]];
    if (nu.dutyCount && nu.dutyCount.N != null) {
      var f = Math.max(0, Math.min(nu.dutyCount.N, total));
      out[j] = f; fixedSum += f;
    } else { out[j] = -1; freeIdx.push(j); freeW.push(nightWeightOf(nu)); }
  }
  var rest = total - fixedSum;
  if (rest < 0) { // 지정 합이 역할 총량 초과(입력 과다) → 전체 비례 축소
    var w2 = [];
    for (j = 0; j < idxs.length; j++) w2.push(out[j] > 0 ? out[j] : 0);
    return normalizeNightCounts(cfg, idxs, allocByWeight(w2, total));
  }
  if (freeIdx.length) {
    var fa = allocByWeight(freeW, rest);
    for (var k = 0; k < freeIdx.length; k++) out[freeIdx[k]] = fa[k];
    return normalizeNightCounts(cfg, idxs, out);
  }
  // 자유 인원이 없는데 잔여>0 → 고정 듀티개수 합이 역할 필요 나이트(=총량)에 못 미침.
  //   (예: 7월 31일, 차지 6명×N:5=30 < 31 / 액팅 4명×N:6=24 < 31)
  //   예전엔 그대로 둬서 합<총량 → 타일링 실패 → 그리디 폴백(미충족·듀티개수차 발생).
  //   매일 '차지1+액팅1' 야간 충원은 하드 제약, 시트의 N개수는 소프트 목표이므로
  //   부족분을 나이트 가중치 비례로 고정 인원에게 올려 합=총량을 보장한다(타일링 성공).
  if (rest > 0) {
    // N:0을 "명시"한 사람은 나이트 면제(하드) → 충원 대상에서 제외(가중치 0).
    //   (예: 이브닝만 받는 간호사) 나머지 인원이 부족분을 나눠 진다.
    var w3 = [], anyNonZero = false;
    for (j = 0; j < idxs.length; j++) {
      var nuj = cfg.nurses[idxs[j]];
      var hardZero = nuj.dutyCount && nuj.dutyCount.N === 0;
      var wj = hardZero ? 0 : nightWeightOf(nuj);
      w3.push(wj);
      if (wj > 0) anyNonZero = true;
    }
    // 역할 전원이 N:0이면(나이트 설 사람이 아무도 없음) 어쩔 수 없이 전체 분배
    if (!anyNonZero) for (j = 0; j < idxs.length; j++) w3[j] = nightWeightOf(cfg.nurses[idxs[j]]);
    var add = allocByWeight(w3, rest);
    for (j = 0; j < idxs.length; j++) out[j] += add[j];
  }
  return normalizeNightCounts(cfg, idxs, out);
}

/* 1박 방지용 나이트 수 정규화 (역할 그룹 내, 합 보존):
   "좋은" 나이트 수 = 0, 또는 ≥2이며 (최대3박↑은 아무 값 / 최대2박은 짝수).
   "나쁜" 수 = 1(누구든 1박 강제), 또는 최대2박인데 홀수.
   나쁜 사람을 ±1 조정하고, 합 유지를 위해 다른 사람을 반대로 ±1(조정 후에도 좋은 상태) 짝지운다.
   (예: 박수진5·편혜경5 → 박수진6·편혜경4 / 또는 max3↑ 동료가 한 개 흡수) */
function normalizeNightCounts(cfg, idxs, out) {
  var n = idxs.length;
  if (!n) return out;
  function mx(k) { return cfg.nurses[idxs[k]].nightMaxLen || cfg.nightLen || 3; }
  function good(k, c) { if (c < 0) return false; if (c === 1) return false; if (mx(k) <= 2 && (c % 2) === 1) return false; return true; }
  function bad(k) { return !good(k, out[k]); }
  var guard = 0;
  while (guard++ < n * 4) {
    var k = -1;
    for (var i = 0; i < n; i++) if (bad(i)) { k = i; break; }
    if (k < 0) break;                                  // 전원 정상
    var fixed = false;
    var dirs = [1, -1];                                // 올림 우선(요청 나이트 수 최대한 보장)
    for (var di = 0; di < dirs.length && !fixed; di++) {
      var dk = dirs[di];
      if (!good(k, out[k] + dk)) continue;
      for (var j = 0; j < n; j++) {                    // 합 보존: 반대 방향으로 받아줄 동료
        if (j === k) continue;
        if (good(j, out[j] - dk)) { out[k] += dk; out[j] -= dk; fixed = true; break; }
      }
    }
    if (!fixed) {                                      // 동료 없음 → 1박 방지 우선(합 ±1은 충원/폴백이 흡수)
      if (good(k, out[k] + 1)) out[k] += 1;
      else if (good(k, out[k] - 1)) out[k] -= 1;
      else out[k] = 0;
    }
  }
  return out;
}

/* 한 역할(차지 또는 액팅)이 nd일을 1인 1명씩 덮도록 2~3블록 타일링 → owner[1..nd] (실패 시 null)
   counts: idxs와 같은 길이의 1인당 나이트 수(합=nd). 없으면 균등 분배. */

function tileNightRole(nd, idxs, rng, counts, cfg) {
  var count = idxs.length;
  if (count === 0) return null;
  var base = Math.floor(nd / count), extra = nd % count;
  var order = shuffleArr(idxs.slice(), rng);
  var queues = [];
  for (var k = 0; k < count; k++) {
    var nightsK = counts ? counts[idxs.indexOf(order[k])] : (base + (k < extra ? 1 : 0));
    var mxLen = (cfg && cfg.nurses[order[k]]) ? cfg.nurses[order[k]].nightMaxLen : 3;
    var isActing = !!(cfg && cfg.nurses[order[k]] && !cfg.nurses[order[k]].charge);
    var sizes = splitNightBlocks(nightsK, rng, mxLen, isActing); // 액팅은 3연속 위주로 묶음
    if (sizes === null) return null;
    queues.push({ nurse: order[k], sizes: sizes });
  }
  var seq = [], remain = 0, gi;
  for (gi = 0; gi < queues.length; gi++) remain += queues[gi].sizes.length;
  var guard = 0;
  while (remain > 0 && guard++ < 100000) {
    queues.sort(function (a, b) { return b.sizes.length - a.sizes.length; });
    var n1 = seq.length >= 1 ? seq[seq.length - 1].nurse : -1; // 직전 블록 주인
    var n2 = seq.length >= 2 ? seq[seq.length - 2].nurse : -1; // 그 앞 블록 주인
    var pick = null, qi;
    // ① 최근 2블록 주인과 모두 다른 사람 (같은 사람 블록 간 거리 ≥3 → 사이에 근무 가능한 날 확보 → 3연속 오프 방지)
    for (qi = 0; qi < queues.length; qi++) if (queues[qi].sizes.length > 0 && queues[qi].nurse !== n1 && queues[qi].nurse !== n2) { pick = queues[qi]; break; }
    // ② 안 되면 직전 1블록 주인과만 다르게
    if (!pick) for (qi = 0; qi < queues.length; qi++) if (queues[qi].sizes.length > 0 && queues[qi].nurse !== n1) { pick = queues[qi]; break; }
    // ③ 그래도 없으면 아무거나
    if (!pick) for (qi = 0; qi < queues.length; qi++) if (queues[qi].sizes.length > 0) { pick = queues[qi]; break; }
    seq.push({ nurse: pick.nurse, size: pick.sizes.shift() }); remain--;
  }
  for (var si = 1; si < seq.length; si++) if (seq[si].nurse === seq[si - 1].nurse) return null; // 같은 너스 연속 = 실패
  var owner = [], day = 1;
  for (var bi = 0; bi < seq.length; bi++) for (var t = 0; t < seq[bi].size; t++) { owner[day] = seq[bi].nurse; day++; }
  return (day - 1 === nd) ? owner : null;
}

/* 앵커 정렬 타일링(백트래킹): "요청오프 다음날(anchorNext)"에 그 사람의 나이트 블록이
   "시작"하도록 처음부터 깐다 → 스왑/분할 없이 D-O-N 충족 + 1박 안 생김.
   - 블록은 2~3(최대2인 사람은 1~2). mx≥3은 1박 금지(자투리 1 방지).
   - 같은 사람 블록 연속 금지. 요청오프/기입력 칸(sched 비어있지 않음)엔 안 놓음.
   - 매 밤 정확히 1명(역할 합 = nd) 충족. 못 짜면 null → 호출부가 기존 랜덤 타일링으로 폴백.
   counts[k] = idxs[k]의 목표 나이트 수. */

function tileNightRoleAligned(nd, idxs, rng, counts, cfg, sched) {
  var count = idxs.length;
  if (!count) return null;
  var rem = [], mxK = [];
  for (var k = 0; k < count; k++) { rem[k] = counts ? counts[k] : 0; mxK[k] = Math.min(cfg.nurses[idxs[k]].nightMaxLen || cfg.nightLen, cfg.nightLen); }
  // 앵커: day → local index (그 날 블록 시작이어야 함)
  var anchorByDay = {}, anchorDays = [];
  for (var k2 = 0; k2 < count; k2++) {
    var a = cfg.anchorNext ? cfg.anchorNext[idxs[k2]] : 0;
    if (a && a >= 1 && a <= nd && anchorByDay[a] == null) { anchorByDay[a] = k2; anchorDays.push(a); }
  }
  anchorDays.sort(function (x, y) { return x - y; });
  function nextAnchorAfter(pos) { for (var z = 0; z < anchorDays.length; z++) if (anchorDays[z] > pos) return anchorDays[z]; return nd + 1; }
  // c가 [pos..end]에 나이트 가능? (칸이 비어 있어야 — 요청오프/기입력 충돌 방지)
  function freeFor(c, pos, end) {
    var gi = idxs[c];
    for (var d = pos; d <= end; d++) { var v = sched[gi][d]; if (v !== '' && v !== 'N') return false; }
    return true;
  }
  var owner = [], order0 = [];
  for (var k3 = 0; k3 < count; k3++) order0.push(k3);
  var steps = 0;
  function fill(pos, last) {
    if (++steps > 3000) return false;               // 빠른 실패(분기 폭발 방지) → 호출부가 폴백
    if (pos > nd) { for (var k = 0; k < count; k++) if (rem[k] !== 0) return false; return true; }
    var forced = anchorByDay[pos];
    var cands;
    if (forced != null) { if (forced === last || rem[forced] <= 0) return false; cands = [forced]; }
    else cands = shuffleArr(order0.slice(), rng).filter(function (k) { return rem[k] > 0 && k !== last; });
    var nA = nextAnchorAfter(pos);
    for (var ci = 0; ci < cands.length; ci++) {
      var c = cands[ci], mx = mxK[c];
      var isCharge = cfg.nurses[idxs[c]].charge;
      var sizes = isCharge ? [3, 2] : [3];   // 액팅=3연속만 / 차지=3 또는 2 (큰 블록 우선)
      if (mx < 3) sizes = isCharge ? [2] : [];       // 나이트최대<3(편혜경·박수진): 2박만 (1박 금지)
      for (var si = 0; si < sizes.length; si++) {
        var sz = sizes[si];
        if (sz > mx || sz > rem[c]) continue;
        if (mx >= 3 && (rem[c] - sz) === 1) continue;   // 자투리 1 방지(1박 금지)
        var end = pos + sz - 1;
        if (end > nd) continue;
        if (nA <= end) continue;                        // 블록이 다음 앵커일을 덮으면 안 됨(앵커=블록시작)
        if (!freeFor(c, pos, end)) continue;            // 요청오프/기입력 충돌
        for (var d = pos; d <= end; d++) owner[d] = c;
        rem[c] -= sz;
        if (fill(end + 1, c)) return true;
        rem[c] += sz;
      }
    }
    return false;
  }
  if (!fill(1, -1)) return null;
  var out = [];
  for (var d2 = 1; d2 <= nd; d2++) out[d2] = idxs[owner[d2]];
  return out;
}

/* 정확 구성: 차지/액팅 각각 타일링 → 매 밤 차지1+액팅1, 전원 정확히 목표 나이트. 전후 오프 부여 */

function constructNights(cfg, sched, rng) {
  var nd = cfg.numDays, N = cfg.nurses.length, charges = [], actings = [];
  for (var i = 0; i < N; i++) cfg.nurses[i].charge ? charges.push(i) : actings.push(i);
  if (!charges.length || !actings.length) return false;
  // 1인당 나이트 수: 듀티 개수(N:n) 지정자는 고정, 나머지는 선호 가중치 분배 (역할 합 = nd)
  var cCounts = roleNightCounts(cfg, charges, nd);
  var aCounts = roleNightCounts(cfg, actings, nd);

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

  // 앵커 정렬 점수: 첫 리퀘스트 오프 다음날(anchorNext)에 그 사람 나이트가 시작되면 +1.
  //   (요청오프 다음날=나이트 D-O-N 규칙을 "스왑 없이" 타일링 단계에서 바로 충족 → 1박/연속오프 안전)
  var anchorCount = 0;
  if (cfg.anchorNext) for (var ac = 0; ac < N; ac++) { var ad = cfg.anchorNext[ac]; if (ad && ad <= nd) anchorCount++; }
  function alignScore(co, ao) {
    if (!cfg.anchorNext) return 0;
    var sc = 0;
    for (var i = 0; i < N; i++) {
      var a = cfg.anchorNext[i];
      if (!a || a > nd) continue;
      var own = cfg.nurses[i].charge ? co : ao;     // 유효 타일링은 리퀘스트오프 날 N을 안 줌 → own[a]===i면 블록 시작
      if (own[a] === i) sc++;
    }
    return sc;
  }

  // 요청/수기입력이 많은 달은 무충돌 타일링이 드물다 → 여러 번 재추첨해 (1) 정확 구성 성공률을 올리고
  // (2) 그 중 앵커(D-O-N)를 가장 많이 충족하는 타일링을 고른다. (전부 충족하면 조기 종료)
  var co = null, ao = null, found = false, bestAlign = -1;
  for (var retry = 0; retry < 150; retry++) {
    var c2, a2;
    if (retry < 8) {
      // 앵커 정렬 빌더 우선(요청오프 다음날에 블록 시작). 실패 시 랜덤 타일링으로 폴백.
      c2 = tileNightRoleAligned(nd, charges, rng, cCounts, cfg, sched) || tileNightRole(nd, charges, rng, cCounts, cfg);
      a2 = tileNightRoleAligned(nd, actings, rng, aCounts, cfg, sched) || tileNightRole(nd, actings, rng, aCounts, cfg);
    } else {
      c2 = tileNightRole(nd, charges, rng, cCounts, cfg);  // 다양성/폴백
      a2 = tileNightRole(nd, actings, rng, aCounts, cfg);
    }
    if (!(c2 && a2 && tilingValid(c2, a2))) continue;
    var al = alignScore(c2, a2);
    if (al > bestAlign) { bestAlign = al; co = c2; ao = a2; found = true; }
    if (al >= anchorCount) break;                   // 앵커 전부 정렬 → 더 볼 필요 없음
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

  // ── 역할 그룹 내 목표 재분배: 듀티 개수(N:n) 지정자는 고정, 나머지는 선호 가중치 ──
  //    그룹 합을 그대로 두므로 매 밤 인원/차지 커버리지는 변하지 않음(빈칸 안 생김).
  [true, false].forEach(function (isCharge) {
    var grp = [], sum = 0;
    for (var gi = 0; gi < N; gi++) if (cfg.nurses[gi].charge === isCharge) { grp.push(gi); sum += nightTarget[gi]; }
    if (!grp.length || sum <= 0) return;
    var counts = roleNightCounts(cfg, grp, Math.round(sum));
    for (var ki = 0; ki < grp.length; ki++) nightTarget[grp[ki]] = counts[ki];
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
    // 사람별 나이트 최대연속 초과 윈도우는 그 사람에게 배정 안 함 (편혜경·박수진=2)
    if (winSize > (cfg.nurses[i].nightMaxLen || cfg.nightLen)) continue;
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
  // gap-3 회피: 이 사람이 이미 가진 3연속 블록이 이 윈도 바로 앞에 "한 블록 거리(gap=3)"로 있으면
  // 앞블록 뒤2오프 + 이블록(3연속) 앞1오프 = 3연속 오프가 강제됨 → 나중에 못 고침. 그런 후보는 맨 뒤로.
  var g3 = {};
  for (var pi = 0; pi < pool.length; pi++) {
    var pn = pool[pi], pe = -1;
    for (var pd = start - 1; pd >= 1 && pd >= start - 5; pd--) if (sched[pn][pd] === 'N') { pe = pd; break; }
    var creates = false;
    if (pe >= 0 && winSize >= cfg.nightLen && (start - pe - 1) === 3) {
      var pl = 0, ps = pe; while (ps >= 1 && sched[pn][ps] === 'N') { pl++; ps--; }
      if (pl >= cfg.nightLen) creates = true; // 앞 블록도 3연속일 때만 3연속 오프 강제
    }
    g3[pn] = creates ? 1 : 0;
  }
  pool.sort(function (a, b) {
    // 0) 두 번째 자리(커버리지 외)는 "액팅 절대 우선" → 매일 밤 차지1+액팅1 구성 유지.
    //    (이걸 부족분보다 아래에 두면, 선호 재분배로 차지 목표가 액팅보다 커질 때
    //     차지가 또 뽑혀 차지2+액팅0 밤이 생긴다 — 실사용에서 발견된 버그)
    if (!requireCharge) {
      var ra = cfg.nurses[a].charge ? 1 : 0, rb = cfg.nurses[b].charge ? 1 : 0;
      if (ra !== rb) return ra - rb;
    }
    if (g3[a] !== g3[b]) return g3[a] - g3[b]; // gap-3 유발 후보는 뒤로 (마지막 수단으로만)
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
  // 2순위: 듀티 개수 목표(미달이면 우선/도달이면 후순위) > 선호 듀티 → 3순위: 랜덤.
  // 이미 선호 목표(prefTargetCount)에 도달한 사람은 보통 사람 취급 → 다른 듀티도 섞임.
  var sat = {}, dr = {};
  for (var pi = 0; pi < pool.length; pi++) {
    var p = pool[pi], dcp = cfg.nurses[p].dutyCount;
    if (dcp && dcp[shift] != null)
      dr[p] = (countNurseShift(sched, p, shift, cfg.numDays) < dcp[shift]) ? 0 : 2;
    else if (PREF_DAY_SORT && cfg.prefNudge !== false && prefSatisfied(cfg, sched, p)) sat[p] = true;
  }
  function rankOf(x) {
    if (dr[x] !== undefined) return dr[x]; // 듀티 개수 지정이 최우선 (넛지 토글과 무관)
    if (PREF_DAY_SORT && cfg.prefNudge !== false)
      return sat[x] ? 1 : dayPrefRank(cfg.nurses[x], shift);
    return 1;
  }
  pool.sort(function (a, b) {
    // 듀티 개수 지정은 명시적 지시 → 근무량 균형보다 우선 (목표 도달 시 자동으로 후순위 전환)
    var da = (dr[a] !== undefined) ? dr[a] : 1, db = (dr[b] !== undefined) ? dr[b] : 1;
    if (da !== db) return da - db;
    // 인원 과잉 모드: 차지를 D/E에 먼저 채움 (차지는 S를 못 하니 정규근무로 오프를 맞춰야 함)
    if (cfg.surplus) {
      var ca = cfg.nurses[a].charge ? 0 : 1, cb = cfg.nurses[b].charge ? 0 : 1;
      if (ca !== cb) return ca - cb;
    }
    var diff = fullWorkload(sched, a, cfg.numDays) - fullWorkload(sched, b, cfg.numDays);
    if (diff !== 0) return diff;
    var pa = rankOf(a), pb = rankOf(b);
    if (pa !== pb) return pa - pb;
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
  var floor = cfg.needFloor || cfg.need;
  var unfilled = 0, hard = 0, offDev = 0, overStaff = 0, targetMiss = 0;

  for (var day = 1; day <= nd; day++) {
    var cntD = countShift(sched, day, 'D'), cntE = countShift(sched, day, 'E');
    var cntN = countShift(sched, day, 'N'), sCnt = countShift(sched, day, 'S');
    // 미달(빨강) = floor(=need) 밑 / 초과 = need 위. N·D·E 동일. S는 별도 보조근무라 커버리지에 안 셈.
    unfilled += Math.max(0, floor.N - cntN) + Math.max(0, floor.D - cntD) + Math.max(0, floor.E - cntE);
    overStaff += Math.max(0, cntN - cfg.need.N) + Math.max(0, cntD - cfg.need.D) + Math.max(0, cntE - cfg.need.E);
    WORK_SHIFTS.forEach(function (sh) {
      if (countShift(sched, day, sh) > 0 && !shiftHasCharge(cfg, sched, day, sh)) hard++; // 차지 없는 근무
    });
  }
  var nightDev = 0;
  var avgNight = (cfg.need.N * nd) / N; // 1인 평균 나이트 일수
  var maxChargeN = 0, minActingN = 1e9, hasActing = false;
  var prefMiss = 0; // 선호 듀티 미반영도(선호 근무를 적게 할수록 ↑) — 적을수록 선호 잘 반영
  var consecOffViol = 0; // 연속오프 한도(차지3/액팅2) 초과 — 탐색이 이런 배치를 피하게 함
  var dutyMiss = 0; // 듀티 개수 목표(예: D:14/N:5/E:0)와의 차이 합 — 적을수록 목표 충족
  var patViol = 0; // 하드 패턴 위반: N다음D/E, E다음D, 연속근무 초과, 나이트 전후 필수오프 미충족
  var nightLenViol = 0; // 사람별 나이트 최대연속 초과(편혜경·박수진 등)
  for (var i = 0; i < N; i++) {
    var off = 0, nights = 0, work = 0, matchPref = 0, cntD = 0, cntE = 0, sCount = 0;
    var pref = cfg.nurses[i].prefShift;
    var maxCO = cfg.nurses[i].charge ? (cfg.maxConsecOffCharge || 3) : (cfg.maxConsecOffActing || 2);
    var nMaxL = cfg.nurses[i].nightMaxLen || cfg.nightLen;
    var offRun = 0, workRun = 0;
    for (var d = 1; d <= nd; d++) {
      var v = sched[i][d], pvd = (d > 1) ? sched[i][d - 1] : '';
      if (v === 'O' || v === '') {
        off++; workRun = 0;
        offRun++; if (offRun > maxCO) consecOffViol++;
      }
      else {
        work++; offRun = 0; workRun++;          // S도 근무로 카운트(연속근무에 포함)
        if (workRun > cfg.maxConsec) patViol++;                 // 연속근무 초과
        if ((v === 'D' || v === 'E' || v === 'S') && pvd === 'N') patViol++; // N 다음날 주간/S 금지
        if (v === 'D' && pvd === 'E') patViol++;                // E 다음날 D 금지
        if (v === 'N') nights++; else if (v === 'D') cntD++; else if (v === 'E') cntE++; else if (v === 'S') sCount++;
        if (pref && v === pref) matchPref++;
      }
    }
    // 나이트 블록 길이 & 전후 필수오프
    for (var nb = 1; nb <= nd; nb++) {
      if (sched[i][nb] === 'N' && sched[i][nb - 1] !== 'N') {
        var blen = 0; while (sched[i][nb + blen] === 'N') blen++;
        var bend = nb + blen - 1;
        if (blen > cfg.nightLen && bend < nd) patViol++;        // 나이트 연속 초과(전역)
        if (blen > nMaxL && bend < nd) nightLenViol += (blen - nMaxL); // 사람별 최대연속 초과
        var ob2 = (blen >= cfg.nightLen) ? cfg.offBeforeNight : 0;
        for (var bx = 1; bx <= ob2; bx++) { var bdd = nb - bx; if (bdd >= 1 && sched[i][bdd] !== 'O' && sched[i][bdd] !== '') patViol++; }
        var oa2 = offAfterFor(cfg, blen);
        for (var ax = 1; ax <= oa2; ax++) { var add = bend + ax; if (add <= nd && sched[i][add] !== 'O' && sched[i][add] !== '') patViol++; }
      }
    }
    // S도 정규 근무로 카운트 → 오프(=O일 수)가 10~11에 들도록. (S는 남는 인원을 쓰는 보조근무)
    if (off < cfg.offMin) offDev += (cfg.offMin - off);
    if (off > cfg.offMax) offDev += (off - cfg.offMax);
    // 듀티 개수 목표: 지정된 듀티마다 |목표-실제| 누적
    var dc = cfg.nurses[i].dutyCount;
    if (dc) {
      if (dc.D != null) dutyMiss += Math.abs(dc.D - cntD);
      if (dc.E != null) dutyMiss += Math.abs(dc.E - cntE);
      if (dc.N != null) dutyMiss += Math.abs(dc.N - nights);
    }
    // 선호/개수지정이 있는 사람의 나이트 편차는 "의도된 편차" → 균등도/역할 벌점에서 제외
    if (!pref && !(dc && dc.N != null)) {
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
    patViol: patViol, nightLenViol: nightLenViol, targetMiss: targetMiss, nightDev: Math.round(nightDev), roleViol: roleViol,
    prefMiss: prefMiss, donMiss: donMiss, dutyMiss: dutyMiss,
    // 하드 패턴·필수인원(floor)·오프는 최상위 가중. 3번째 데이/이브닝(targetMiss)은
    // "있으면 좋은" 가벼운 벌점 — S/여유로 채우되 못 채워도 빨강 아님.
    total: patViol * 120 + unfilled * 100 + overStaff * 60 + hard * 80 + offDev * 30 + consecOffViol * 130 +
      donMiss * 50 + dutyMiss * 45 + nightLenViol * 20 + targetMiss * 8 + nightDev * 6 + roleViol * 10 + prefMiss
  };
}

/* ===================== 시트에 기록 ===================== */

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { SHIFT, WORK_SHIFTS, PREF_DAY_SORT, NIGHT_STRENGTH_MAP, PREFMISS_WEIGHT_MAP, PREF_TARGET_RATIO, setupDayPicker, diagnose31, parseReqOff, parseDutyCount, minDayKey, parsePref, dayPrefRank, parsePrefStrength, nightWeightOf, prefMissWeightOf, countNurseShift, prefTargetCount, prefSatisfied, ensureGrid, tryBuild, convertOverstaffToS, assignSupport, repairHardViolations, repairStaffing, limitConsecutiveOff, breakOffRuns, swapOffOut, countOffRow, lockNightOff, nightArrangementOK, fixNightCounts, nightBlocksOf, rebuildNightOffs, canHostNightBlock, enforceFirstOffNight, isLocked, chooseFillShift, topUpUnderworked, rotateShiftDeadlocks, countChargeOnShift, makesLongOff, assignNights, offAfterFor, splitNightBlocks, shuffleArr, allocByWeight, roleNightCounts, tileNightRole, tileNightRoleAligned, constructNights, assignNightsGreedy, placeNight, pickWindowNurse, fillDayEvening, countEligible, fillDayShift, pickDayCandidate, fullWorkload, canWork, countShift, shiftHasCharge, workloadUpTo, makeRng, evaluate };
}
