/*
 * 듀티 스케줄 코어 — 진입점.
 * Code.gs(구글시트)에서 추출한 순수 알고리즘(_algorithm.js)을 감싸,
 * 시트 없이 "평범한 JS 객체 입력 → 스케줄 출력"으로 쓸 수 있게 한다.
 * (서버/도커/Lambda 어디서든 이 모듈만 require 하면 됨)
 *
 *   const { generateSchedule } = require('./scheduler');
 *   const result = generateSchedule({ year, month, nurses: [...] });
 */
const A = require('./_algorithm');

/* 입력(평범한 객체)을 알고리즘 내부 cfg 로 변환. readSettings(시트)와 동일 규칙. */
function buildConfig(input) {
  input = input || {};
  const need = Object.assign({ D: 2, E: 2, N: 2 }, input.need || {});
  const cfg = {
    year: input.year,
    month: input.month,
    need: { D: need.D, E: need.E, N: need.N },
    maxConsec: numOr(input.maxConsec, 4),
    nightLen: numOr(input.nightLen, 3),
    offBeforeNight: numOr(input.offBeforeNight, 1),
    offAfterNight: numOr(input.offAfterNight, 1),
    maxNightBlocks: numOr(input.maxNightBlocks, 2),
    nightMode: numOr(input.nightMode, 2), // 1=인원우선 2=균형 3=나이트균등
    offMin: numOr(input.offMin, 10),
    offMax: numOr(input.offMax, 11),
    attempts: numOr(input.attempts, 300),
  };
  // 탐색 한도(선택): 빡빡한 달에 더 굴릴 상한/시간예산
  cfg.maxAttempts = numOr(input.maxAttempts, 6000);
  cfg.timeBudgetSec = numOr(input.timeBudgetSec, 45);
  cfg.needFloor = { D: need.D, E: need.E, N: need.N };

  cfg.nurses = (input.nurses || []).map(function (n) {
    const role = (n.role || '액팅').toString().trim();
    let nMax = parseInt(n.nightMax, 10);
    if (!(nMax >= 1)) nMax = cfg.nightLen;
    nMax = Math.min(nMax, cfg.nightLen);
    const prevNightDays = Number(n.prevNightDays) || 0;
    return {
      name: (n.name || '').toString().trim(),
      charge: role === '차지' || role === 'charge' || n.charge === true,
      reqOff: typeof n.reqOff === 'object' && n.reqOff !== null ? n.reqOff : A.parseReqOff(n.reqOff),
      dutyCount: typeof n.dutyCount === 'object' && n.dutyCount !== null ? n.dutyCount : A.parseDutyCount(n.dutyCount),
      prevNightDays: prevNightDays,
      prevBlocks: Math.round(prevNightDays / (cfg.nightLen || 3)),
      prefShift: n.prefShift && /^[DEN]$/.test(n.prefShift) ? n.prefShift : A.parsePref(n.prefShift),
      prefStrength: numOr(n.prefStrength, 2),
      nightMaxLen: nMax,
    };
  }).filter(function (n) { return n.name; });

  cfg.numDays = new Date(cfg.year, cfg.month, 0).getDate();
  const coreWork = (cfg.need.D + cfg.need.E + cfg.need.N) * cfg.numDays;
  const minTotalWork = cfg.nurses.length * (cfg.numDays - cfg.offMax);
  cfg.surplus = coreWork < minTotalWork;
  // 표에 직접 박아둔 칸(고정) — 선택. { nurseIndex: { day: 'D'|'E'|'N'|'O' } }
  cfg.preset = input.preset || null;
  return cfg;
}

function numOr(v, dflt) {
  const n = Number(v);
  return isFinite(n) && v !== '' && v !== null && v !== undefined ? n : dflt;
}

function isClean(s) {
  return s.unfilled === 0 && s.overStaff === 0 && s.offDev === 0 &&
    s.consecOffViol === 0 && s.hard === 0 && s.patViol === 0;
}

/*
 * 스케줄 생성. (generateDuty의 탐색 루프 — 시트 I/O 제거판)
 * input: buildConfig 참고. 추가 옵션: seed(고정 시드, 재현용)
 * 반환: { year, month, numDays, attempts, clean, score, nurses:[{name,role,shifts:[...]}] }
 */
function generateSchedule(input) {
  const cfg = buildConfig(input);
  if (!cfg.nurses.length) throw new Error('간호사 목록(nurses)이 비어 있습니다.');
  if (!cfg.year || !cfg.month) throw new Error('year, month 가 필요합니다.');

  const base = (input && input.seed != null)
    ? (input.seed >>> 0)
    : Math.floor(Math.random() * 2000000000);

  let best = null;
  const minA = Math.max(1, cfg.attempts || 300);
  const hardCap = Math.max(minA, cfg.maxAttempts || 6000);
  const budgetMs = (cfg.timeBudgetSec || 45) * 1000;
  const startMs = Date.now();
  let triedCount = 0;

  for (let a = 0; a < hardCap; a++) {
    cfg.prefNudge = (a % 2 === 0);
    const rng = A.makeRng(base + a * 2654435761 + 12345);
    const sched = A.tryBuild(cfg, rng);
    const score = A.evaluate(cfg, sched);
    triedCount = a + 1;
    if (!best || score.total < best.score.total) best = { sched: sched, score: score };
    if (best.score.total === 0) break;
    if (a + 1 >= minA) {
      if (isClean(best.score)) break;
      if (Date.now() - startMs > budgetMs) break;
    }
  }

  const nd = cfg.numDays;
  const nurses = cfg.nurses.map(function (n, i) {
    const shifts = [];
    for (let d = 1; d <= nd; d++) shifts.push(best.sched[i][d] || 'O');
    return { name: n.name, role: n.charge ? '차지' : '액팅', shifts: shifts };
  });

  return {
    year: cfg.year,
    month: cfg.month,
    numDays: nd,
    attempts: triedCount,
    clean: isClean(best.score),
    score: best.score,
    nurses: nurses,
  };
}

module.exports = { generateSchedule, buildConfig, algorithm: A };
