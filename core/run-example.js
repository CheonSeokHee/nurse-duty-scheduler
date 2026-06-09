#!/usr/bin/env node
/*
 * 코어 단독 실행 데모.
 *   node core/run-example.js            # 기본 11명(2026/6) 표 생성·출력
 *   node core/run-example.js 2026 7     # 연/월 지정
 *   node core/run-example.js 2026 6 42  # 시드 고정(재현용)
 *
 * 입력 JSON 파일로 돌리려면:
 *   node core/run-example.js --in myinput.json
 */
const { generateSchedule } = require('./scheduler');

// 구글시트 "기본 11명" 과 동일한 예시 입력
const DEFAULT_INPUT = {
  year: 2026, month: 6,
  need: { D: 2, E: 2, N: 2 },
  maxConsec: 4, nightLen: 3, offBeforeNight: 1, offAfterNight: 1,
  maxNightBlocks: 2, nightMode: 3, offMin: 10, offMax: 11, attempts: 300,
  nurses: [
    { name: '편혜경', role: '차지', reqOff: '21,22', dutyCount: 'D:14/N:5/E:0', nightMax: 2 },
    { name: '이선정', role: '차지', reqOff: '14,15', dutyCount: 'N:5' },
    { name: '박수진', role: '차지', reqOff: '19,20', dutyCount: 'N:5', nightMax: 2 },
    { name: '전초희', role: '차지', reqOff: '11,12', dutyCount: 'N:5' },
    { name: '김경진', role: '차지', reqOff: '7,8', dutyCount: 'N:5' },
    { name: '박지연', role: '차지', reqOff: '', dutyCount: 'N:5' },
    { name: '서문휘정', role: '액팅', reqOff: '12,13', dutyCount: 'N:6' },
    { name: '이서현', role: '액팅', reqOff: '27,28', dutyCount: 'N:6' },
    { name: '정선희', role: '액팅', reqOff: '', dutyCount: 'N:6' },
    { name: '곽예은', role: '액팅', reqOff: '14,15', dutyCount: 'N:6' },
    { name: '오혜경', role: '액팅', reqOff: '', dutyCount: '' },
  ],
};

function parseArgs(argv) {
  const a = argv.slice(2);
  const inIdx = a.indexOf('--in');
  if (inIdx >= 0) {
    const fs = require('fs');
    return JSON.parse(fs.readFileSync(a[inIdx + 1], 'utf8'));
  }
  const input = JSON.parse(JSON.stringify(DEFAULT_INPUT));
  if (a[0]) input.year = Number(a[0]);
  if (a[1]) input.month = Number(a[1]);
  if (a[2]) input.seed = Number(a[2]);
  return input;
}

function printTable(r) {
  const pad = (s, n) => (s + ' '.repeat(n)).slice(0, n);
  const days = Array.from({ length: r.numDays }, (_, i) => i + 1);
  const head = pad('이름', 9) + pad('역할', 5) + days.map(d => String(d).padStart(2)).join('');
  console.log(head);
  console.log('-'.repeat(head.length));
  for (const n of r.nurses) {
    const row = pad(n.name, 9) + pad(n.role, 5) + n.shifts.map(s => s.padStart(2)).join('');
    console.log(row);
  }
  // 합계(검증용)
  console.log('-'.repeat(head.length));
  ['D', 'E', 'N', 'S'].forEach(sh => {
    const counts = days.map(d => r.nurses.filter(n => n.shifts[d - 1] === sh).length);
    console.log(pad(sh + ' 합', 14) + counts.map(c => String(c).padStart(2)).join(''));
  });
}

const input = parseArgs(process.argv);
const t0 = Date.now();
const r = generateSchedule(input);
const ms = Date.now() - t0;

printTable(r);
console.log('');
console.log(`${r.year}/${r.month} (${r.numDays}일) · ${r.attempts}회 탐색 · ${ms}ms`);
console.log('검증통과(빈칸·위반0):', r.clean ? '✅' : '⚠️', JSON.stringify({
  미충족: r.score.unfilled, 초과: r.score.overStaff, 연속오프: r.score.consecOffViol,
  패턴위반: r.score.patViol, 오프편차: r.score.offDev, 듀티개수차: r.score.dutyMiss,
  'D-O-N미충족': r.score.donMiss,
}));
