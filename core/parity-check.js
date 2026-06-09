#!/usr/bin/env node
/*
 * 동일성 검증: core(_algorithm.js) 와 원본 Code.gs 가 같은 입력·시드에서
 * "완전히 같은 스케줄"을 내는지 확인한다. (추출이 충실한지 보증)
 *
 *   node core/parity-check.js
 */
const fs = require('fs');
const path = require('path');
const { buildConfig } = require('./scheduler');
const A = require('./_algorithm');

// 원본 Code.gs 를 GAS API 스텁과 함께 평가 → tryBuild/evaluate/makeRng 확보
function loadCodeGs() {
  const src = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  const sandbox = {
    SpreadsheetApp: { getActiveSpreadsheet() { return {}; }, getUi() { return { alert() {} }; }, newDataValidation() { return { requireValueInList() { return this; }, build() { return {}; } }; } },
    PropertiesService: { getDocumentProperties() { return { getProperty() { return null; }, setProperty() {}, deleteProperty() {} }; } },
    Utilities: {}, console: console, Math: Math, Date: Date, Number: Number, parseInt: parseInt, parseFloat: parseFloat, JSON: JSON, Array: Array, String: String, Object: Object, isFinite: isFinite,
  };
  const fn = new Function('with(this){' + src + '; return { tryBuild: tryBuild, evaluate: evaluate, makeRng: makeRng };}');
  return fn.call(sandbox);
}

// run-example.js 와 동일한 기본 입력(복제)
const INPUT = {
  year: 2026, month: 6, need: { D: 2, E: 2, N: 2 },
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

const orig = loadCodeGs();
const cfg = buildConfig(INPUT);
const nd = cfg.numDays;

let mismatches = 0, checks = 0;
for (let a = 0; a < 50; a++) {
  // 두 엔진에 동일한 cfg 깊은 복사 + 동일 시드 제공
  const seed = 12345 + a * 2654435761;
  const cfgA = JSON.parse(JSON.stringify(cfg)); cfgA.prefNudge = (a % 2 === 0);
  const cfgB = JSON.parse(JSON.stringify(cfg)); cfgB.prefNudge = (a % 2 === 0);
  const sA = A.tryBuild(cfgA, A.makeRng(seed));
  const sB = orig.tryBuild(cfgB, orig.makeRng(seed));
  for (let i = 0; i < cfg.nurses.length; i++) {
    for (let d = 1; d <= nd; d++) {
      checks++;
      if ((sA[i][d] || 'O') !== (sB[i][d] || 'O')) mismatches++;
    }
  }
}
console.log('비교 칸 수:', checks, '| 불일치:', mismatches);
console.log(mismatches === 0
  ? '✅ core 와 Code.gs 결과 100% 동일 — 추출 충실함'
  : '❌ 불일치 발견 — 추출 점검 필요');
process.exit(mismatches === 0 ? 0 : 1);
