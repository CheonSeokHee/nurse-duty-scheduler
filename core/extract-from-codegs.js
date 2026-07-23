#!/usr/bin/env node
/*
 * Code.gs(구글시트 버전)에서 "순수 스케줄 알고리즘"만 뽑아 core/_algorithm.js 로 출력한다.
 * - 시트 I/O(SpreadsheetApp)·메뉴·PropertiesService 의존 함수는 제외.
 * - Code.gs의 함수 본문을 "있는 그대로" 복사하므로 알고리즘이 갈라지지 않는다(포크 방지).
 *   Code.gs를 고친 뒤 `node core/extract-from-codegs.js` 만 다시 돌리면 core가 동기화됨.
 *
 * 사용: node core/extract-from-codegs.js
 */
const fs = require('fs');
const path = require('path');

const SRC = path.join(__dirname, '..', 'Code.gs');
const OUT = path.join(__dirname, '_algorithm.js');

// 시트/메뉴/저장 I/O 함수 — core에서 제외 (GAS 전용)
const EXCLUDE_FUNCS = new Set([
  'onOpen', 'setupSheets', 'readSettings', 'drawDutyTemplate', 'formatLayout',
  'applyShiftColors', 'generateDuty', 'computePreset', 'countPresetCells',
  'presetKeys', 'loadPresetState', 'savePresetState', 'saveLastGenerated',
  'resetPresetLock', 'writeSchedule', 'checkRules', 'clearDutyValues',
  'syncPrevMonthDuty', 'syncDaysCell', 'onEdit',
]);
// 시트 전용 상수 — 제외 (순수 알고리즘 상수만 유지)
const EXCLUDE_VARS = new Set([
  'SETTINGS_SHEET', 'DUTY_SHEET', 'COLORS', 'NURSE_START_ROW', 'DUTY_DATA_START_ROW',
]);

const src = fs.readFileSync(SRC, 'utf8');
const lines = src.split('\n');

// 최상위 선언(컬럼0의 function/var) 경계 찾기
const decls = [];
const re = /^(?:function\s+(\w+)|var\s+([A-Za-z_]\w*)\s*=)/;
for (let i = 0; i < lines.length; i++) {
  const m = lines[i].match(re);
  if (m) decls.push({ line: i, name: m[1] || m[2], kind: m[1] ? 'fn' : 'var' });
}

// 각 선언의 본문 = [선언줄, 다음 선언줄). 선언 바로 앞의 주석/빈 줄은 그 선언에 붙임.
function leadingCommentStart(declLine, prevEnd) {
  let s = declLine;
  while (s - 1 >= prevEnd) {
    const t = lines[s - 1].trim();
    if (t === '' || t.startsWith('//') || t.startsWith('*') || t.startsWith('/*')) s--;
    else break;
  }
  return s;
}

const kept = [];
let prevEnd = 0;
for (let d = 0; d < decls.length; d++) {
  const cur = decls[d];
  const end = (d + 1 < decls.length) ? decls[d + 1].line : lines.length;
  const drop = (cur.kind === 'fn' && EXCLUDE_FUNCS.has(cur.name)) ||
               (cur.kind === 'var' && EXCLUDE_VARS.has(cur.name));
  if (!drop) {
    const start = leadingCommentStart(cur.line, prevEnd);
    kept.push(lines.slice(start, end).join('\n').replace(/\s+$/, ''));
  }
  prevEnd = end;
}

const header = `/*
 * 자동 생성 파일 — 직접 수정하지 마세요.
 * 원본: Code.gs (구글시트 듀티 배정 알고리즘) → 순수 함수만 추출.
 * 재생성: node core/extract-from-codegs.js
 *
 * 이 파일은 시트/메뉴 I/O가 없는 "스케줄 코어"입니다.
 * 호출 진입점은 core/scheduler.js (buildConfig / generateSchedule) 를 사용하세요.
 */
/* eslint-disable */
`;

// CommonJS export — 추출된 모든 최상위 식별자를 내보낸다.
const exportNames = [];
for (const d of decls) {
  const drop = (d.kind === 'fn' && EXCLUDE_FUNCS.has(d.name)) ||
               (d.kind === 'var' && EXCLUDE_VARS.has(d.name));
  if (!drop) exportNames.push(d.name);
}
const footer = `\nif (typeof module !== 'undefined' && module.exports) {\n  module.exports = { ${exportNames.join(', ')} };\n}\n`;

fs.writeFileSync(OUT, header + '\n' + kept.join('\n\n') + '\n' + footer, 'utf8');
console.log('생성됨:', path.relative(process.cwd(), OUT));
console.log('  추출 함수/상수:', exportNames.length, '개');
console.log('  제외(I/O):', [...EXCLUDE_FUNCS].length, '함수 +', [...EXCLUDE_VARS].length, '상수');
