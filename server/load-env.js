/*
 * 초경량 .env 로더 (의존성 0).  프로젝트 루트의 .env 를 읽어 process.env 에 채운다.
 * 이미 환경에 있는 값은 덮어쓰지 않는다(셸 export 가 우선).
 * 형식: KEY=VALUE 한 줄에 하나. # 주석, 빈 줄 무시. 따옴표는 벗겨준다.
 */
const fs = require('fs');
const path = require('path');

function loadEnv(file) {
  const p = file || path.join(__dirname, '..', '.env');
  let text;
  try {
    text = fs.readFileSync(p, 'utf8');
  } catch (e) {
    return false; // .env 없으면 조용히 통과(셸 env 만으로도 동작)
  }
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) continue;
    const eq = line.indexOf('=');
    if (eq === -1) continue;
    const key = line.slice(0, eq).trim();
    let val = line.slice(eq + 1).trim();
    if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
      val = val.slice(1, -1);
    }
    if (key && process.env[key] === undefined) process.env[key] = val;
  }
  return true;
}

module.exports = { loadEnv };
