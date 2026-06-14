/*
 * 0단계 API 서버 — core/scheduler.js 를 HTTP 로 감싼 최소 껍데기.
 * 의존성 없음(Node 내장 http 만 사용). 실행: node server/index.js
 *
 *   POST /api/schedule   body(JSON): { year, month, need, nurses:[...], seed? ... }
 *                        → 200 { ok:true, result:{...} }  /  400 { ok:false, error }
 *   GET  /               → 브라우저에서 바로 눌러볼 수 있는 테스트 페이지
 *   GET  /api/health     → { ok:true }
 *
 * "서버에서 듀티표가 진짜 생성되는가" 를 먼저 눈으로 확인하기 위한 단계.
 */
require('./load-env').loadEnv(); // .env → process.env (SDK 가 키를 읽기 전에 먼저)
const http = require('http');
const fs = require('fs');
const path = require('path');
const { generateSchedule } = require('../core/scheduler');
const { chat } = require('./chat');
const store = require('./store');
const { handleAuth, currentUser } = require('./auth');

const PORT = process.env.PORT || 3000;

function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'GET,POST,PUT,DELETE,OPTIONS',
  });
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    let size = 0;
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > 2 * 1024 * 1024) { reject(new Error('요청 본문이 너무 큽니다(2MB 초과).')); req.destroy(); return; }
      data += chunk;
    });
    req.on('end', () => resolve(data));
    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const url = (req.url || '/').split('?')[0];

  if (req.method === 'OPTIONS') { sendJson(res, 204, {}); return; }

  if (req.method === 'GET' && url === '/api/health') {
    sendJson(res, 200, { ok: true, service: 'duty-scheduler', time: new Date().toISOString() });
    return;
  }

  // 인증 라우트(로그인/로그아웃/가입/me/사용자관리)
  if (url.startsWith('/api/auth/')) {
    try {
      const handled = await handleAuth(req, res, url, (s, o) => sendJson(res, s, o), () => readBody(req));
      if (handled) return;
    } catch (err) {
      sendJson(res, 400, { ok: false, error: String(err && err.message || err) });
      return;
    }
  }

  // 보호 게이트: 그 외 모든 /api 는 로그인 필요
  if (url.startsWith('/api/') && url !== '/api/health') {
    if (!currentUser(req)) { sendJson(res, 401, { ok: false, error: '로그인이 필요합니다.' }); return; }
  }

  if (req.method === 'POST' && url === '/api/schedule') {
    const t0 = Date.now();
    try {
      const raw = await readBody(req);
      const input = raw ? JSON.parse(raw) : {};
      const result = generateSchedule(input);
      sendJson(res, 200, { ok: true, elapsedMs: Date.now() - t0, result });
    } catch (err) {
      sendJson(res, 400, { ok: false, error: String(err && err.message || err) });
    }
    return;
  }

  if (req.method === 'POST' && url === '/api/chat') {
    const t0 = Date.now();
    try {
      const raw = await readBody(req);
      const body = raw ? JSON.parse(raw) : {};
      const out = await chat(body);
      sendJson(res, 200, { ok: true, elapsedMs: Date.now() - t0, ...out });
    } catch (err) {
      sendJson(res, 400, { ok: false, error: String(err && err.message || err) });
    }
    return;
  }

  // ── 병동(ward) + 듀티표 이력 API ──
  if (url.startsWith('/api/wards')) {
    try {
      const handled = await handleWards(req, res, url);
      if (handled) return;
    } catch (err) {
      sendJson(res, 400, { ok: false, error: String(err && err.message || err) });
      return;
    }
  }

  if (req.method === 'GET' && url === '/') {
    const file = path.join(__dirname, 'public', 'index.html');
    fs.readFile(file, (e, buf) => {
      if (e) { res.writeHead(500); res.end('test page missing'); return; }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(buf);
    });
    return;
  }

  sendJson(res, 404, { ok: false, error: 'not found' });
});

/*
 * /api/wards 라우팅. 경로:
 *   GET    /api/wards                          병동 목록
 *   POST   /api/wards            {name}        병동 생성
 *   GET    /api/wards/:id                       병동 상세(명단+이력메타)
 *   PUT    /api/wards/:id        {name}         병동 이름 변경
 *   DELETE /api/wards/:id                        병동 삭제
 *   PUT    /api/wards/:id/roster {settings,nurses}  명단/설정 저장
 *   GET    /api/wards/:id/schedules/:sid          저장된 듀티표 본문
 *   POST   /api/wards/:id/schedules {label,config,result}  듀티표 이력 저장
 *   DELETE /api/wards/:id/schedules/:sid          이력 삭제
 * 반환값 true = 처리 완료(응답 보냄).
 */
async function handleWards(req, res, url) {
  const parts = url.split('/').filter(Boolean); // ['api','wards', id?, 'schedules'?, sid?]
  const m = req.method;
  const wid = parts[2];
  const sub = parts[3];
  const sid = parts[4];

  if (!wid) {
    if (m === 'GET') { sendJson(res, 200, { ok: true, wards: store.listWards() }); return true; }
    if (m === 'POST') {
      const body = JSON.parse((await readBody(req)) || '{}');
      sendJson(res, 200, { ok: true, ward: store.createWard(body.name) });
      return true;
    }
    return false;
  }

  // /api/wards/:id/schedules ...
  if (sub === 'schedules') {
    if (m === 'POST') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const meta = store.addSchedule(wid, body.label, body.config, body.result);
      if (!meta) { sendJson(res, 404, { ok: false, error: '병동을 찾을 수 없습니다.' }); return true; }
      sendJson(res, 200, { ok: true, schedule: meta });
      return true;
    }
    if (m === 'GET' && sid) {
      const s = store.getSchedule(wid, sid);
      if (!s) { sendJson(res, 404, { ok: false, error: '듀티표를 찾을 수 없습니다.' }); return true; }
      sendJson(res, 200, { ok: true, schedule: s });
      return true;
    }
    if (m === 'DELETE' && sid) {
      const ok = store.deleteSchedule(wid, sid);
      sendJson(res, ok ? 200 : 404, { ok });
      return true;
    }
    return false;
  }

  // /api/wards/:id/roster
  if (sub === 'roster' && m === 'PUT') {
    const body = JSON.parse((await readBody(req)) || '{}');
    const r = store.saveRoster(wid, body);
    if (!r) { sendJson(res, 404, { ok: false, error: '병동을 찾을 수 없습니다.' }); return true; }
    sendJson(res, 200, { ok: true, roster: r });
    return true;
  }

  // /api/wards/:id
  if (!sub) {
    if (m === 'GET') {
      const w = store.getWard(wid);
      if (!w) { sendJson(res, 404, { ok: false, error: '병동을 찾을 수 없습니다.' }); return true; }
      sendJson(res, 200, { ok: true, ward: w });
      return true;
    }
    if (m === 'PUT') {
      const body = JSON.parse((await readBody(req)) || '{}');
      const w = store.renameWard(wid, body.name);
      sendJson(res, w ? 200 : 404, { ok: !!w, ward: w });
      return true;
    }
    if (m === 'DELETE') {
      const ok = store.deleteWard(wid);
      sendJson(res, ok ? 200 : 404, { ok });
      return true;
    }
  }
  return false;
}

server.listen(PORT, () => {
  console.log(`듀티 스케줄러 API 가동 → http://localhost:${PORT}`);
  console.log(`  • 브라우저에서 열어 테스트: http://localhost:${PORT}/`);
  console.log(`  • POST http://localhost:${PORT}/api/schedule`);
});
