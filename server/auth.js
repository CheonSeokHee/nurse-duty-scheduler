/*
 * 사용자 계정 인증 — 의존성 0 (Node 내장 crypto).
 *
 * - 비밀번호: scrypt + 랜덤 salt 해시로 저장 (평문 저장 안 함)
 * - 세션: 로그인 시 랜덤 토큰 발급 → httpOnly 쿠키 'sid' 로 전달, 서버(db)에 보관
 * - 첫 사용자: 계정이 0개면 누구나 "첫 관리자" 가입 가능. 이후 가입은 admin 만.
 *
 * ⚠️ 사내 PoC 수준: HTTPS 강제·요청 제한(rate limit)·CSRF 토큰은 아직 없음.
 *    실서비스로 가면 HTTPS 뒤에 두고 보강 필요.
 */
const crypto = require('crypto');
const store = require('./store');

const SESSION_DAYS = 7;
// 운영(HTTPS) 환경에서 COOKIE_SECURE=1 이면 쿠키에 Secure 플래그를 붙인다.
const SECURE_COOKIE = process.env.COOKIE_SECURE === '1';

// 로그인 무차별 대입 방어 — IP+아이디별 실패 횟수 제한(메모리).
const RL_MAX = 10, RL_WINDOW_MS = 10 * 60 * 1000;
const rlMap = new Map();
function rlKey(req, username) {
  const ip = (req.headers['x-forwarded-for'] || '').split(',')[0].trim() ||
    (req.socket && req.socket.remoteAddress) || '';
  return ip + '|' + (username || '').toLowerCase();
}
function rlBlocked(key) {
  const e = rlMap.get(key);
  if (!e) return false;
  if (Date.now() > e.resetAt) { rlMap.delete(key); return false; }
  return e.count >= RL_MAX;
}
function rlFail(key) {
  const e = rlMap.get(key);
  if (!e || Date.now() > e.resetAt) rlMap.set(key, { count: 1, resetAt: Date.now() + RL_WINDOW_MS });
  else e.count++;
}
function rlReset(key) { rlMap.delete(key); }

function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(String(password), salt, 64).toString('hex');
  return { salt, hash };
}

function verifyPassword(password, salt, hash) {
  const calc = crypto.scryptSync(String(password), salt, 64).toString('hex');
  const a = Buffer.from(calc, 'hex'), b = Buffer.from(hash, 'hex');
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

function parseCookies(req) {
  const out = {};
  const raw = req.headers.cookie || '';
  raw.split(';').forEach((p) => {
    const i = p.indexOf('=');
    if (i > -1) out[p.slice(0, i).trim()] = decodeURIComponent(p.slice(i + 1).trim());
  });
  return out;
}

function publicUser(u) {
  return u ? { id: u.id, username: u.username, role: u.role } : null;
}

// 요청의 쿠키로 현재 로그인 사용자 조회 (없으면 null)
function currentUser(req) {
  const token = parseCookies(req)['sid'];
  const s = store.getSession(token);
  if (!s) return null;
  return store.getUser(s.userId);
}

function setSessionCookie(res, token) {
  const maxAge = SESSION_DAYS * 24 * 3600;
  const secure = SECURE_COOKIE ? '; Secure' : '';
  res.setHeader('Set-Cookie',
    `sid=${token}; HttpOnly; SameSite=Lax; Path=/; Max-Age=${maxAge}${secure}`);
}
function clearSessionCookie(res) {
  const secure = SECURE_COOKIE ? '; Secure' : '';
  res.setHeader('Set-Cookie', `sid=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0${secure}`);
}

function validUsername(u) { return /^[A-Za-z0-9._-]{2,32}$/.test(u || ''); }

/*
 * /api/auth/* 라우팅. send=(status,obj)=>void, readBody=()=>Promise<string>
 * 반환 true = 처리 완료.
 */
async function handleAuth(req, res, url, send, readBody) {
  const m = req.method;

  // 로그인 상태/부트스트랩 정보
  if (url === '/api/auth/me' && m === 'GET') {
    send(200, { ok: true, user: publicUser(currentUser(req)), needsBootstrap: store.countUsers() === 0 });
    return true;
  }

  if (url === '/api/auth/register' && m === 'POST') {
    const body = JSON.parse((await readBody()) || '{}');
    const username = (body.username || '').toString().trim();
    const password = (body.password || '').toString();
    const firstUser = store.countUsers() === 0;
    if (!firstUser) {
      // 첫 사용자가 아니면 admin 만 계정 생성 가능
      const me = currentUser(req);
      if (!me || me.role !== 'admin') { send(403, { ok: false, error: '관리자만 계정을 만들 수 있습니다.' }); return true; }
    }
    if (!validUsername(username)) { send(400, { ok: false, error: '아이디는 영문/숫자/._- 2~32자.' }); return true; }
    if (password.length < 6) { send(400, { ok: false, error: '비밀번호는 6자 이상이어야 합니다.' }); return true; }
    if (store.getUserByName(username)) { send(400, { ok: false, error: '이미 존재하는 아이디입니다.' }); return true; }
    const { salt, hash } = hashPassword(password);
    const role = firstUser ? 'admin' : (body.role === 'admin' ? 'admin' : 'user');
    const u = store.createUser({ username, salt, hash, role });
    send(200, { ok: true, user: publicUser(u) });
    return true;
  }

  if (url === '/api/auth/login' && m === 'POST') {
    const body = JSON.parse((await readBody()) || '{}');
    const key = rlKey(req, body.username);
    if (rlBlocked(key)) {
      send(429, { ok: false, error: '로그인 시도가 너무 많습니다. 10분 후 다시 시도하세요.' });
      return true;
    }
    const u = store.getUserByName(body.username);
    if (!u || !verifyPassword(body.password || '', u.salt, u.hash)) {
      rlFail(key);
      send(401, { ok: false, error: '아이디 또는 비밀번호가 올바르지 않습니다.' });
      return true;
    }
    rlReset(key);
    const token = crypto.randomBytes(32).toString('hex');
    const expiresAt = new Date(Date.now() + SESSION_DAYS * 24 * 3600 * 1000).toISOString();
    store.putSession(token, u.id, expiresAt);
    setSessionCookie(res, token);
    send(200, { ok: true, user: publicUser(u) });
    return true;
  }

  if (url === '/api/auth/logout' && m === 'POST') {
    const token = parseCookies(req)['sid'];
    if (token) store.deleteSession(token);
    clearSessionCookie(res);
    send(200, { ok: true });
    return true;
  }

  // 사용자 관리 (admin)
  if (url === '/api/auth/users' && m === 'GET') {
    const me = currentUser(req);
    if (!me || me.role !== 'admin') { send(403, { ok: false, error: '관리자 전용' }); return true; }
    send(200, { ok: true, users: store.listUsers() });
    return true;
  }
  if (url.startsWith('/api/auth/users/') && m === 'DELETE') {
    const me = currentUser(req);
    if (!me || me.role !== 'admin') { send(403, { ok: false, error: '관리자 전용' }); return true; }
    const uid = url.split('/').pop();
    if (uid === me.id) { send(400, { ok: false, error: '자기 계정은 삭제할 수 없습니다.' }); return true; }
    send(200, { ok: store.deleteUser(uid) });
    return true;
  }

  return false;
}

module.exports = { handleAuth, currentUser };
