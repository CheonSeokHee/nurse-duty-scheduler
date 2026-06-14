/*
 * 2단계 데이터 저장소 — 의존성 0, JSON 파일 1개(server/data/db.json).
 * 병동(ward)별로 명단·설정(roster)과 생성한 듀티표 이력(schedules)을 보관한다.
 * 규모가 커지면 이 모듈만 Postgres 등으로 교체하면 됨(인터페이스 유지).
 *
 * 구조:
 *   { wards: { <id>: {
 *       id, name, createdAt,
 *       roster:    { settings:{...}, nurses:[...] },
 *       schedules: [ { id, savedAt, label, config, result } ]
 *   } } }
 */
const fs = require('fs');
const path = require('path');

const DATA_DIR = path.join(__dirname, 'data');
const DB_FILE = path.join(DATA_DIR, 'db.json');

function ensure() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(DB_FILE)) fs.writeFileSync(DB_FILE, JSON.stringify({ wards: {}, users: {}, sessions: {} }, null, 2));
}

function read() {
  ensure();
  let db;
  try {
    db = JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
  } catch (e) {
    db = {};
  }
  db.wards = db.wards || {};
  db.users = db.users || {};
  db.sessions = db.sessions || {};
  return db;
}

// 원자적 저장(임시파일 → rename)으로 쓰는 중 깨짐 방지.
function write(db) {
  ensure();
  const tmp = DB_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(db, null, 2));
  fs.renameSync(tmp, DB_FILE);
}

function id(prefix) {
  return prefix + '_' + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36);
}

// ── 병동 ──
function listWards() {
  const db = read();
  return Object.values(db.wards).map((w) => ({
    id: w.id,
    name: w.name,
    nurseCount: (w.roster && w.roster.nurses ? w.roster.nurses.length : 0),
    scheduleCount: (w.schedules ? w.schedules.length : 0),
    createdAt: w.createdAt,
  })).sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));
}

function createWard(name) {
  const db = read();
  const wid = id('ward');
  db.wards[wid] = {
    id: wid,
    name: (name || '새 병동').toString().trim() || '새 병동',
    createdAt: new Date().toISOString(),
    roster: { settings: {}, nurses: [] },
    schedules: [],
  };
  write(db);
  return db.wards[wid];
}

function getWard(wid) {
  const db = read();
  const w = db.wards[wid];
  if (!w) return null;
  // 이력은 메타만(목록용), 본문은 별도 조회.
  return {
    id: w.id, name: w.name, createdAt: w.createdAt,
    roster: w.roster || { settings: {}, nurses: [] },
    schedules: (w.schedules || []).map((s) => ({ id: s.id, savedAt: s.savedAt, label: s.label })),
  };
}

function renameWard(wid, name) {
  const db = read();
  if (!db.wards[wid]) return null;
  db.wards[wid].name = (name || '').toString().trim() || db.wards[wid].name;
  write(db);
  return db.wards[wid];
}

function deleteWard(wid) {
  const db = read();
  if (!db.wards[wid]) return false;
  delete db.wards[wid];
  write(db);
  return true;
}

// ── 명단/설정 ──
function saveRoster(wid, roster) {
  const db = read();
  if (!db.wards[wid]) return null;
  db.wards[wid].roster = {
    settings: (roster && roster.settings) || {},
    nurses: (roster && roster.nurses) || [],
  };
  write(db);
  return db.wards[wid].roster;
}

// ── 듀티표 이력 ──
function addSchedule(wid, label, config, result) {
  const db = read();
  if (!db.wards[wid]) return null;
  const entry = {
    id: id('sch'),
    savedAt: new Date().toISOString(),
    label: (label || '').toString().trim() ||
      ((config && config.year) ? config.year + '/' + config.month : '듀티표'),
    config, result,
  };
  if (!db.wards[wid].schedules) db.wards[wid].schedules = [];
  db.wards[wid].schedules.unshift(entry); // 최신순
  write(db);
  return { id: entry.id, savedAt: entry.savedAt, label: entry.label };
}

function getSchedule(wid, sid) {
  const db = read();
  const w = db.wards[wid];
  if (!w) return null;
  return (w.schedules || []).find((s) => s.id === sid) || null;
}

function deleteSchedule(wid, sid) {
  const db = read();
  const w = db.wards[wid];
  if (!w || !w.schedules) return false;
  const before = w.schedules.length;
  w.schedules = w.schedules.filter((s) => s.id !== sid);
  if (w.schedules.length === before) return false;
  write(db);
  return true;
}

// ── 사용자(계정) ──
function countUsers() { return Object.keys(read().users).length; }

function getUserByName(username) {
  const db = read();
  const key = (username || '').toString().trim().toLowerCase();
  return Object.values(db.users).find((u) => u.username.toLowerCase() === key) || null;
}

function getUser(uid) { return read().users[uid] || null; }

function listUsers() {
  return Object.values(read().users)
    .map((u) => ({ id: u.id, username: u.username, role: u.role, createdAt: u.createdAt }))
    .sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));
}

// rec: { username, salt, hash, role }
function createUser(rec) {
  const db = read();
  const uid = id('user');
  db.users[uid] = {
    id: uid,
    username: rec.username,
    salt: rec.salt,
    hash: rec.hash,
    role: rec.role || 'user',
    createdAt: new Date().toISOString(),
  };
  write(db);
  return db.users[uid];
}

function deleteUser(uid) {
  const db = read();
  if (!db.users[uid]) return false;
  delete db.users[uid];
  // 해당 사용자의 세션도 정리
  for (const t of Object.keys(db.sessions)) if (db.sessions[t].userId === uid) delete db.sessions[t];
  write(db);
  return true;
}

// ── 세션 ──
function putSession(token, userId, expiresAt) {
  const db = read();
  db.sessions[token] = { userId, expiresAt };
  write(db);
}

function getSession(token) {
  if (!token) return null;
  const db = read();
  const s = db.sessions[token];
  if (!s) return null;
  if (s.expiresAt && Date.parse(s.expiresAt) < Date.now()) { // 만료 정리
    delete db.sessions[token]; write(db); return null;
  }
  return s;
}

function deleteSession(token) {
  const db = read();
  if (!db.sessions[token]) return false;
  delete db.sessions[token];
  write(db);
  return true;
}

module.exports = {
  listWards, createWard, getWard, renameWard, deleteWard,
  saveRoster, addSchedule, getSchedule, deleteSchedule,
  countUsers, getUserByName, getUser, listUsers, createUser, deleteUser,
  putSession, getSession, deleteSession,
};
