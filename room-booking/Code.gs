/**
 * 장소 예약 시스템 (Google Apps Script)
 * - 모든 시간은 대한민국 표준시(Asia/Seoul) 기준
 * - 스프레드시트에 연결된(컨테이너 바운드) 스크립트로 사용하세요.
 * - 제목, 장소, 횟수 제한은 코드가 아니라 스프레드시트([설정], [장소] 시트)에서 바꿉니다.
 */

// ===================== 고정 설정 =====================
const TZ = 'Asia/Seoul';
const SH = { USERS: '학생', RES: '예약', HOL: '휴무일', LOG: '로그', PLACE: '장소', CFG: '설정', CHK: '퇴실확인' };
const PERIODS = ['아침', '점심', '방과후'];
const OPEN_HOUR = 8;                        // 전주 토요일 08:00 오픈
const CLOSE_HOUR = 18;                      // 전주 일요일 18:00 마감
const SESSION_SEC = 21600;                  // 로그인 유지 6시간
const MAX_FAIL = 5;                         // 로그인 5회 실패 시
const LOCK_SEC = 600;                       // 10분 잠금
const DOW = ['일', '월', '화', '수', '목', '금', '토'];
const RES_HEAD = ['예약ID', '날짜', '장소ID', '시간대', '아이디', '이름', '팀명', '예약시각', '상태', '취소시각', '장소명(자동)'];
const DEFAULT_TITLE = '동아리실·요가실 예약';
const DEFAULT_PLACES = ['동아리실', '요가실'];
function cfgRows_() {
  return [
    ['사이트제목', DEFAULT_TITLE, '로그인 화면, 상단, 브라우저 탭에 표시되는 제목'],
    ['아침 최대횟수', '2', '한 팀이 한 주에 예약할 수 있는 아침(08:00~08:20) 최대 횟수'],
    ['점심 최대횟수', '2', '한 팀이 한 주에 예약할 수 있는 점심(12:30~13:10) 최대 횟수'],
    ['방과후 최대횟수', '2', '한 팀이 한 주에 예약할 수 있는 방과후 최대 횟수'],
    ['아침 사진마감', '08:40', '아침 예약의 퇴실 확인 사진 제출 마감 시각 (당일, 24시간제 HH:MM)'],
    ['방과후 사진마감', '16:50', '방과후 예약의 퇴실 확인 사진 제출 마감 시각 (당일, 24시간제 HH:MM)'],
    ['사진 촬영 방식', '라이브카메라만', '라이브카메라만 = 앱 안에서 실시간 촬영만 허용 / 카메라앱허용 = 라이브 카메라가 안 될 때 휴대폰 카메라 앱 촬영도 허용'],
    ['퇴실확인 시작일', addDays_(mondayOf_(todayKst_()), 7), '이 날짜 이후 예약부터 퇴실 확인이 필요합니다. 비우면 퇴실 확인 기능이 꺼집니다. (YYYY-MM-DD)']
  ];
}

class UserError extends Error {
  constructor(msg, title) { super(msg); this.userMsg = msg; this.title = title || '알림'; }
}

// ===================== 웹앱 진입 =====================
function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.title = getSettings_().title;
  return t.evaluate().setTitle(t.title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

// ===================== 시트 메뉴 / 초기 설정 =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('예약시스템')
    .addItem('① 초기 설정 (시트 만들기/업데이트)', 'setup')
    .addItem('② 비밀번호 암호화', 'hashAllPasswords')
    .addToUi();
}

/** [장소] 시트에 새 줄을 추가하면 번호(P3, P4 ...)를 자동으로 붙입니다. */
function onEdit(e) {
  try { if (e && e.range && e.range.getSheet().getName() === SH.PLACE) ensurePlaceIds_(); } catch (x) { /* 무시 */ }
}

function setup() {
  const ss = ss_();
  const mk = (name, head) => {
    let sh = ss.getSheetByName(name), isNew = false;
    if (!sh) { sh = ss.insertSheet(name); isNew = true; }
    if (isNew || sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, head.length).setValues([head]);
      sh.setFrozenRows(1);
    }
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8efff');
    return { sh, isNew };
  };
  const setHead = (sh, col, text) => { sh.getRange(1, col).setValue(text).setFontWeight('bold').setBackground('#e8efff'); };

  // 학생
  const u = mk(SH.USERS, ['아이디', '비밀번호', '이름', '팀명', '역할']);
  u.sh.getRange('A:E').setNumberFormat('@');
  u.sh.getRange('E2:E300').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['학생', '매니저', '학생+매니저', '교사'], true).build());
  if (u.isNew) {
    u.sh.getRange(2, 1, 3, 5).setValues([
      ['선생님', '1234', '기술가정 선생님', '', '교사'],
      ['김민준', '1111', '김민준', '푸른팀', '학생'],
      ['이하늘', '2222', '이하늘', '푸른팀', '학생']
    ]);
  }
  u.sh.setColumnWidths(1, 5, 120);

  // 장소
  const p = mk(SH.PLACE, ['번호(수정 금지)', '장소명', '사용(O/X)']);
  p.sh.getRange('A:C').setNumberFormat('@');
  p.sh.getRange('C2:C100').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['O', 'X'], true).build());
  if (p.isNew) {
    p.sh.getRange(2, 1, DEFAULT_PLACES.length, 3).setValues(DEFAULT_PLACES.map((n, i) => ['P' + (i + 1), n, 'O']));
  }
  p.sh.setColumnWidths(1, 3, 130);
  ensurePlaceIds_();

  // 설정
  const c = mk(SH.CFG, ['항목', '값', '설명']);
  c.sh.getRange('A:B').setNumberFormat('@');
  const have = c.sh.getLastRow() >= 2 ? c.sh.getRange(2, 1, c.sh.getLastRow() - 1, 1).getDisplayValues().map(r => String(r[0]).trim()) : [];
  cfgRows_().forEach(r => { if (have.indexOf(r[0]) < 0) c.sh.getRange(c.sh.getLastRow() + 1, 1, 1, 3).setNumberFormat('@').setValues([r]); });
  c.sh.setColumnWidth(1, 140); c.sh.setColumnWidth(2, 220); c.sh.setColumnWidth(3, 420);

  // 예약
  const r = mk(SH.RES, RES_HEAD);
  setHead(r.sh, 3, RES_HEAD[2]); setHead(r.sh, 11, RES_HEAD[10]);
  r.sh.getRange('B:C').setNumberFormat('@');
  r.sh.getRange('H:H').setNumberFormat('@');
  r.sh.getRange('J:J').setNumberFormat('@');
  r.sh.setColumnWidths(1, 11, 110);
  migrateReservations_();

  // 휴무일
  const h = mk(SH.HOL, ['날짜', '사유', '시간대(비우면 전체)']);
  setHead(h.sh, 3, '시간대(비우면 전체)');
  h.sh.getRange('A:C').setNumberFormat('@');
  h.sh.getRange('C2:C300').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['전체'].concat(PERIODS), true).build());
  h.sh.setColumnWidths(1, 3, 150);

  // 퇴실확인 (사진 제출 기록)
  const k = mk(SH.CHK, ['예약ID', '날짜', '시간대', '장소ID', '팀명', '제출자ID', '제출자', '제출시각', '촬영방식', '사진링크', '파일ID', '면제', '메모', '장소명(자동)']);
  k.sh.getRange('A:M').setNumberFormat('@');
  k.sh.setColumnWidths(1, 14, 110);
  const cfgSh = ss.getSheetByName(SH.CFG);
  if (cfgSh && cfgSh.getLastRow() >= 2) {
    cfgSh.getRange(2, 1, cfgSh.getLastRow() - 1, 1).getDisplayValues().forEach((r, i) => {
      if (String(r[0]).trim() === '사진 촬영 방식') {
        cfgSh.getRange(i + 2, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['라이브카메라만', '카메라앱허용'], true).build());
      }
    });
  }
  getPhotoFolder_();   // 구글 드라이브 사진 폴더 생성 (처음에는 권한 승인 창이 뜹니다)

  // 로그
  const l = mk(SH.LOG, ['시각', '아이디', '동작', '내용']);
  l.sh.getRange('A:A').setNumberFormat('@');

  const first = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  ss.setSpreadsheetTimeZone(TZ);
  SpreadsheetApp.getUi().alert('초기 설정이 끝났습니다.\n\n' +
    '• [학생] 시트: 아이디(본명)·비밀번호·이름·팀명·역할(학생/매니저/학생+매니저/교사)\n' +
    '• [장소] 시트: 장소 추가/이름 변경 (번호는 건드리지 마세요)\n' +
    '• [설정] 시트: 사이트 제목, 시간대별 최대 횟수, 사진 마감 시각, 촬영 방식, 퇴실확인 시작일\n' +
    '• [퇴실확인] 시트: 사진 제출 기록 (자동)\n' +
    '• [휴무일] 시트: 날짜, 사유, 시간대(비우면 하루 전체)');
}

/** 예약 시트의 장소 값을 이름에서 번호(P1, P2…)로 바꾸고, 장소명 자동 열을 채웁니다. */
function migrateReservations_() {
  const sh = ss_().getSheetByName(SH.RES);
  if (!sh || sh.getLastRow() < 2) return;
  PLACES_MEMO = null;
  const n = sh.getLastRow() - 1;
  const rng = sh.getRange(2, 3, n, 1);
  const vals = rng.getDisplayValues();
  const out = vals.map(r => [normPlace_(r[0])]);
  rng.setNumberFormat('@').setValues(out);
  sh.getRange(2, 11, n, 1).setFormulas(out.map((r, i) => [placeFormula_(i + 2)]));
}
function placeFormula_(row) {
  return "=IFERROR(VLOOKUP(C" + row + ",'" + SH.PLACE + "'!A:B,2,FALSE),C" + row + ")";
}

/** [학생] 시트의 비밀번호를 SHA-256으로 변환합니다. (변환 후에는 원래 비밀번호를 볼 수 없습니다) */
function hashAllPasswords() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('비밀번호 암호화',
    '평문 비밀번호를 모두 암호화합니다.\n변환 후에는 시트에서 원래 비밀번호를 볼 수 없습니다.\n(재설정하려면 새 비밀번호를 평문으로 입력한 뒤 다시 실행)\n\n진행할까요?',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  const sh = sheet_(SH.USERS), last = sh.getLastRow();
  if (last < 2) return;
  const rng = sh.getRange(2, 2, last - 1, 1);
  const vals = rng.getDisplayValues();
  let n = 0;
  const out = vals.map(r => {
    const p = String(r[0]).trim();
    if (!p || p.indexOf('sha256:') === 0) return [r[0]];
    n++; return [hashPw_(p)];
  });
  rng.setNumberFormat('@').setValues(out);
  ui.alert(n + '개의 비밀번호를 암호화했습니다.');
}

// ===================== 유틸 (날짜·시간, KST) =====================
function pad_(n) { return n < 10 ? '0' + n : '' + n; }
function parseYmd_(s) {
  const m = String(s || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
}
function fmtYmd_(d) { return d.getUTCFullYear() + '-' + pad_(d.getUTCMonth() + 1) + '-' + pad_(d.getUTCDate()); }
function normDate_(s) { const d = parseYmd_(s); return d ? fmtYmd_(d) : ''; }
function addDays_(s, n) { const d = parseYmd_(s); d.setUTCDate(d.getUTCDate() + n); return fmtYmd_(d); }
function dowOf_(s) { return parseYmd_(s).getUTCDay(); }
function mondayOf_(s) { return addDays_(s, -((dowOf_(s) + 6) % 7)); }
function todayKst_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowStamp_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function kstDate_(ymd, h) { return new Date(ymd + 'T' + pad_(h) + ':00:00+09:00'); }
function label_(ymd) { const d = parseYmd_(ymd); return (d.getUTCMonth() + 1) + '/' + d.getUTCDate() + '(' + DOW[d.getUTCDay()] + ')'; }

/** 해당 주(월요일 기준)의 예약 가능 기간: 전주 토 08:00 ~ 전주 일 18:00 */
function windowOf_(monday) {
  const sat = addDays_(monday, -2), sun = addDays_(monday, -1);
  const open = kstDate_(sat, OPEN_HOUR), close = kstDate_(sun, CLOSE_HOUR), now = new Date();
  return {
    state: now < open ? 'before' : (now < close ? 'open' : 'closed'),
    openMs: open.getTime(), closeMs: close.getTime(),
    openLabel: label_(sat) + ' ' + pad_(OPEN_HOUR) + ':00',
    closeLabel: label_(sun) + ' ' + pad_(CLOSE_HOUR) + ':00'
  };
}
function slotTime_(period, dow) {
  if (period === '아침') return '08:00~08:20';
  if (period === '점심') return '12:30~13:10';
  return (dow === 2 || dow === 4) ? '16:00~16:30' : '15:00~16:30';
}
function defaultMonday_() {
  const t = todayKst_(), d = dowOf_(t), mon = mondayOf_(t);
  return (d === 6 || d === 0) ? addDays_(mon, 7) : mon;      // 주말이면 다음 주(예약 대상 주)
}
function defaultReportMonday_() {
  // 예약이 마감된 가장 최근 주: 일요일 18시가 지나면 다음 주, 아니면 이번 주
  const mon = mondayOf_(todayKst_()), next = addDays_(mon, 7);
  return windowOf_(next).state === 'closed' ? next : mon;
}

// ===================== 시트 접근 =====================
function ss_() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  if (s) return s;
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('스프레드시트가 연결되어 있지 않습니다.');
  return SpreadsheetApp.openById(id);
}
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('시트 [' + name + ']가 없습니다. 메뉴에서 초기 설정을 실행하세요.');
  return sh;
}

// --- 설정 ---
var SETTINGS_MEMO = null;
function getSettings_() {
  if (SETTINGS_MEMO) return SETTINGS_MEMO;
  const s = { title: DEFAULT_TITLE, limits: { '아침': 2, '점심': 2, '방과후': 2 },
    deadline: { '아침': { h: 8, m: 40 }, '방과후': { h: 16, m: 50 } }, cam: 'live', checkoutStart: '' };
  try {
    const sh = ss_().getSheetByName(SH.CFG);
    if (sh && sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(r => {
        const k = String(r[0]).trim(), v = String(r[1]).trim();
        if (k === '사이트제목') { if (v) s.title = v; return; }
        if (k === '사진 촬영 방식') { s.cam = v.indexOf('앱') >= 0 ? 'app' : 'live'; return; }
        if (k === '퇴실확인 시작일') { s.checkoutStart = normDate_(v); return; }
        const dm = /^(아침|방과후) 사진마감$/.exec(k);
        if (dm) { const t = hm_(v); if (t && t.h < 24 && t.m < 60) s.deadline[dm[1]] = t; return; }
        PERIODS.forEach(p => {
          if (k === p + ' 최대횟수') { const n = parseInt(v, 10); if (!isNaN(n) && n >= 0) s.limits[p] = n; }
        });
      });
    }
  } catch (e) { /* 기본값 사용 */ }
  SETTINGS_MEMO = s;
  return s;
}

// --- 장소 ---
var PLACES_MEMO = null;
function getPlaces_(includeOff) {
  if (!PLACES_MEMO) {
    const sh = ss_().getSheetByName(SH.PLACE);
    let list = [];
    if (!sh) {
      list = DEFAULT_PLACES.map((n, i) => ({ id: 'P' + (i + 1), name: n, on: true }));
    } else if (sh.getLastRow() >= 2) {
      list = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues().map(r => ({
        id: String(r[0]).trim(), name: String(r[1]).trim(),
        on: !/^(x|n|0|아니오|미사용)$/i.test(String(r[2]).trim())
      })).filter(p => p.id && p.name);
    }
    PLACES_MEMO = list;
  }
  return includeOff ? PLACES_MEMO : PLACES_MEMO.filter(p => p.on);
}
function activePlaces_() {
  const l = getPlaces_(false);
  if (!l.length) throw new UserError('사용 중인 장소가 없습니다. [장소] 시트에 장소를 입력하세요.');
  return l;
}
function placeName_(id) { const p = getPlaces_(true).find(x => x.id === id); return p ? p.name : id; }
/** 예전 데이터(장소 이름)도 번호로 바꿔서 읽습니다. */
function normPlace_(v) {
  v = String(v).trim();
  const all = getPlaces_(true);
  if (all.some(p => p.id === v)) return v;
  const byName = all.find(p => p.name === v);
  return byName ? byName.id : v;
}
function ensurePlaceIds_() {
  const sh = ss_().getSheetByName(SH.PLACE);
  if (!sh || sh.getLastRow() < 2) return;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
  let max = 0;
  v.forEach(r => { const m = /^P(\d+)$/.exec(String(r[0]).trim()); if (m) max = Math.max(max, +m[1]); });
  let changed = false;
  const out = v.map(r => {
    if (!String(r[0]).trim() && String(r[1]).trim()) { changed = true; max++; return ['P' + max]; }
    return [r[0]];
  });
  if (changed) { sh.getRange(2, 1, out.length, 1).setNumberFormat('@').setValues(out); PLACES_MEMO = null; }
}

// --- 사용자 ---
function getUsers_() {
  const sh = sheet_(SH.USERS), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getDisplayValues().map((r, i) => {
    const rs = String(r[4]);
    const teacher = rs.indexOf('교사') >= 0, mgr = rs.indexOf('매니저') >= 0;
    const student = rs.indexOf('학생') >= 0 || (!teacher && !mgr);
    return {
      row: i + 2,
      id: String(r[0]).trim(),
      pw: String(r[1]).trim(),
      name: String(r[2]).trim() || String(r[0]).trim(),
      team: String(r[3]).trim(),
      role: teacher ? 'teacher' : (student ? 'student' : 'manager'),
      report: teacher || mgr                 // 예약 달력 출력 탭을 볼 수 있는가
    };
  }).filter(u => u.id);
}

// --- 예약 ---
function getRes_() {
  const sh = sheet_(SH.RES), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 10).getDisplayValues().map((r, i) => ({
    row: i + 2, id: r[0], date: normDate_(r[1]), place: normPlace_(r[2]), period: r[3],
    uid: r[4], name: r[5], team: r[6], at: r[7], status: String(r[8]).trim()
  })).filter(r => r.id && r.date);
}
function getActiveRes_() { return getRes_().filter(r => r.status === '예약'); }

// --- 휴무일 ---
function periodOfCell_(s) {
  s = String(s || '');
  for (let i = 0; i < PERIODS.length; i++) if (s.indexOf(PERIODS[i]) >= 0) return PERIODS[i];
  return '전체';
}
function getHolidays_() {
  const sh = ss_().getSheetByName(SH.HOL);
  if (!sh || sh.getLastRow() < 2) return [];
  const out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues().forEach((r, i) => {
    const d = normDate_(r[0]);
    if (d) out.push({ row: i + 2, date: d, reason: String(r[1]).trim() || '휴무', period: periodOfCell_(r[2]) });
  });
  return out;
}
function holidayFor_(hols, date, period) {
  const h = hols.find(x => x.date === date && (x.period === '전체' || x.period === period));
  return h ? h.reason : null;
}
function allDayHoliday_(hols, date) {
  const h = hols.find(x => x.date === date && x.period === '전체');
  return h ? h.reason : null;
}

function log_(uid, action, detail) {
  try { sheet_(SH.LOG).appendRow([nowStamp_(), uid, action, detail || '']); } catch (e) { /* 로그 실패는 무시 */ }
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new UserError('접속자가 많아 처리가 지연되고 있습니다. 잠시 후 다시 시도해 주세요.');
  try { const r = fn(); SpreadsheetApp.flush(); return r; } finally { lock.releaseLock(); }
}
function newId_() { return 'R' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5); }

// ===================== 인증 =====================
function normId_(s) { return String(s || '').normalize('NFC').trim().toLowerCase(); }
function keyOf_(prefix, s) {
  return prefix + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8));
}
function salt_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SALT');
  if (!s) { s = Utilities.getUuid(); p.setProperty('SALT', s); }
  return s;
}
function hashPw_(pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt_() + ':' + String(pw).normalize('NFC'), Utilities.Charset.UTF_8);
  return 'sha256:' + bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function safeEq_(a, b) {
  if (a.length !== b.length) return false;
  let d = 0; for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function checkPw_(stored, input) {
  input = String(input).trim();
  if (stored.indexOf('sha256:') === 0) return safeEq_(stored, hashPw_(input));
  return safeEq_(stored.normalize('NFC'), input.normalize('NFC'));
}
function sessErr_() { const e = new Error('session'); e.code = 'SESSION'; return e; }
function pub_(u) { return { id: u.id, name: u.name, team: u.team, role: u.role, report: u.report }; }
function defaults_() { return { bookMonday: defaultMonday_(), reportMonday: defaultReportMonday_() }; }

function login(id, pw) {
  try {
    id = String(id || '').trim(); pw = String(pw || '').trim();
    if (!id || !pw) return { ok: false, msg: '아이디와 비밀번호를 입력하세요.' };
    const cache = CacheService.getScriptCache(), fk = keyOf_('fail:', normId_(id));
    const fails = parseInt(cache.get(fk) || '0', 10);
    if (fails >= MAX_FAIL) return { ok: false, msg: '로그인 실패가 ' + MAX_FAIL + '회 이상입니다. 10분 후 다시 시도하세요.' };
    const u = getUsers_().find(x => normId_(x.id) === normId_(id));
    if (!u || !checkPw_(u.pw, pw)) {
      cache.put(fk, String(fails + 1), LOCK_SEC);
      log_(id, '로그인실패', (fails + 1) + '회');
      return { ok: false, msg: '아이디 또는 비밀번호가 올바르지 않습니다. (' + (fails + 1) + '/' + MAX_FAIL + ')' };
    }
    cache.remove(fk);
    try { ensurePlaceIds_(); } catch (e) { /* 무시 */ }
    const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
    cache.put('sess:' + token, u.id, SESSION_SEC);
    log_(u.id, '로그인', '');
    return { ok: true, token: token, user: pub_(u), defaults: defaults_(), nowMs: Date.now() };
  } catch (e) {
    console.error(e.stack || e);
    return { ok: false, msg: '서버 오류: ' + e.message };
  }
}
function restore(token) {
  return api_(token, u => ({ ok: true, user: pub_(u), defaults: defaults_(), nowMs: Date.now() }));
}
function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess:' + token);
  return { ok: true };
}
function auth_(token) {
  if (!token) throw sessErr_();
  const cache = CacheService.getScriptCache(), uid = cache.get('sess:' + token);
  if (!uid) throw sessErr_();
  const u = getUsers_().find(x => x.id === uid);       // 시트 변경(팀 이동, 삭제)을 즉시 반영
  if (!u) throw sessErr_();
  cache.put('sess:' + token, uid, SESSION_SEC);
  return u;
}
/** roles: 허용 역할 배열(없으면 모두 허용) */
function api_(token, fn, roles) {
  try {
    const u = auth_(token);
    if (roles && roles.indexOf(u.role) < 0) throw new UserError('권한이 없습니다.');
    return fn(u);
  } catch (e) {
    if (e && e.code === 'SESSION') return { ok: false, code: 'SESSION', msg: '로그인이 만료되었습니다. 다시 로그인해 주세요.' };
    if (e instanceof UserError) return { ok: false, msg: e.userMsg, title: e.title };
    console.error(e.stack || e);
    return { ok: false, msg: '서버 오류가 발생했습니다. 잠시 후 다시 시도해 주세요. (' + e.message + ')' };
  }
}

function changePassword(token, oldPw, newPw) {
  return api_(token, u => {
    newPw = String(newPw || '').trim();
    if (!checkPw_(u.pw, oldPw)) throw new UserError('현재 비밀번호가 올바르지 않습니다.');
    if (newPw.length < 4) throw new UserError('새 비밀번호는 4자 이상이어야 합니다.');
    if (checkPw_(u.pw, newPw)) throw new UserError('현재와 다른 비밀번호를 입력하세요.');
    sheet_(SH.USERS).getRange(u.row, 2).setNumberFormat('@').setValue(hashPw_(newPw));
    log_(u.id, '비밀번호변경', '');
    return { ok: true };
  });
}

// ===================== 주간 조회 =====================
function weekDates_(monday) { return [0, 1, 2, 3, 4].map(i => addDays_(monday, i)); }
function periodIdx_(p) { return PERIODS.indexOf(p); }
const TEAM_COLORS = 12;
/** 팀 이름 -> 색 번호(0~11). [학생] 시트에 처음 나온 순서대로 배정되어 항상 같은 팀은 같은 색입니다. */
function teamIndexMap_(extraTeams) {
  const m = {}; let n = 0;
  getUsers_().forEach(u => { if (u.role === 'student' && u.team && !(u.team in m)) m[u.team] = n++; });
  (extraTeams || []).forEach(t => { if (t && !(t in m)) m[t] = n++; });
  return m;
}
function listText_(arr) {
  return arr.slice().sort((a, b) => a.date.localeCompare(b.date) || periodIdx_(a.period) - periodIdx_(b.period))
    .map(r => '• ' + r.name + ' — ' + label_(r.date) + ' ' + r.period + ' ' + placeName_(r.place)).join('\n');
}

function getWeek(token, monday) {
  return api_(token, u => {
    monday = mondayOf_(normDate_(monday) || todayKst_());
    const dates = weekDates_(monday), hols = getHolidays_(), win = windowOf_(monday), places = activePlaces_();
    const settings = getSettings_();
    const isT = u.role === 'teacher', bookable = win.state === 'open' || isT;
    const all = getActiveRes_().filter(r => dates.indexOf(r.date) >= 0);
    const tmap = isT ? teamIndexMap_(all.map(r => r.team)) : null;

    const days = dates.map(d => {
      const dow = dowOf_(d);
      const periods = PERIODS.map(p => {
        const time = slotTime_(p, dow), ph = holidayFor_(hols, d, p);
        const slots = places.map(pl => {
          const r = all.find(x => x.date === d && x.place === pl.id && x.period === p);
          const s = { date: d, place: pl.id, placeName: pl.name, period: p, time: time, status: 'free', team: '', name: '', resId: '', canCancel: false, reason: '' };
          if (r) {
            s.resId = r.id;
            if (!isT && r.uid === u.id) { s.status = 'mine'; s.name = r.name; s.team = r.team; }
            else if (!isT && u.team && r.team === u.team) { s.status = 'team'; s.name = r.name; s.team = r.team; }
            else { s.status = 'other'; if (isT) { s.name = r.name; s.team = r.team; s.ci = tmap[r.team] % TEAM_COLORS; } }   // 학생에게는 다른 팀 이름을 보내지 않음
            s.canCancel = isT || (r.uid === u.id && win.state === 'open');
          } else if (ph !== null) { s.status = 'holiday'; s.reason = ph; }
          else if (!bookable) s.status = 'closed';
          return s;
        });
        return { name: p, time: time, slots: slots };
      });
      return { date: d, label: label_(d), dow: dow, holiday: allDayHoliday_(hols, d), periods: periods };
    });

    let team = null;
    if (!isT) {
      const mine = all.filter(r => r.team && r.team === u.team);
      team = {
        name: u.team,
        meters: PERIODS.map(p => ({ period: p, n: mine.filter(r => r.period === p).length, max: settings.limits[p] })),
        members: getUsers_().filter(x => x.role === 'student' && x.team === u.team).map(x => x.name),
        list: mine.sort((a, b) => a.date.localeCompare(b.date) || periodIdx_(a.period) - periodIdx_(b.period))
          .map(r => ({ label: label_(r.date), period: r.period, placeName: placeName_(r.place), time: slotTime_(r.period, dowOf_(r.date)), name: r.name, mine: r.uid === u.id }))
      };
    }
    if (team) {
      const viol = u.team ? weekViolations_(u.team, addDays_(monday, -7)) : [];
      team.block = { blocked: viol.length > 0, list: viol.map(v => ({ label: v.label, period: v.period, placeName: v.placeName })) };
    }
    return {
      ok: true, monday: monday, range: label_(dates[0]) + ' ~ ' + label_(dates[4]),
      days: days, team: team, bookable: bookable, isTeacher: isT,
      window: win, nowMs: Date.now(), check: checkInfo_()
    };
  }, ['teacher', 'student']);
}

function getMine(token) {
  return api_(token, u => {
    const from = addDays_(todayKst_(), -28);
    const list = getActiveRes_().filter(r => r.uid === u.id && r.date >= from)
      .sort((a, b) => a.date.localeCompare(b.date) || periodIdx_(a.period) - periodIdx_(b.period))
      .map(r => ({
        resId: r.id, date: r.date, label: label_(r.date), period: r.period, placeName: placeName_(r.place),
        time: slotTime_(r.period, dowOf_(r.date)),
        canCancel: windowOf_(mondayOf_(r.date)).state === 'open',
        past: r.date < todayKst_()
      }));
    return { ok: true, list: list };
  }, ['student']);
}

// ===================== 예약 / 취소 =====================
function bookSlot(token, date, place, period, asId) {
  return api_(token, u => {
    date = normDate_(date);
    if (!date || [1, 2, 3, 4, 5].indexOf(dowOf_(date)) < 0) throw new UserError('평일(월~금)만 예약할 수 있습니다.');
    const pl = getPlaces_(false).find(x => x.id === place);
    if (!pl || PERIODS.indexOf(period) < 0) throw new UserError('잘못된 요청입니다.');
    const isT = u.role === 'teacher';
    let who = u;
    if (isT) {
      if (!asId) throw new UserError('교사 계정은 학생을 선택해서 예약해 주세요.');
      who = getUsers_().find(x => x.id === asId && x.role === 'student');
      if (!who) throw new UserError('학생을 찾을 수 없습니다.');
    }
    if (!who.team) throw new UserError('팀이 지정되지 않았습니다. 선생님께 문의하세요.', '예약 불가');

    const hol = holidayFor_(getHolidays_(), date, period);
    if (hol !== null) throw new UserError(label_(date) + ' ' + period + '은(는) 예약할 수 없습니다. (' + hol + ')', '예약 불가');

    const monday = mondayOf_(date), win = windowOf_(monday);
    if (!isT) {
      if (win.state === 'before') throw new UserError('아직 예약이 열리지 않았습니다.\n오픈: ' + win.openLabel, '예약 불가');
      if (win.state === 'closed') throw new UserError('예약이 마감되었습니다.\n마감: ' + win.closeLabel, '예약 불가');
    }
    const limit = getSettings_().limits[period];
    if (!isT) {                                    // 지난주 퇴실 확인 미제출 팀은 이번 주 예약 불가
      const viol = weekViolations_(who.team, addDays_(monday, -7));
      if (viol.length) {
        throw new UserError(who.team + ' 팀은 지난주 퇴실 확인 사진을 제출하지 않아 이번 예약 기간에는 예약할 수 없습니다.\n\n[미제출 내역]\n' +
          viol.map(v => '• ' + v.label + ' ' + v.period + ' ' + v.placeName).join('\n') + '\n\n선생님께 문의하세요.', '예약 불가');
      }
    }

    return withLock_(() => {
      const all = getActiveRes_(), dates = weekDates_(monday);
      const taken = all.find(r => r.date === date && r.place === place && r.period === period);
      if (taken) {
        throw new UserError(taken.team === who.team
          ? '우리 팀의 ' + taken.name + '님이 이미 예약한 자리입니다.'
          : '이미 다른 팀이 예약한 자리입니다.', '예약 불가');
      }
      const teamWeek = all.filter(r => r.team === who.team && dates.indexOf(r.date) >= 0);

      const dup = teamWeek.find(r => r.date === date && r.period === period);
      if (dup) {
        throw new UserError('같은 팀의 ' + dup.name + '님이 이미 ' + label_(date) + ' ' + period + '에 ' + placeName_(dup.place) + '을(를) 예약했습니다.\n한 팀은 같은 날 같은 시간대에 한 곳만 예약할 수 있습니다.', '예약 불가');
      }
      const cnt = teamWeek.filter(r => r.period === period).length;
      if (cnt >= limit) {
        throw new UserError(who.team + ' 팀은 이번 주 ' + period + ' 시간대를 이미 ' + cnt + '개 예약했습니다. (최대 ' + limit + '개)\n\n[우리 팀 예약 내역]\n' + listText_(teamWeek), '예약 불가');
      }
      const sh = sheet_(SH.RES), row = sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, 10).setNumberFormat('@')
        .setValues([[newId_(), date, place, period, who.id, who.name, who.team, nowStamp_(), '예약', '']]);
      sh.getRange(row, 11).setFormula(placeFormula_(row));
      log_(u.id, isT ? '교사대리예약' : '예약', who.team + '/' + who.name + ' ' + date + ' ' + period + ' ' + pl.name);
      return { ok: true };
    });
  }, ['teacher', 'student']);
}

function cancelRes(token, resId) {
  return api_(token, u => withLock_(() => {
    const r = getActiveRes_().find(x => x.id === resId);
    if (!r) throw new UserError('이미 취소되었거나 존재하지 않는 예약입니다.');
    const isT = u.role === 'teacher';
    if (!isT && r.uid !== u.id) throw new UserError('본인이 한 예약만 취소할 수 있습니다.', '취소 불가');
    if (!isT && windowOf_(mondayOf_(r.date)).state !== 'open') {
      throw new UserError('예약·취소 가능 기간(전주 토요일 08:00 ~ 일요일 18:00)이 아닙니다.', '취소 불가');
    }
    sheet_(SH.RES).getRange(r.row, 9, 1, 2).setNumberFormat('@').setValues([['취소', nowStamp_()]]);
    log_(u.id, isT ? '교사취소' : '취소', r.team + '/' + r.name + ' ' + r.date + ' ' + r.period + ' ' + placeName_(r.place));
    return { ok: true };
  }), ['teacher', 'student']);
}

// ===================== 교사 / 매니저 기능 =====================
function listStudents(token) {
  return api_(token, () => ({
    ok: true,
    list: getUsers_().filter(x => x.role === 'student' && x.team)
      .sort((a, b) => a.team.localeCompare(b.team, 'ko') || a.name.localeCompare(b.name, 'ko'))
      .map(x => ({ id: x.id, name: x.name, team: x.team, label: x.team + ' · ' + x.name }))
  }), ['teacher']);
}

/** 예약 달력(교사, 매니저 조회 가능) */
function getReport(token, monday) {
  return api_(token, u => {
    if (!u.report) throw new UserError('권한이 없습니다.');
    monday = mondayOf_(normDate_(monday) || defaultReportMonday_());
    const dates = weekDates_(monday), hols = getHolidays_(), win = windowOf_(monday), places = activePlaces_();
    const all = getActiveRes_().filter(r => dates.indexOf(r.date) >= 0);
    const tmap = teamIndexMap_(all.map(r => r.team));
    const rows = [];
    PERIODS.forEach(p => places.forEach(pl => {
      const t1 = slotTime_(p, 1), t2 = slotTime_(p, 2);       // 월 기준 / 화 기준 (방과후는 요일별로 다름)
      rows.push({
        period: p, placeName: pl.name,
        timeLines: t1 === t2 ? [t1] : [t1, '(화·목 ' + t2 + ')'],
        cells: dates.map(d => {
          const r = all.find(x => x.date === d && x.place === pl.id && x.period === p);
          return {
            time: slotTime_(p, dowOf_(d)), team: r ? r.team : null, name: r ? r.name : null, uid: r ? r.uid : null,
            ci: r ? tmap[r.team] % TEAM_COLORS : null,
            holiday: holidayFor_(hols, d, p)
          };
        })
      });
    }));
    const list = all.slice().sort((a, b) => a.date.localeCompare(b.date) || periodIdx_(a.period) - periodIdx_(b.period) || placeName_(a.place).localeCompare(placeName_(b.place), 'ko'))
      .map(r => ({ date: r.date, label: label_(r.date), dow: DOW[dowOf_(r.date)], period: r.period, placeName: placeName_(r.place), time: slotTime_(r.period, dowOf_(r.date)), team: r.team, name: r.name, uid: r.uid, at: r.at }));
    const teams = {};
    getUsers_().filter(x => x.role === 'student' && x.team).forEach(x => { teams[x.team] = true; });
    all.forEach(r => { teams[r.team] = true; });
    const summary = Object.keys(teams).sort((a, b) => a.localeCompare(b, 'ko')).map(t => {
      const m = all.filter(r => r.team === t), c = {};
      PERIODS.forEach(p => { c[p] = m.filter(r => r.period === p).length; });
      return { team: t, counts: c, total: m.length, ci: tmap[t] % TEAM_COLORS };
    });
    return {
      ok: true, monday: monday, range: label_(dates[0]) + ' ~ ' + label_(dates[4]),
      rangeFull: dates[0].replace(/-/g, '.') + ' ~ ' + dates[4].replace(/-/g, '.'),
      title: getSettings_().title, periods: PERIODS,
      days: dates.map(d => ({ date: d, label: label_(d), holiday: allDayHoliday_(hols, d) })),
      rows: rows, list: list, summary: summary, windowClosed: win.state === 'closed', closeLabel: win.closeLabel,
      printedAt: nowStamp_()
    };
  });
}

function listHolidays(token) {
  return api_(token, () => ({
    ok: true,
    list: getHolidays_().sort((a, b) => a.date.localeCompare(b.date) || a.period.localeCompare(b.period))
      .map(h => ({ date: h.date, label: label_(h.date), period: h.period, reason: h.reason }))
  }), ['teacher']);
}
function setHoliday(token, date, reason, period, remove) {
  return api_(token, u => withLock_(() => {
    date = normDate_(date);
    if (!date) throw new UserError('날짜를 입력하세요.');
    period = periodOfCell_(period);
    const sh = sheet_(SH.HOL);
    const hit = getHolidays_().find(h => h.date === date && h.period === period);
    if (remove) {
      if (hit) sh.deleteRow(hit.row);
    } else if (hit) {
      sh.getRange(hit.row, 2).setValue(String(reason || '휴무'));
    } else {
      sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setNumberFormat('@').setValues([[date, String(reason || '휴무'), period]]);
    }
    log_(u.id, remove ? '휴무일삭제' : '휴무일등록', date + ' ' + period + ' ' + (reason || ''));
    return { ok: true };
  }), ['teacher']);
}


// ===================== 퇴실 확인 (아침·방과후 사진 제출) =====================
function needsCheckout_(period) { return period === '아침' || period === '방과후'; }
function hm_(s) { const m = /^(\d{1,2}):(\d{2})$/.exec(String(s).trim()); return m ? { h: +m[1], m: +m[2] } : null; }
function kstDateHM_(ymd, h, m) { return new Date(ymd + 'T' + pad_(h) + ':' + pad_(m) + ':00+09:00'); }

/** 안내문에 쓸 마감 시각 정보 */
function checkInfo_() {
  const st = getSettings_();
  const f = t => pad_(t.h) + ':' + pad_(t.m);
  return { enabled: !!st.checkoutStart, start: st.checkoutStart, cam: st.cam,
    deadlines: { '아침': f(st.deadline['아침']), '방과후': f(st.deadline['방과후']) } };
}
/** 제출 가능 시간: 예약한 시간 시작 ~ 당일 마감 */
function checkoutWindow_(date, period) {
  const startStr = slotTime_(period, dowOf_(date)).split('~')[0], s = hm_(startStr), dl = getSettings_().deadline[period];
  return { start: kstDateHM_(date, s.h, s.m), end: kstDateHM_(date, dl.h, dl.m), startLabel: startStr, endLabel: pad_(dl.h) + ':' + pad_(dl.m) };
}
function placeFormulaChk_(row) {
  return "=IFERROR(VLOOKUP(D" + row + ",'" + SH.PLACE + "'!A:B,2,FALSE),D" + row + ")";
}
function getCheckoutRows_() {
  const sh = ss_().getSheetByName(SH.CHK), map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 13).getDisplayValues().forEach((r, i) => {
    if (!r[0]) return;
    map[r[0]] = { row: i + 2, resId: r[0], date: normDate_(r[1]), period: r[2], place: r[3], team: r[4], uid: r[5], name: r[6],
      at: r[7], method: r[8], link: r[9], fileId: r[10], exempt: String(r[11]).trim() !== '', memo: r[12] };
  });
  return map;
}
function getPhotoFolder_() {
  const p = PropertiesService.getScriptProperties(), id = p.getProperty('PHOTO_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* 폴더가 없으면 새로 만듦 */ } }
  const f = DriveApp.createFolder('퇴실확인 사진 (' + getSettings_().title + ')');
  p.setProperty('PHOTO_FOLDER', f.getId());
  return f;
}

/** 예약 1건의 퇴실 확인 상태: wait(시작 전) / open(제출 가능) / done / late(미제출=위반) / excused(면제) */
function checkoutStatus_(r, sub, hols, now) {
  const w = checkoutWindow_(r.date, r.period);
  const base = {
    resId: r.id, date: r.date, label: label_(r.date), period: r.period, place: r.place, placeName: placeName_(r.place),
    time: slotTime_(r.period, dowOf_(r.date)), team: r.team, reserver: r.name,
    startLabel: w.startLabel, endLabel: w.endLabel, startMs: w.start.getTime(), endMs: w.end.getTime(),
    st: '', by: '', at: '', method: '', reason: '', hasPhoto: false
  };
  if (sub && sub.fileId) { base.st = 'done'; base.by = sub.name; base.at = sub.at; base.method = sub.method; base.hasPhoto = true; return base; }
  if (sub && sub.exempt) { base.st = 'excused'; base.reason = sub.memo || '교사 면제'; return base; }
  const hol = holidayFor_(hols, r.date, r.period);
  if (hol !== null) { base.st = 'excused'; base.reason = '휴무일 (' + hol + ')'; return base; }
  const t = now.getTime();
  base.st = t < w.start.getTime() ? 'wait' : (t <= w.end.getTime() ? 'open' : 'late');
  return base;
}
function sortItems_(a, b) { return a.date.localeCompare(b.date) || periodIdx_(a.period) - periodIdx_(b.period) || a.placeName.localeCompare(b.placeName, 'ko'); }
/** 기간(dates) 안에서 퇴실 확인이 필요한 예약들의 상태 (team 이 비어 있으면 전체 팀) */
function checkoutItems_(team, dates) {
  const start = getSettings_().checkoutStart;
  if (!start) return [];
  const subs = getCheckoutRows_(), hols = getHolidays_(), now = new Date();
  return getActiveRes_().filter(r => (!team || r.team === team) && dates.indexOf(r.date) >= 0 && r.date >= start && needsCheckout_(r.period))
    .map(r => checkoutStatus_(r, subs[r.id], hols, now)).sort(sortItems_);
}
function weekViolations_(team, monday) {
  if (!team) return [];
  return checkoutItems_(team, weekDates_(monday)).filter(x => x.st === 'late');
}
function pubItem_(x) {
  return { resId: x.resId, date: x.date, label: x.label, period: x.period, placeName: x.placeName, time: x.time, team: x.team, reserver: x.reserver,
    startLabel: x.startLabel, endLabel: x.endLabel, startMs: x.startMs, endMs: x.endMs, st: x.st, by: x.by, at: x.at, method: x.method, reason: x.reason, hasPhoto: x.hasPhoto };
}

/** 학생: 우리 팀의 아침·방과후 예약과 제출 상태 */
function getCheckout(token, monday) {
  return api_(token, u => {
    if (!u.team) throw new UserError('팀이 지정되지 않았습니다. 선생님께 문의하세요.');
    monday = mondayOf_(normDate_(monday) || todayKst_());
    const dates = weekDates_(monday), items = checkoutItems_(u.team, dates);
    return {
      ok: true, monday: monday, range: label_(dates[0]) + ' ~ ' + label_(dates[4]), team: u.team,
      items: items.map(pubItem_), nextBlock: items.filter(x => x.st === 'late').map(pubItem_),
      check: checkInfo_(), nowMs: Date.now()
    };
  }, ['student']);
}

/** 학생: 퇴실 확인 사진 제출 (우리 팀 예약이면 팀원 누구나 가능) */
function submitCheckout(token, resId, photo, meta) {
  return api_(token, u => {
    const settings = getSettings_();
    if (!settings.checkoutStart) throw new UserError('퇴실 확인 기능이 꺼져 있습니다.');
    meta = meta || {}; const ck = meta.checks || {};
    if (!(ck.light && ck.heat && ck.lock)) throw new UserError('전등 OFF, 난방 OFF, 문단속을 모두 확인하고 체크해 주세요.', '제출 불가');
    const method = meta.method === 'live' ? 'live' : 'app';
    if (settings.cam === 'live' && method !== 'live') throw new UserError('이 앱에서는 앱 안의 카메라로 실시간 촬영한 사진만 제출할 수 있습니다.', '제출 불가');
    if (method === 'app') {
      const age = Date.now() - Number(meta.lastModified || 0);
      if (!meta.lastModified || age > 10 * 60 * 1000 || age < -5 * 60 * 1000) throw new UserError('방금 찍은 사진만 제출할 수 있어요. 카메라로 다시 촬영해 주세요.', '제출 불가');
    }
    const bytes = Utilities.base64Decode(String(photo || '').replace(/^data:image\/jpeg;base64,/, ''));
    if (bytes.length < 5000 || bytes.length > 4 * 1024 * 1024 || (bytes[0] & 255) !== 0xFF || (bytes[1] & 255) !== 0xD8) {
      throw new UserError('사진을 확인할 수 없습니다. 다시 촬영해 주세요.', '제출 불가');
    }
    return withLock_(() => {
      const res = getActiveRes_().find(x => x.id === resId);
      if (!res) throw new UserError('예약을 찾을 수 없습니다.');
      if (res.team !== u.team) throw new UserError('우리 팀의 예약만 제출할 수 있어요.', '제출 불가');
      if (!needsCheckout_(res.period)) throw new UserError('점심시간은 퇴실 확인 사진을 내지 않아도 돼요.');
      if (res.date < settings.checkoutStart) throw new UserError('이 예약은 퇴실 확인 대상이 아닙니다.');
      const subs = getCheckoutRows_(), old = subs[resId];
      if (old && old.fileId) throw new UserError('이미 제출된 예약입니다. (' + old.name + ', ' + old.at + ')', '제출 불가');
      const w = checkoutWindow_(res.date, res.period), now = Date.now();
      if (now < w.start.getTime()) throw new UserError('아직 제출 시간이 아닙니다.\n제출 가능: ' + label_(res.date) + ' ' + w.startLabel + ' ~ ' + w.endLabel, '제출 불가');
      if (now > w.end.getTime()) throw new UserError('제출 시간이 지났습니다. (마감 ' + w.endLabel + ')\n선생님께 문의하세요.', '제출 불가');
      const stamp = nowStamp_();
      const fname = [res.date, res.period, placeName_(res.place), res.team, u.name, stamp.slice(11, 16).replace(':', '')].join('_').replace(/[\\/:*?"<>|]/g, '') + '.jpg';
      const file = getPhotoFolder_().createFile(Utilities.newBlob(bytes, 'image/jpeg', fname));
      const sh = sheet_(SH.CHK), row = old ? old.row : sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, 13).setNumberFormat('@').setValues([[res.id, res.date, res.period, res.place, res.team, u.id, u.name, stamp, method, file.getUrl(), file.getId(), '', old ? old.memo : '']]);
      sh.getRange(row, 14).setFormula(placeFormulaChk_(row));
      log_(u.id, '퇴실확인제출', res.team + ' ' + res.date + ' ' + res.period + ' ' + placeName_(res.place) + ' (' + method + ')');
      return { ok: true, at: stamp };
    });
  }, ['student']);
}

/** 교사: 한 주의 모든 팀 퇴실 확인 현황 */
function getCheckoutAdmin(token, monday) {
  return api_(token, () => {
    monday = mondayOf_(normDate_(monday) || todayKst_());
    const dates = weekDates_(monday), items = checkoutItems_('', dates);
    const teams = {};
    items.forEach(x => {
      const t = teams[x.team] = teams[x.team] || { team: x.team, total: 0, done: 0, late: 0, excused: 0 };
      t.total++; if (x.st === 'done') t.done++; else if (x.st === 'late') t.late++; else if (x.st === 'excused') t.excused++;
    });
    return {
      ok: true, monday: monday, range: label_(dates[0]) + ' ~ ' + label_(dates[4]),
      items: items.map(pubItem_), teams: Object.keys(teams).sort((a, b) => a.localeCompare(b, 'ko')).map(k => teams[k]),
      blockedTeams: Object.keys(teams).filter(k => teams[k].late > 0).sort((a, b) => a.localeCompare(b, 'ko')),
      nextWeekLabel: label_(addDays_(monday, 7)), check: checkInfo_(), nowMs: Date.now()
    };
  }, ['teacher']);
}

/** 교사: 제출된 사진 보기 */
function getPhoto(token, resId) {
  return api_(token, () => {
    const sub = getCheckoutRows_()[resId];
    if (!sub || !sub.fileId) throw new UserError('제출된 사진이 없습니다.');
    const blob = DriveApp.getFileById(sub.fileId).getBlob();
    return { ok: true, dataUrl: 'data:image/jpeg;base64,' + Utilities.base64Encode(blob.getBytes()), by: sub.name, at: sub.at, method: sub.method, link: sub.link };
  }, ['teacher']);
}

/** 교사: 미제출 면제(해제) */
function setExempt(token, resId, on, memo) {
  return api_(token, u => withLock_(() => {
    const res = getRes_().find(x => x.id === resId);
    if (!res) throw new UserError('예약을 찾을 수 없습니다.');
    const sh = sheet_(SH.CHK), old = getCheckoutRows_()[resId];
    if (on) {
      if (old && old.fileId) throw new UserError('이미 사진이 제출된 예약입니다.');
      const row = old ? old.row : sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, 13).setNumberFormat('@').setValues([[res.id, res.date, res.period, res.place, res.team, '', '', '', '', '', '', '면제', String(memo || '교사 면제')]]);
      sh.getRange(row, 14).setFormula(placeFormulaChk_(row));
    } else if (old && !old.fileId) {
      sh.deleteRow(old.row);
    } else if (old) {
      sh.getRange(old.row, 12, 1, 2).setNumberFormat('@').setValues([['', '']]);
    }
    log_(u.id, on ? '퇴실면제' : '퇴실면제해제', res.team + ' ' + res.date + ' ' + res.period);
    return { ok: true };
  }), ['teacher']);
}
