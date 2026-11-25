/**
 * 동아리실·요가실 예약 시스템 (Google Apps Script)
 * - 모든 시간은 대한민국 표준시(Asia/Seoul) 기준
 * - 스프레드시트에 연결된(컨테이너 바운드) 스크립트로 사용하세요.
 */

// ===================== 설정 =====================
const TZ = 'Asia/Seoul';
const SH = { USERS: '학생', RES: '예약', HOL: '휴무일', LOG: '로그' };
const PLACES = ['동아리실', '요가실'];
const PERIODS = ['점심', '방과후'];
const LIMIT = { lunch: 2, total: 4 };      // 팀당 '주당' 최대 예약 수
const OPEN_HOUR = 8;                        // 전주 토요일 08:00 오픈
const CLOSE_HOUR = 18;                      // 전주 일요일 18:00 마감
const SESSION_SEC = 21600;                  // 로그인 유지 6시간
const MAX_FAIL = 5;                         // 로그인 5회 실패 시
const LOCK_SEC = 600;                       // 10분 잠금
const DOW = ['일', '월', '화', '수', '목', '금', '토'];
const RES_HEAD = ['예약ID', '날짜', '장소', '시간대', '아이디', '이름', '팀명', '예약시각', '상태', '취소시각'];

class UserError extends Error {
  constructor(msg, title) { super(msg); this.userMsg = msg; this.title = title || '알림'; }
}

// ===================== 웹앱 진입 =====================
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('동아리실·요가실 예약')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

// ===================== 시트 메뉴 / 초기 설정 =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('예약시스템')
    .addItem('① 초기 설정 (시트 만들기)', 'setup')
    .addItem('② 비밀번호 암호화', 'hashAllPasswords')
    .addToUi();
}

function setup() {
  const ss = ss_();
  const mk = (name, head) => {
    let sh = ss.getSheetByName(name), isNew = false;
    if (!sh) { sh = ss.insertSheet(name); isNew = true; }
    if (isNew || sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8efff');
      sh.setFrozenRows(1);
    }
    return { sh, isNew };
  };
  const u = mk(SH.USERS, ['아이디', '비밀번호', '이름', '팀명', '역할']);
  u.sh.getRange('A:E').setNumberFormat('@');
  u.sh.getRange('E2:E300').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['학생', '교사'], true).build());
  if (u.isNew) {
    u.sh.getRange(2, 1, 3, 5).setValues([
      ['선생님', '1234', '기술가정 선생님', '', '교사'],
      ['민준', '1111', '김민준', '푸른팀', '학생'],
      ['하늘', '2222', '이하늘', '푸른팀', '학생']
    ]);
  }
  u.sh.setColumnWidths(1, 5, 120);

  const r = mk(SH.RES, RES_HEAD);
  r.sh.getRange('B:B').setNumberFormat('@');
  r.sh.getRange('H:H').setNumberFormat('@');
  r.sh.getRange('J:J').setNumberFormat('@');
  r.sh.setColumnWidths(1, 10, 110);

  const h = mk(SH.HOL, ['날짜', '사유']);
  h.sh.getRange('A:A').setNumberFormat('@');
  const l = mk(SH.LOG, ['시각', '아이디', '동작', '내용']);
  l.sh.getRange('A:A').setNumberFormat('@');

  const first = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  ss.setSpreadsheetTimeZone(TZ);
  SpreadsheetApp.getUi().alert('초기 설정이 끝났습니다.\n[학생] 시트에 아이디·비밀번호·이름·팀명·역할(학생/교사)을 입력하세요.');
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
function getUsers_() {
  const sh = sheet_(SH.USERS), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getDisplayValues().map((r, i) => ({
    row: i + 2,
    id: String(r[0]).trim(),
    pw: String(r[1]).trim(),
    name: String(r[2]).trim() || String(r[0]).trim(),
    team: String(r[3]).trim(),
    role: String(r[4]).indexOf('교사') >= 0 ? 'teacher' : 'student'
  })).filter(u => u.id);
}
function getRes_() {
  const sh = sheet_(SH.RES), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 10).getDisplayValues().map((r, i) => ({
    row: i + 2, id: r[0], date: normDate_(r[1]), place: r[2], period: r[3],
    uid: r[4], name: r[5], team: r[6], at: r[7], status: String(r[8]).trim()
  })).filter(r => r.id && r.date);
}
function getActiveRes_() { return getRes_().filter(r => r.status === '예약'); }
function getHolidays_() {
  const sh = sheet_(SH.HOL), last = sh.getLastRow(), map = {};
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 2).getDisplayValues().forEach(r => {
    const d = normDate_(r[0]); if (d) map[d] = String(r[1]).trim() || '휴무';
  });
  return map;
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
function pub_(u) { return { id: u.id, name: u.name, team: u.team, role: u.role }; }
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
function api_(token, fn, role) {
  try {
    const u = auth_(token);
    if (role && u.role !== role) throw new UserError('권한이 없습니다.');
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
function listText_(arr) {
  return arr.slice().sort((a, b) => a.date.localeCompare(b.date) || periodIdx_(a.period) - periodIdx_(b.period))
    .map(r => '• ' + r.name + ' — ' + label_(r.date) + ' ' + r.period + ' ' + r.place).join('\n');
}

function getWeek(token, monday) {
  return api_(token, u => {
    monday = mondayOf_(normDate_(monday) || todayKst_());
    const dates = weekDates_(monday), hol = getHolidays_(), win = windowOf_(monday);
    const isT = u.role === 'teacher', bookable = win.state === 'open' || isT;
    const all = getActiveRes_().filter(r => dates.indexOf(r.date) >= 0);

    const days = dates.map(d => {
      const holiday = hol[d] === undefined ? null : hol[d];
      const slots = [];
      PERIODS.forEach(p => PLACES.forEach(pl => {
        const r = all.find(x => x.date === d && x.place === pl && x.period === p);
        const s = { date: d, place: pl, period: p, time: slotTime_(p, dowOf_(d)), status: 'free', team: '', name: '', resId: '', canCancel: false };
        if (r) {
          s.resId = r.id; s.team = r.team;
          if (!isT && r.uid === u.id) { s.status = 'mine'; s.name = r.name; }
          else if (!isT && u.team && r.team === u.team) { s.status = 'team'; s.name = r.name; }
          else { s.status = 'other'; if (isT) s.name = r.name; }
          s.canCancel = isT || (r.uid === u.id && win.state === 'open');
        } else if (holiday !== null) s.status = 'holiday';
        else if (!bookable) s.status = 'closed';
        slots.push(s);
      }));
      return { date: d, label: label_(d), dow: dowOf_(d), holiday: holiday, slots: slots };
    });

    let team = null;
    if (!isT) {
      const mine = all.filter(r => r.team && r.team === u.team);
      team = {
        name: u.team,
        lunch: mine.filter(r => r.period === '점심').length,
        total: mine.length,
        limitLunch: LIMIT.lunch, limitTotal: LIMIT.total,
        members: getUsers_().filter(x => x.role === 'student' && x.team === u.team).map(x => x.name),
        list: mine.sort((a, b) => a.date.localeCompare(b.date) || periodIdx_(a.period) - periodIdx_(b.period))
          .map(r => ({ label: label_(r.date), period: r.period, place: r.place, time: slotTime_(r.period, dowOf_(r.date)), name: r.name, mine: r.uid === u.id }))
      };
    }
    return {
      ok: true, monday: monday, range: label_(dates[0]) + ' ~ ' + label_(dates[4]),
      days: days, team: team, bookable: bookable, isTeacher: isT,
      window: win, nowMs: Date.now()
    };
  });
}

function getMine(token) {
  return api_(token, u => {
    const from = addDays_(todayKst_(), -28);
    const list = getActiveRes_().filter(r => r.uid === u.id && r.date >= from)
      .sort((a, b) => a.date.localeCompare(b.date) || periodIdx_(a.period) - periodIdx_(b.period))
      .map(r => ({
        resId: r.id, date: r.date, label: label_(r.date), period: r.period, place: r.place,
        time: slotTime_(r.period, dowOf_(r.date)),
        canCancel: windowOf_(mondayOf_(r.date)).state === 'open',
        past: r.date < todayKst_()
      }));
    return { ok: true, list: list };
  }, 'student');
}

// ===================== 예약 / 취소 =====================
function bookSlot(token, date, place, period, asId) {
  return api_(token, u => {
    date = normDate_(date);
    if (!date || [1, 2, 3, 4, 5].indexOf(dowOf_(date)) < 0) throw new UserError('평일(월~금)만 예약할 수 있습니다.');
    if (PLACES.indexOf(place) < 0 || PERIODS.indexOf(period) < 0) throw new UserError('잘못된 요청입니다.');
    const isT = u.role === 'teacher';
    let who = u;
    if (isT) {
      if (!asId) throw new UserError('교사 계정은 학생을 선택해서 예약해 주세요.');
      who = getUsers_().find(x => x.id === asId && x.role === 'student');
      if (!who) throw new UserError('학생을 찾을 수 없습니다.');
    }
    if (!who.team) throw new UserError('팀이 지정되지 않았습니다. 선생님께 문의하세요.', '예약 불가');

    const hol = getHolidays_()[date];
    if (hol !== undefined) throw new UserError(label_(date) + '은(는) 예약할 수 없는 날입니다. (' + hol + ')', '예약 불가');

    const monday = mondayOf_(date), win = windowOf_(monday);
    if (!isT) {
      if (win.state === 'before') throw new UserError('아직 예약이 열리지 않았습니다.\n오픈: ' + win.openLabel, '예약 불가');
      if (win.state === 'closed') throw new UserError('예약이 마감되었습니다.\n마감: ' + win.closeLabel, '예약 불가');
    }

    return withLock_(() => {
      const all = getActiveRes_(), dates = weekDates_(monday);
      const taken = all.find(r => r.date === date && r.place === place && r.period === period);
      if (taken) {
        throw new UserError(taken.team === who.team
          ? '우리 팀의 ' + taken.name + '님이 이미 예약한 자리입니다.'
          : '이미 다른 팀(' + taken.team + ')이 예약한 자리입니다.', '예약 불가');
      }
      const teamWeek = all.filter(r => r.team === who.team && dates.indexOf(r.date) >= 0);

      const dup = teamWeek.find(r => r.date === date && r.period === period);
      if (dup) {
        throw new UserError('같은 팀의 ' + dup.name + '님이 이미 ' + label_(date) + ' ' + period + '에 ' + dup.place + '을(를) 예약했습니다.\n한 팀은 같은 날 같은 시간대에 한 곳만 예약할 수 있습니다.', '예약 불가');
      }
      const lunch = teamWeek.filter(r => r.period === '점심').length;
      if (period === '점심' && lunch >= LIMIT.lunch) {
        throw new UserError(who.team + ' 팀은 이번 주 점심시간을 이미 ' + lunch + '개 예약했습니다. (최대 ' + LIMIT.lunch + '개)\n\n[우리 팀 예약 내역]\n' + listText_(teamWeek), '예약 불가');
      }
      if (teamWeek.length >= LIMIT.total) {
        throw new UserError(who.team + ' 팀은 이번 주 이미 ' + teamWeek.length + '개를 예약했습니다. (점심+방과후 최대 ' + LIMIT.total + '개)\n\n[우리 팀 예약 내역]\n' + listText_(teamWeek), '예약 불가');
      }
      const sh = sheet_(SH.RES), row = sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, 10).setNumberFormat('@')
        .setValues([[newId_(), date, place, period, who.id, who.name, who.team, nowStamp_(), '예약', '']]);
      log_(u.id, isT ? '교사대리예약' : '예약', who.team + '/' + who.name + ' ' + date + ' ' + period + ' ' + place);
      return { ok: true };
    });
  });
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
    log_(u.id, isT ? '교사취소' : '취소', r.team + '/' + r.name + ' ' + r.date + ' ' + r.period + ' ' + r.place);
    return { ok: true };
  }));
}

// ===================== 교사 기능 =====================
function listStudents(token) {
  return api_(token, () => ({
    ok: true,
    list: getUsers_().filter(x => x.role === 'student' && x.team)
      .sort((a, b) => a.team.localeCompare(b.team, 'ko') || a.name.localeCompare(b.name, 'ko'))
      .map(x => ({ id: x.id, name: x.name, team: x.team }))
  }), 'teacher');
}

function getReport(token, monday) {
  return api_(token, () => {
    monday = mondayOf_(normDate_(monday) || defaultReportMonday_());
    const dates = weekDates_(monday), hol = getHolidays_(), win = windowOf_(monday);
    const all = getActiveRes_().filter(r => dates.indexOf(r.date) >= 0);
    const rows = [];
    PERIODS.forEach(p => PLACES.forEach(pl => {
      rows.push({
        period: p, place: pl,
        cells: dates.map(d => {
          const r = all.find(x => x.date === d && x.place === pl && x.period === p);
          return { time: slotTime_(p, dowOf_(d)), team: r ? r.team : null, name: r ? r.name : null, uid: r ? r.uid : null };
        })
      });
    }));
    const list = all.slice().sort((a, b) => a.date.localeCompare(b.date) || periodIdx_(a.period) - periodIdx_(b.period) || PLACES.indexOf(a.place) - PLACES.indexOf(b.place))
      .map(r => ({ date: r.date, label: label_(r.date), dow: DOW[dowOf_(r.date)], period: r.period, place: r.place, time: slotTime_(r.period, dowOf_(r.date)), team: r.team, name: r.name, uid: r.uid, at: r.at }));
    const teams = {};
    getUsers_().filter(x => x.role === 'student' && x.team).forEach(x => { teams[x.team] = true; });
    all.forEach(r => { teams[r.team] = true; });
    const summary = Object.keys(teams).sort((a, b) => a.localeCompare(b, 'ko')).map(t => {
      const m = all.filter(r => r.team === t);
      return { team: t, lunch: m.filter(r => r.period === '점심').length, after: m.filter(r => r.period === '방과후').length, total: m.length };
    });
    return {
      ok: true, monday: monday, range: label_(dates[0]) + ' ~ ' + label_(dates[4]),
      rangeFull: dates[0].replace(/-/g, '.') + ' ~ ' + dates[4].replace(/-/g, '.'),
      days: dates.map(d => ({ date: d, label: label_(d), holiday: hol[d] === undefined ? null : hol[d] })),
      rows: rows, list: list, summary: summary, windowClosed: win.state === 'closed', closeLabel: win.closeLabel,
      printedAt: nowStamp_()
    };
  }, 'teacher');
}

function listHolidays(token) {
  return api_(token, () => {
    const hol = getHolidays_();
    return { ok: true, list: Object.keys(hol).sort().map(d => ({ date: d, label: label_(d), reason: hol[d] })) };
  }, 'teacher');
}
function setHoliday(token, date, reason, remove) {
  return api_(token, u => withLock_(() => {
    date = normDate_(date);
    if (!date) throw new UserError('날짜를 입력하세요.');
    const sh = sheet_(SH.HOL), last = sh.getLastRow();
    const vals = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues() : [];
    const idx = vals.findIndex(r => normDate_(r[0]) === date);
    if (remove) {
      if (idx >= 0) sh.deleteRow(idx + 2);
    } else if (idx >= 0) {
      sh.getRange(idx + 2, 2).setValue(String(reason || '휴무'));
    } else {
      sh.getRange(last + 1, 1, 1, 2).setNumberFormat('@').setValues([[date, String(reason || '휴무')]]);
    }
    log_(u.id, remove ? '휴무일삭제' : '휴무일등록', date + ' ' + (reason || ''));
    return { ok: true };
  }), 'teacher');
}
