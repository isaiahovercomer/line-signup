/**
 * LINE 群組一鍵報名系統 — Google Apps Script 後端
 * 貼到 Google 試算表的「擴充功能 → Apps Script」。更新程式後執行一次 setup() 補齊欄位，再部署新版本。
 */

const CONFIG = {
  CHANNEL_ID: '2012412693', // LINE Developers → 你的 LINE Login channel → Basic settings
};

const SH = { ROSTER: '名單', EVENTS: '活動', LOG: '報名紀錄', PUBLISHERS: '發布者' };

const HEADERS = {
  '名單': ['姓名', '手機', 'LINE userId', 'LINE名稱'],
  '活動': ['活動代碼', '活動名稱', '日期時間', '地點', '說明', '圖片網址', '按鈕文字', '開放報名', '建立者', '建立時間', '人數上限', '截止時間'],
  '報名紀錄': ['時間', '活動代碼', '活動名稱', '姓名', 'LINE名稱', 'LINE userId'],
  '發布者': ['LINE userId', '名稱', '備註'],
};

// 「活動」工作表的欄位位置（從 0 開始）
const EV = { ID: 0, TITLE: 1, DATE: 2, PLACE: 3, DESC: 4, IMAGE: 5, BUTTON: 6, OPEN: 7, CREATOR: 8, CREATED: 9, CAP: 10, DEADLINE: 11 };

/** 建立工作表並補齊標題列；更新程式後再執行一次也安全 */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const h = HEADERS[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  ss.getSheetByName(SH.ROSTER).getRange('B:B').setNumberFormat('@');   // 手機保留開頭的 0
  ss.getSheetByName(SH.EVENTS).getRange('C:C').setNumberFormat('@');   // 日期時間存成文字
  ss.getSheetByName(SH.EVENTS).getRange('L:L').setNumberFormat('@');   // 截止時間存成文字
}

function doGet() {
  return json({ ok: true, msg: '報名系統運作中' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'event') {
      const ev = getEvent(req.eventId);
      out = { ok: true, event: ev ? withStats(ev) : null };
    } else {
      const user = verifyIdToken(req.idToken);
      const admin = () => requirePublisher(user);              // 只有「發布者」工作表裡的管理員
      const own = () => requireOwner(user, req.eventId);       // 活動建立者或管理員
      switch (req.action) {
        case 'register':     out = register(user, req.eventId); break;
        case 'bind':         out = bind(user, req.eventId, req.name, req.phone); break;
        case 'me':           out = { ok: true, userId: user.userId, displayName: user.displayName, isAdmin: isPublisher(user.userId) }; break;
        case 'createEvent':  out = { ok: true, event: createEvent(user, req.event || {}) }; break;
        case 'listEvents':   out = { ok: true, events: listEvents(user) }; break;
        case 'updateEvent':  own(); out = { ok: true, event: updateEvent(req.eventId, req.event || {}) }; break;
        case 'eventDetail':  own(); out = { ok: true, ...eventDetail(req.eventId) }; break;
        case 'setOpen':      own(); out = { ok: true, event: setOpen(req.eventId, !!req.open) }; break;
        case 'listMembers':  admin(); out = { ok: true, members: listMembers() }; break;
        case 'addMember':    admin(); out = { ok: true, members: addMember(req.name, req.phone) }; break;
        case 'unbindMember': admin(); out = { ok: true, members: unbindMember(req.row, req.name) }; break;
        case 'removeMember': admin(); out = { ok: true, members: removeMember(req.row, req.name) }; break;
        default: throw new Error('未知的動作');
      }
    }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return json(out);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 向 LINE 驗證 ID token，防止有人偽造身分 */
function verifyIdToken(idToken) {
  if (!idToken) throw new Error('缺少登入資訊，請在 LINE 中重新開啟');
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: idToken, client_id: CONFIG.CHANNEL_ID },
    muteHttpExceptions: true,
  });
  const d = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) {
    const m = d.error_description || '驗證失敗';
    throw new Error(/expired/i.test(m) ? 'TOKEN_EXPIRED' : 'LINE 身分驗證失敗：' + m);
  }
  return { userId: d.sub, displayName: d.name || '' };
}

/* ---------- 工具 ---------- */
function sheet(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function rows(name) { const v = sheet(name).getDataRange().getValues(); v.shift(); return v; }
function tz() { return Session.getScriptTimeZone(); }
function normPhone(v) { return String(v || '').replace(/\D/g, '').slice(-9); }
function isTrue(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function fmt(d, pattern) { return Utilities.formatDate(d, tz(), pattern || 'yyyy/MM/dd HH:mm'); }
function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** 截止時間統一存成「yyyy-MM-dd HH:mm」文字 */
function normDeadline(v) {
  if (v instanceof Date) return fmt(v, 'yyyy-MM-dd HH:mm');
  const s = String(v || '').trim().replace('T', ' ').slice(0, 16);
  return /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(s) ? s : '';
}
function deadlineDate(s) {
  return s ? Utilities.parseDate(s, tz(), 'yyyy-MM-dd HH:mm') : null;
}

function rowToEvent(r) {
  return {
    id: String(r[EV.ID]), title: String(r[EV.TITLE]),
    date: r[EV.DATE] instanceof Date ? fmt(r[EV.DATE]) : String(r[EV.DATE] || ''),
    place: String(r[EV.PLACE] || ''), desc: String(r[EV.DESC] || ''), image: String(r[EV.IMAGE] || ''),
    button: String(r[EV.BUTTON] || '我要報名'), open: isTrue(r[EV.OPEN]),
    capacity: parseInt(r[EV.CAP], 10) || 0, deadline: normDeadline(r[EV.DEADLINE]),
  };
}

/** 活動狀態：open 報名中、full 額滿、ended 已過截止時間、closed 主辦人停止 */
function stateOf(ev, count) {
  if (!ev.open) return 'closed';
  const dl = deadlineDate(ev.deadline);
  if (dl && new Date() > dl) return 'ended';
  if (ev.capacity && count >= ev.capacity) return 'full';
  return 'open';
}

function logOf(eventId) { return rows(SH.LOG).filter(r => String(r[1]) === String(eventId)); }

function withStats(ev, log) {
  const count = (log || logOf(ev.id)).length;
  return Object.assign({}, ev, { count, state: stateOf(ev, count) });
}

function cleanEventInput(ev) {
  const title = String(ev.title || '').trim();
  if (!title) throw new Error('請填寫活動名稱');
  return {
    title,
    date: String(ev.date || '').trim(),
    place: String(ev.place || '').trim(),
    desc: String(ev.desc || '').trim(),
    image: /^https:\/\//.test(ev.image || '') ? String(ev.image).trim() : '',
    button: String(ev.button || '我要報名').trim().slice(0, 20) || '我要報名',
    capacity: Math.max(0, parseInt(ev.capacity, 10) || 0),
    deadline: normDeadline(ev.deadline),
  };
}

/* ---------- 發布者 ---------- */
function isPublisher(uid) { return rows(SH.PUBLISHERS).some(r => String(r[0]).trim() === uid); }
function requirePublisher(user) { if (!isPublisher(user.userId)) throw new Error('只有管理員可以使用這個功能'); }

/** 「建立者」欄位存成「LINE名稱 (userId)」，從括號裡取出 userId */
function creatorOf(row) { const m = /\((U[0-9a-f]+)\)\s*$/.exec(String(row[EV.CREATOR] || '')); return m ? m[1] : ''; }
function requireOwner(user, eventId) {
  if (isPublisher(user.userId)) return;
  if (creatorOf(findEventRow(eventId).row) !== user.userId) throw new Error('只有活動建立者可以管理這個活動');
}

/* ---------- 活動 ---------- */
function getEvent(id) {
  const r = rows(SH.EVENTS).find(r => String(r[EV.ID]) === String(id));
  return r ? rowToEvent(r) : null;
}

function findEventRow(id) {
  const data = sheet(SH.EVENTS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (String(data[i][EV.ID]) === String(id)) return { index: i + 1, row: data[i] };
  throw new Error('找不到這個活動');
}

function createEvent(user, input) {
  const ev = cleanEventInput(input);
  const id = 'E' + fmt(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  const row = [id, ev.title, ev.date, ev.place, ev.desc, ev.image, ev.button, true,
    user.displayName + ' (' + user.userId + ')', new Date(), ev.capacity || '', ev.deadline];
  sheet(SH.EVENTS).appendRow(row);
  return withStats(rowToEvent(row), []);
}

function updateEvent(id, input) {
  const ev = cleanEventInput(input);
  return withLock(() => {
    const f = findEventRow(id);
    const sh = sheet(SH.EVENTS);
    sh.getRange(f.index, EV.TITLE + 1, 1, 6).setValues([[ev.title, ev.date, ev.place, ev.desc, ev.image, ev.button]]);
    sh.getRange(f.index, EV.CAP + 1, 1, 2).setValues([[ev.capacity || '', ev.deadline]]);
    return withStats(getEvent(id));
  });
}

/** 管理員看到全部活動，其他人只看到自己建立的 */
function listEvents(user) {
  const log = rows(SH.LOG);
  const counts = {};
  log.forEach(r => { counts[r[1]] = (counts[r[1]] || 0) + 1; });
  const all = isPublisher(user.userId);
  return rows(SH.EVENTS).filter(r => all || creatorOf(r) === user.userId).reverse().slice(0, 50).map(r => {
    const ev = rowToEvent(r);
    ev.creator = String(r[EV.CREATOR] || '').replace(/\s*\(U[0-9a-f]+\)\s*$/, '');
    const count = counts[ev.id] || 0;
    return Object.assign(ev, { count, state: stateOf(ev, count) });
  });
}

function eventDetail(id) {
  const ev = getEvent(id);
  if (!ev) throw new Error('找不到這個活動');
  const log = logOf(id);
  return {
    event: withStats(ev, log),
    attendees: log.map(r => ({
      time: r[0] instanceof Date ? fmt(r[0], 'M/d HH:mm') : String(r[0]),
      name: String(r[3]), lineName: String(r[4] || ''),
    })),
  };
}

function setOpen(id, open) {
  return withLock(() => {
    const f = findEventRow(id);
    sheet(SH.EVENTS).getRange(f.index, EV.OPEN + 1).setValue(open);
    return withStats(getEvent(id));
  });
}

/* ---------- 成員名單 ---------- */
function listMembers() {
  return rows(SH.ROSTER).map((r, i) => ({
    row: i + 2, name: String(r[0] || ''), phone: String(r[1] || ''),
    bound: !!String(r[2] || '').trim(), lineName: String(r[3] || ''),
  })).filter(m => m.name);
}

function addMember(name, phone) {
  name = String(name || '').trim();
  const digits = String(phone || '').replace(/\D/g, '');
  if (!name) throw new Error('請填寫姓名');
  if (digits.length < 9) throw new Error('請填寫完整手機號碼');
  return withLock(() => {
    const dup = rows(SH.ROSTER).some(r => String(r[0]).trim() === name && normPhone(r[1]) === normPhone(digits));
    if (dup) throw new Error('名單裡已經有這個人');
    const sh = sheet(SH.ROSTER);
    const at = sh.getLastRow() + 1;
    sh.getRange(at, 2).setNumberFormat('@');
    sh.getRange(at, 1, 1, 4).setValues([[name, digits, '', '']]);
    return listMembers();
  });
}

/** row 與 name 必須同時符合，避免名單被別人改動時改錯人 */
function memberAt(row, name) {
  const sh = sheet(SH.ROSTER);
  row = parseInt(row, 10);
  if (!row || row < 2 || row > sh.getLastRow() || String(sh.getRange(row, 1).getValue()).trim() !== String(name).trim()) {
    throw new Error('名單已經變動，請重新整理後再試');
  }
  return { sh, row };
}

function unbindMember(row, name) {
  return withLock(() => {
    const m = memberAt(row, name);
    m.sh.getRange(m.row, 3, 1, 2).clearContent();
    return listMembers();
  });
}

function removeMember(row, name) {
  return withLock(() => {
    const m = memberAt(row, name);
    m.sh.deleteRow(m.row);
    return listMembers();
  });
}

/* ---------- 報名 ---------- */
function register(user, eventId) {
  return withLock(() => {
    const ev = getEvent(eventId);
    if (!ev) throw new Error('找不到這個活動');
    const log = logOf(ev.id);
    const names = () => log.map(r => String(r[3]));
    const state = stateOf(ev, log.length);

    const rs = sheet(SH.ROSTER);
    const data = rs.getDataRange().getValues();
    let idx = -1;
    for (let i = 1; i < data.length; i++) if (String(data[i][2]).trim() === user.userId) { idx = i; break; }
    if (idx < 0) {
      if (state !== 'open') return { ok: true, status: state, event: withStats(ev, log) };
      return { ok: true, status: 'need_bind', event: withStats(ev, log) };
    }

    const name = String(data[idx][0]);
    if (user.displayName && data[idx][3] !== user.displayName) rs.getRange(idx + 1, 4).setValue(user.displayName);

    if (log.some(r => String(r[5]) === user.userId)) {
      return { ok: true, status: 'already', name, event: withStats(ev, log), names: names() };
    }
    if (state !== 'open') return { ok: true, status: state, name, event: withStats(ev, log) };

    const entry = [new Date(), ev.id, ev.title, name, user.displayName, user.userId];
    sheet(SH.LOG).appendRow(entry);
    log.push(entry);
    return { ok: true, status: 'registered', name, event: withStats(ev, log), names: names() };
  });
}

/** 第一次報名：用姓名＋手機比對名單並綁定 LINE；名單上沒有就自動加入 */
function bind(user, eventId, name, phone) {
  name = String(name || '').trim();
  const p = normPhone(phone);
  if (!name || p.length < 9) throw new Error('請填寫姓名和完整手機號碼');

  const result = withLock(() => {
    const rs = sheet(SH.ROSTER);
    const data = rs.getDataRange().getValues();
    if (data.some((r, i) => i > 0 && String(r[2]).trim() === user.userId)) return null;
    let idx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === name && normPhone(data[i][1]) === p) { idx = i; break; }
    }
    if (idx < 0) {
      // 第一次報名、名單上還沒有這個人：自動加入名單並綁定
      const at = rs.getLastRow() + 1;
      rs.getRange(at, 2).setNumberFormat('@');
      rs.getRange(at, 1, 1, 4).setValues([[name, String(phone).replace(/\D/g, ''), user.userId, user.displayName]]);
      return null;
    }
    const bound = String(data[idx][2]).trim();
    if (bound && bound !== user.userId) return { ok: true, status: 'bound_other', event: getEvent(eventId) };
    rs.getRange(idx + 1, 3, 1, 2).setValues([[user.userId, user.displayName]]);
    return null;
  });
  return result || register(user, eventId);
}
