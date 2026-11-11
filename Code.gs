/**
 * LINE 介面活動登記建立平台 — Google Apps Script 後端
 * 貼到 Google 試算表的「擴充功能 → Apps Script」。更新程式後執行一次 setup() 補齊欄位，再部署新版本。
 */

const CONFIG = {
  CHANNEL_ID: '2012412693', // LINE Developers → 你的 LINE Login channel → Basic settings
  IMAGE_FOLDER: 'LINE 活動圖片', // 上傳的活動圖片會存在你雲端硬碟的這個資料夾
  LIFF_ID: '2012412693-a7phfbWa', // 提醒信裡的連結
};

const SH = { ROSTER: '名單', EVENTS: '活動', LOG: '報名紀錄', PUBLISHERS: '發布者' };

const HEADERS = {
  '名單': ['姓名', '手機', 'LINE userId', 'LINE名稱'],
  '活動': ['活動代碼', '活動名稱', '日期時間', '地點', '說明', '圖片網址', '按鈕文字', '開放報名', '建立者', '建立時間', '人數上限', '截止時間', '開始時間', '結束時間', '問題', '色調', '提醒時段', '已發提醒', '已寄通知'],
  '報名紀錄': ['時間', '活動代碼', '活動名稱', '姓名', 'LINE名稱', 'LINE userId', '回答', '回答資料'],
  '發布者': ['LINE userId', '名稱', '備註'],
};

// 「活動」工作表的欄位位置（從 0 開始）
const EV = { ID: 0, TITLE: 1, DATE: 2, PLACE: 3, DESC: 4, IMAGE: 5, BUTTON: 6, OPEN: 7, CREATOR: 8, CREATED: 9, CAP: 10, DEADLINE: 11, START: 12, END: 13, QUESTIONS: 14, THEME: 15, REMIND: 16, REMINDED: 17, NOTIFIED: 18 };

/** 建立工作表並補齊標題列；更新程式後再執行一次也安全（第一次會要求雲端硬碟權限，用來存上傳的圖片） */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const h = HEADERS[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  ss.getSheetByName(SH.ROSTER).getRange('B:B').setNumberFormat('@');   // 手機保留開頭的 0
  ['C:C', 'L:L', 'M:M', 'N:N'].forEach(c => ss.getSheetByName(SH.EVENTS).getRange(c).setNumberFormat('@')); // 時間存成文字
  imageFolder();
  ensureReminderTrigger();
}

function doGet() {
  return json({ ok: true, msg: '活動登記系統運作中' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'event') {
      const ev = getEvent(req.eventId);
      out = { ok: true, event: ev ? withStats(ev) : null };
    } else {
      // 有平台自己的通行證就用通行證（60 天有效），否則向 LINE 驗證 ID token
      const user = req.session && req.action !== 'session' ? userFromSession(req.session) : verifyIdToken(req.idToken);
      const admin = () => requirePublisher(user);              // 只有「發布者」工作表裡的管理員
      const own = () => requireOwner(user, req.eventId);       // 活動建立者或管理員
      switch (req.action) {
        case 'register':     out = register(user, req.eventId, req.answers); break;
        case 'bind':         out = bind(user, req.eventId, req.name, req.phone, req.answers); break;
        case 'cancel':       out = cancel(user, req.eventId); break;
        case 'updateAnswers': out = updateAnswers(user, req.eventId, req.answers); break;
        case 'myData':       out = { ok: true, ...myData(user) }; break;
        case 'deleteMyData': out = { ok: true, ...deleteMyData(user) }; break;
        case 'session':      out = { ok: true, ...createSession(user) }; break;
        case 'logout':       deleteSession(req.session); out = { ok: true }; break;
        case 'me':           out = { ok: true, userId: user.userId, displayName: user.displayName, isAdmin: isPublisher(user.userId) }; break;
        case 'uploadImage':  out = { ok: true, url: uploadImage(user, req.data, req.mime) }; break;
        case 'createEvent':  out = { ok: true, event: createEvent(user, req.event || {}) }; break;
        case 'listEvents':   out = { ok: true, events: listEvents(user), templates: listTemplates(user), notifyEmail: getNotifyEmail(user) }; break;
        case 'setNotifyEmail': out = { ok: true, notifyEmail: setNotifyEmail(user, req.email) }; break;
        case 'markReminded': own(); out = { ok: true, event: markReminded(req.eventId, req.offsets) }; break;
        case 'deleteTemplate': out = { ok: true, templates: deleteTemplate(user, req.title) }; break;
        case 'updateEvent':  own(); out = { ok: true, event: updateEvent(req.eventId, req.event || {}) }; break;
        case 'eventDetail':  own(); out = { ok: true, ...eventDetail(req.eventId) }; break;
        case 'setOpen':      own(); out = { ok: true, event: setOpen(req.eventId, !!req.open) }; break;
        case 'deleteEvent':  own(); out = { ok: true, removed: deleteEvent(req.eventId) }; break;
        case 'deleteEvents': out = { ok: true, ...deleteEvents(user, req.eventIds) }; break;
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

/* ---------- 平台通行證：登入一次，60 天內不用再登入 ----------
   通行證存在隱藏的「登入」工作表，只有試算表擁有者看得到。 */
const SESSION_SHEET = '登入';
const SESSION_DAYS = 60;

function sessionSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SESSION_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SESSION_SHEET);
    sh.appendRow(['通行證', 'LINE userId', 'LINE名稱', '到期', '建立']);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function createSession(user) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const now = new Date(), exp = new Date(now.getTime() + SESSION_DAYS * 864e5);
  withLock(() => {
    const sh = sessionSheet();
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) if (data[i][3] instanceof Date && data[i][3] < now) sh.deleteRow(i + 1); // 順手清掉過期的
    sh.appendRow([token, user.userId, user.displayName, exp, now]);
  });
  return { token, expires: exp.getTime() };
}

function userFromSession(token) {
  const data = sessionSheet().getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token && data[i][3] instanceof Date && data[i][3] > now) return { userId: String(data[i][1]), displayName: String(data[i][2]) };
  }
  throw new Error('SESSION_EXPIRED');
}

function deleteSession(token) {
  if (!token) return;
  withLock(() => {
    const sh = sessionSheet(), data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) if (data[i][0] === token) sh.deleteRow(i + 1);
  });
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
function parseJson(s, fallback) { try { return JSON.parse(s); } catch (e) { return fallback; } }
function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** 時間統一存成「yyyy-MM-dd HH:mm」文字 */
function normTime(v) {
  if (v instanceof Date) return fmt(v, 'yyyy-MM-dd HH:mm');
  const s = String(v || '').trim().replace('T', ' ').slice(0, 16);
  return /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(s) ? s : '';
}
function parseTime(s) { return s ? Utilities.parseDate(s, tz(), 'yyyy-MM-dd HH:mm') : null; }

/** 「11/20（五）09:00–12:00」；跨日時兩邊都寫日期 */
function dayLabel(d) { return fmt(d, 'M/d') + '（' + '一二三四五六日'.charAt(Number(fmt(d, 'u')) - 1) + '）'; }
function fmtRange(start, end) {
  const a = parseTime(start);
  if (!a) return '';
  let t = dayLabel(a) + ' ' + fmt(a, 'HH:mm');
  const b = parseTime(end);
  if (b) t += fmt(a, 'yyyyMMdd') === fmt(b, 'yyyyMMdd') ? '–' + fmt(b, 'HH:mm') : ' – ' + dayLabel(b) + ' ' + fmt(b, 'HH:mm');
  return t;
}

/** 問題存成 { q, options }；options 有值就是下拉選單。舊資料只有文字也能讀 */
function cleanQuestions(list) {
  return (Array.isArray(list) ? list : []).map(x => {
    const q = String(typeof x === 'string' ? x : (x && x.q) || '').trim().slice(0, 40);
    const options = [...new Set((x && Array.isArray(x.options) ? x.options : []).map(o => String(o || '').trim().slice(0, 30)).filter(Boolean))].slice(0, 60);
    return { q, options };
  }).filter(x => x.q).slice(0, 5);
}

/** 圖卡色調代碼（顏色定義在網頁那邊）；沒填或不認得就用經典色 */
const THEMES = ['classic', 'feast', 'prayer', 'gospel', 'youth', 'kids', 'study', 'service'];
function cleanTheme(v) { v = String(v || '').trim(); return THEMES.includes(v) ? v : 'classic'; }

/** 提醒時段：活動開始前幾小時（7 天、24 小時、1 小時），存成「168,24,1」 */
const REMIND_HOURS = [168, 24, 1];
function cleanRemind(v) {
  const list = Array.isArray(v) ? v : String(v || '').split(',');
  return REMIND_HOURS.filter(h => list.map(Number).includes(h));
}

function rowToEvent(r) {
  return {
    id: String(r[EV.ID]), title: String(r[EV.TITLE]),
    date: r[EV.DATE] instanceof Date ? fmt(r[EV.DATE]) : String(r[EV.DATE] || ''),
    place: String(r[EV.PLACE] || ''), desc: String(r[EV.DESC] || ''), image: String(r[EV.IMAGE] || ''),
    button: String(r[EV.BUTTON] || '我要登記'), open: isTrue(r[EV.OPEN]),
    capacity: parseInt(r[EV.CAP], 10) || 0, deadline: normTime(r[EV.DEADLINE]),
    start: normTime(r[EV.START]), end: normTime(r[EV.END]),
    questions: cleanQuestions(parseJson(r[EV.QUESTIONS], [])),
    theme: cleanTheme(r[EV.THEME]),
    remind: cleanRemind(r[EV.REMIND]), reminded: cleanRemind(r[EV.REMINDED]),
  };
}

/** 活動狀態：open 報名中、full 額滿、ended 已過截止時間、closed 主辦人停止 */
function stateOf(ev, count) {
  if (!ev.open) return 'closed';
  const dl = parseTime(ev.deadline);
  if (dl && new Date() > dl) return 'ended';
  if (ev.capacity && count >= ev.capacity) return 'full';
  return 'open';
}

function logOf(eventId) { return rows(SH.LOG).filter(r => String(r[1]) === String(eventId)); }

function withStats(ev, log) {
  log = log || logOf(ev.id);
  const count = log.length;
  return Object.assign({}, ev, { count, guests: guestsOf(ev, log), state: stateOf(ev, count) });
}

/* 「同行人數」「帶幾位家人」「兒童姓名」這類題目：換算成人數，加到總人數 */
const GUEST_RE = /同行|攜伴|攜眷|陪同|朋友|家人|眷屬|親友|配偶|兒童|小孩|孩子|子女|帶.{0,4}(人|誰|幾)|人數|幾位|幾人/;
const NONE_RE = /^(無|沒有|没有|沒|不|否|免|0|-|—|－|x|n\/?a|none|no)$/i;
const ZH_NUM = { 一: 1, 二: 2, 兩: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };
/** 一個回答代表幾個人：「3」「大人2小孩1」「兩位」「王小明、李大華」「無」 */
function countPeople(v) {
  v = String(v || '').trim().replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  if (!v || NONE_RE.test(v)) return 0;
  const nums = v.match(/\d+/g);
  if (nums) return Math.min(/歲/.test(v) ? (v.match(/歲/g) || []).length : nums.reduce((s, n) => s + +n, 0), 99); // 「小孩3歲、5歲」算 2 人
  const zh = v.match(/^([一二兩三四五六七八九十])\s*(位|個|人|名)?/);
  if (zh) return ZH_NUM[zh[1]];
  return Math.min(v.split(/[、,，\/;；\s]+|和|及|與|跟/).filter(x => x.trim()).length, 30); // 填姓名：有幾個名字就算幾人
}
/** 所有「同行／家人／兒童／人數」這類題目加起來；活動沒有這類題目回傳 null */
function guestsFrom(questions, answersList) {
  const idx = (questions || []).map((q, i) => GUEST_RE.test(String((q && q.q) || q)) ? i : -1).filter(i => i >= 0);
  if (!idx.length) return null;
  return answersList.reduce((s, ans) => s + idx.reduce((t, i) => t + countPeople((ans || [])[i]), 0), 0);
}
function guestsOf(ev, log) { return guestsFrom(ev.questions, log.map(answersOf)); }

function cleanEventInput(ev) {
  const title = String(ev.title || '').trim();
  if (!title) throw new Error('請填寫活動名稱');
  const start = normTime(ev.start);
  let end = normTime(ev.end);
  if (!start && end) throw new Error('有結束時間時，也要填開始時間');
  if (start && end && end <= start) throw new Error('結束時間要晚於開始時間');
  const questions = cleanQuestions(ev.questions);
  return {
    title, start, end,
    date: start ? fmtRange(start, end) : String(ev.date || '').trim(),
    place: String(ev.place || '').trim(),
    desc: String(ev.desc || '').trim(),
    image: /^https:\/\//.test(ev.image || '') ? String(ev.image).trim() : '',
    button: String(ev.button || '我要登記').trim().slice(0, 20) || '我要登記',
    capacity: Math.max(0, parseInt(ev.capacity, 10) || 0),
    deadline: normTime(ev.deadline),
    questions,
    theme: cleanTheme(ev.theme),
    remind: start ? cleanRemind(ev.remind) : [],
  };
}

/* ---------- 發布者與權限 ---------- */
function isPublisher(uid) { return rows(SH.PUBLISHERS).some(r => String(r[0]).trim() === uid); }
function requirePublisher(user) { if (!isPublisher(user.userId)) throw new Error('只有管理員可以使用這個功能'); }

/** 「建立者」欄位存成「LINE名稱 (userId)」，從括號裡取出 userId */
function creatorOf(row) { const m = /\((U[0-9a-f]+)\)\s*$/.exec(String(row[EV.CREATOR] || '')); return m ? m[1] : ''; }
function requireOwner(user, eventId) {
  if (isPublisher(user.userId)) return;
  if (creatorOf(findEventRow(eventId).row) !== user.userId) throw new Error('只有活動建立者可以管理這個活動');
}

/* ---------- 圖片上傳 ---------- */
function imageFolder() {
  const it = DriveApp.getFoldersByName(CONFIG.IMAGE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.IMAGE_FOLDER);
}

function uploadImage(user, data, mime) {
  mime = /^image\/(jpeg|png)$/.test(mime) ? mime : 'image/jpeg';
  const b64 = String(data || '').replace(/^data:[^,]+,/, '');
  if (!b64) throw new Error('沒有收到圖片');
  if (b64.length > 7000000) throw new Error('圖片太大，請換一張小一點的');
  const name = 'event-' + fmt(new Date(), 'yyyyMMdd-HHmmss') + (mime === 'image/png' ? '.png' : '.jpg');
  const file = imageFolder().createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime, name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // LINE 圖卡要能公開讀取
  file.setDescription('上傳者：' + user.displayName);
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}

/* ---------- 活動 ---------- */
function getEvent(id) {
  const r = rows(SH.EVENTS).find(r => String(r[EV.ID]) === String(id));
  return r ? rowToEvent(r) : null;
}

function findEventRow(id) {
  const data = sheet(SH.EVENTS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (String(data[i][EV.ID]) === String(id)) return { index: i + 1, row: data[i] };
  throw new Error('找不到這個活動，可能已經被刪除');
}

function createEvent(user, input) {
  const ev = cleanEventInput(input);
  const id = 'E' + fmt(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  const row = [id, ev.title, ev.date, ev.place, ev.desc, ev.image, ev.button, true,
    user.displayName + ' (' + user.userId + ')', new Date(), ev.capacity || '', ev.deadline,
    ev.start, ev.end, JSON.stringify(ev.questions), ev.theme, ev.remind.join(','), '', ''];
  sheet(SH.EVENTS).appendRow(row);
  return withStats(rowToEvent(row), []);
}

function updateEvent(id, input) {
  const ev = cleanEventInput(input);
  return withLock(() => {
    const f = findEventRow(id);
    const sh = sheet(SH.EVENTS);
    sh.getRange(f.index, EV.TITLE + 1, 1, 6).setValues([[ev.title, ev.date, ev.place, ev.desc, ev.image, ev.button]]);
    sh.getRange(f.index, EV.CAP + 1, 1, 6).setValues([[ev.capacity || '', ev.deadline, ev.start, ev.end, JSON.stringify(ev.questions), ev.theme]]);
    sh.getRange(f.index, EV.REMIND + 1).setValue(ev.remind.join(','));
    // 開始時間改了，提醒重新算
    if (normTime(f.row[EV.START]) !== ev.start) sh.getRange(f.index, EV.REMINDED + 1, 1, 2).setValues([['', '']]);
    return withStats(getEvent(id));
  });
}

/* ---------- 範本：活動刪除後仍保留內容（不含任何登記者資料），可一鍵套用 ----------
   存在隱藏的「範本」工作表，同一個建立者、同一個活動名稱只留最新一筆。 */
const TEMPLATE_SHEET = '範本';

function templateSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(TEMPLATE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TEMPLATE_SHEET);
    sh.appendRow(['建立者 userId', '活動名稱', '內容', '更新時間']);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function saveTemplates(eventRows) {
  if (!eventRows.length) return;
  const sh = templateSheet();
  const data = sh.getDataRange().getValues();
  eventRows.forEach(r => {
    const uid = creatorOf(r), ev = rowToEvent(r);
    if (!uid || !ev.title) return;
    const body = JSON.stringify({ title: ev.title, place: ev.place, desc: ev.desc, image: ev.image, button: ev.button,
      capacity: ev.capacity, questions: ev.questions, theme: ev.theme });
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === uid && String(data[i][1]) === ev.title) { sh.deleteRow(i + 1); data.splice(i, 1); }
    }
    sh.appendRow([uid, ev.title, body, new Date()]);
    data.push([uid, ev.title, body, new Date()]);
  });
}

function listTemplates(user) {
  templateSheet();
  return rows(TEMPLATE_SHEET)
    .filter(r => String(r[0]) === user.userId)
    .sort((a, b) => (b[3] instanceof Date ? b[3] : 0) - (a[3] instanceof Date ? a[3] : 0))
    .slice(0, 20)
    .map(r => Object.assign(parseJson(r[2], {}) || {}, { saved: true, title: String(r[1]) }));
}

function deleteTemplate(user, title) {
  withLock(() => {
    const sh = templateSheet(), data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]) === user.userId && String(data[i][1]) === String(title)) sh.deleteRow(i + 1);
  });
  return listTemplates(user);
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
    ev.mine = creatorOf(r) === user.userId; // 套用舊活動時只列自己建立的
    const count = counts[ev.id] || 0;
    return Object.assign(ev, { count, state: stateOf(ev, count) });
  });
}

function answersOf(r) { return parseJson(r[7], []) || []; }

/** 接龍上的一行：「簡晨光，三杯雞」 */
function rollLine(r) { return [String(r[3])].concat(answersOf(r).filter(Boolean)).join('，'); }

function eventDetail(id) {
  const ev = getEvent(id);
  if (!ev) throw new Error('找不到這個活動，可能已經被刪除');
  const log = logOf(id);
  return {
    event: withStats(ev, log),
    attendees: log.map(r => ({
      time: r[0] instanceof Date ? fmt(r[0], 'M/d HH:mm') : String(r[0]),
      name: String(r[3]), lineName: String(r[4] || ''), answers: answersOf(r),
    })),
  };
}

/** 一次刪除多個活動與它們的登記紀錄（無法復原）。先逐一確認權限，全部通過才刪 */
function deleteEvents(user, ids) {
  ids = [...new Set((Array.isArray(ids) ? ids : []).map(String))].slice(0, 200);
  if (!ids.length) throw new Error('沒有選取活動');
  ids.forEach(id => requireOwner(user, id));
  const want = new Set(ids);
  return withLock(() => {
    const delRows = (name, col) => {
      const sh = sheet(name), data = sh.getDataRange().getValues();
      let n = 0;
      for (let i = data.length - 1; i >= 1; i--) if (want.has(String(data[i][col]))) { sh.deleteRow(i + 1); n++; }
      return n;
    };
    saveTemplates(rows(SH.EVENTS).filter(r => want.has(String(r[EV.ID])))); // 先存成範本，之後還能一鍵套用
    const logs = delRows(SH.LOG, 1);
    const events = delRows(SH.EVENTS, EV.ID);
    return { events, logs };
  });
}

/** 刪除活動與它的報名紀錄（無法復原）；回傳刪掉幾筆報名 */
function deleteEvent(id) {
  return withLock(() => {
    const f = findEventRow(id);
    saveTemplates([f.row]); // 先存成範本，之後還能一鍵套用
    const log = sheet(SH.LOG);
    const data = log.getDataRange().getValues();
    let removed = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]) === String(id)) { log.deleteRow(i + 1); removed++; }
    }
    sheet(SH.EVENTS).deleteRow(f.index);
    return removed;
  });
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
/** answers 是依照活動問題順序的回答陣列；活動有問題但還沒收到回答時，回傳 need_answers 請前端顯示問題 */
function register(user, eventId, answers) {
  return withLock(() => {
    const ev = getEvent(eventId);
    if (!ev) throw new Error('找不到這個活動，可能已經被刪除');
    const log = logOf(ev.id);
    const lines = () => log.map(rollLine);
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

    const mine = log.find(r => String(r[5]) === user.userId);
    if (mine) return { ok: true, status: 'already', name, event: withStats(ev, log), names: lines(), answers: answersOf(mine) };
    if (state !== 'open') return { ok: true, status: state, name, event: withStats(ev, log) };
    if (ev.questions.length && !Array.isArray(answers)) return { ok: true, status: 'need_answers', name, event: withStats(ev, log) };

    const { ans, readable } = cleanAnswers(ev, answers);
    const entry = [new Date(), ev.id, ev.title, name, user.displayName, user.userId, readable, JSON.stringify(ans)];
    sheet(SH.LOG).appendRow(entry);
    log.push(entry);
    return { ok: true, status: 'registered', name, event: withStats(ev, log), names: lines() };
  });
}

/* ---------- 當事人查看與刪除自己的資料（個資法） ---------- */
function myData(user) {
  const r = rows(SH.ROSTER).find(r => String(r[2]).trim() === user.userId);
  const events = {};
  rows(SH.EVENTS).forEach(e => { events[String(e[EV.ID])] = String(e[EV.TITLE]); });
  const regs = rows(SH.LOG).filter(l => String(l[5]) === user.userId).map(l => ({
    title: events[String(l[1])] || String(l[2]),
    time: l[0] instanceof Date ? fmt(l[0], 'yyyy/M/d HH:mm') : String(l[0]),
  }));
  return {
    found: !!r || regs.length > 0,
    name: r ? String(r[0]) : '', phone: r ? String(r[1]) : '', lineName: r ? String(r[3] || '') : user.displayName,
    registrations: regs,
  };
}

/** 刪除名單資料與所有登記紀錄；回傳刪了幾筆 */
function deleteMyData(user) {
  return withLock(() => {
    let roster = 0, logs = 0;
    const del = (name, test) => {
      const sh = sheet(name), data = sh.getDataRange().getValues();
      let n = 0;
      for (let i = data.length - 1; i >= 1; i--) if (test(data[i])) { sh.deleteRow(i + 1); n++; }
      return n;
    };
    roster = del(SH.ROSTER, r => String(r[2]).trim() === user.userId);
    logs = del(SH.LOG, r => String(r[5]) === user.userId);
    sessionSheet(); del(SESSION_SHEET, r => String(r[1]) === user.userId); // 通行證也一併刪除
    templateSheet(); del(TEMPLATE_SHEET, r => String(r[0]) === user.userId); // 自己建立的範本也一併刪除
    notifySheet(); del(NOTIFY_SHEET, r => String(r[0]) === user.userId);      // 提醒通知信箱也一併刪除
    return { roster, logs };
  });
}

/** 依活動問題整理回答：下拉選單只收選項裡的答案；readable 是給試算表看的「我可以預備：三杯雞｜所屬大區：10大區」 */
function cleanAnswers(ev, answers) {
  const ans = ev.questions.map((q, i) => {
    const a = String((answers || [])[i] || '').trim().slice(0, 100);
    return q.options.length && !q.options.includes(a) ? '' : a;
  });
  const readable = ev.questions.map((q, i) => ans[i] ? q.q.replace(/[？?]\s*$/, '') + '：' + ans[i] : '').filter(Boolean).join('｜');
  return { ans, readable };
}

/** 修改自己的回答：報名順序與時間不變 */
function updateAnswers(user, eventId, answers) {
  return withLock(() => {
    const ev = getEvent(eventId);
    if (!ev) throw new Error('找不到這個活動，可能已經被刪除');
    const sh = sheet(SH.LOG);
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === ev.id && String(data[i][5]) === user.userId) {
        const { ans, readable } = cleanAnswers(ev, answers);
        sh.getRange(i + 1, 7, 1, 2).setValues([[readable, JSON.stringify(ans)]]);
        const log = logOf(ev.id);
        return { ok: true, status: 'updated', name: String(data[i][3]), event: withStats(ev, log), names: log.map(rollLine) };
      }
    }
    return { ok: true, status: 'not_registered', event: withStats(ev) };
  });
}

/** 取消自己的報名：刪掉報名紀錄，名額釋出 */
function cancel(user, eventId) {
  return withLock(() => {
    const ev = getEvent(eventId);
    if (!ev) throw new Error('找不到這個活動，可能已經被刪除');
    const sh = sheet(SH.LOG);
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]) === ev.id && String(data[i][5]) === user.userId) {
        sh.deleteRow(i + 1);
        const log = logOf(ev.id);
        return { ok: true, status: 'cancelled', name: String(data[i][3]), event: withStats(ev, log), names: log.map(rollLine) };
      }
    }
    return { ok: true, status: 'not_registered', event: withStats(ev) };
  });
}

/** 第一次報名：用姓名＋手機比對名單並綁定 LINE；名單上沒有就自動加入 */
function bind(user, eventId, name, phone, answers) {
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
  return result || register(user, eventId, answers);
}

/* ---------- 活動提醒 ----------
   個人 LINE 帳號不能由程式自動發到群組，所以做法是：
   到了提醒時間 → 寄 Email 通知建立者（有留信箱的話），平台上也會標示「該發提醒了」
   → 建立者按一下「發提醒到群組」，用自己的帳號把提醒卡發出去。
   remindTick 由時間觸發器每 15 分鐘執行一次（setup() 會自動建立）。 */
const NOTIFY_SHEET = '通知信箱';

function notifySheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(NOTIFY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(NOTIFY_SHEET);
    sh.appendRow(['LINE userId', 'Email', '更新時間']);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function getNotifyEmail(user) {
  notifySheet();
  const r = rows(NOTIFY_SHEET).find(r => String(r[0]) === user.userId);
  return r ? String(r[1]) : '';
}

function setNotifyEmail(user, email) {
  email = String(email || '').trim();
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email 格式不對，請再檢查一次');
  return withLock(() => {
    const sh = notifySheet(), data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]) === user.userId) sh.deleteRow(i + 1);
    if (email) sh.appendRow([user.userId, email, new Date()]);
    return email;
  });
}

/** 建立者按了「發提醒到群組」：記下已經發過的時段 */
function markReminded(id, offsets) {
  return withLock(() => {
    const f = findEventRow(id);
    const done = cleanRemind(f.row[EV.REMINDED]).concat(cleanRemind(offsets));
    sheet(SH.EVENTS).getRange(f.index, EV.REMINDED + 1).setValue(cleanRemind(done).join(','));
    return withStats(getEvent(id));
  });
}

/** 「還有 7 天」「明天 18:00 開始」「45 分鐘後開始」 */
function untilText(start) {
  const ms = start - new Date();
  if (ms < 90 * 60000) return Math.max(1, Math.round(ms / 60000)) + ' 分鐘後開始';
  const day = d => fmt(d, 'yyyyMMdd');
  const tomorrow = new Date(Date.now() + 86400000);
  if (day(start) === day(new Date())) return '今天 ' + fmt(start, 'HH:mm') + ' 開始';
  if (day(start) === day(tomorrow)) return '明天 ' + fmt(start, 'HH:mm') + ' 開始';
  if (ms < 48 * 3600000) return Math.round(ms / 3600000) + ' 小時後開始';
  return '還有 ' + Math.round(ms / 86400000) + ' 天';
}

function remindTick() {
  const sh = sheet(SH.EVENTS), data = sh.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    const r = data[i], ev = rowToEvent(r);
    const start = parseTime(ev.start);
    if (!start || start <= now || !ev.remind.length) continue;
    const due = ev.remind.filter(h => now >= new Date(start.getTime() - h * 3600000));
    const notified = cleanRemind(r[EV.NOTIFIED]);
    if (!due.some(h => !notified.includes(h) && !ev.reminded.includes(h))) continue;
    sh.getRange(i + 1, EV.NOTIFIED + 1).setValue(cleanRemind(notified.concat(due)).join(','));
    const to = getNotifyEmail({ userId: creatorOf(r) });
    if (!to) continue;
    const when = untilText(start);
    const link = 'https://liff.line.me/' + CONFIG.LIFF_ID + '?remind=' + encodeURIComponent(ev.id);
    const count = logOf(ev.id).length;
    try {
      MailApp.sendEmail({
        to, name: 'LINE簡易活動登記',
        subject: '⏰ 該發提醒了：' + ev.title + '（' + when + '）',
        htmlBody: '<p>「<b>' + ev.title + '</b>」' + when + '。</p>'
          + '<p>時間：' + (ev.date || '') + (ev.place ? '<br>地點：' + ev.place : '') + '<br>目前登記：' + count + ' 人' + (ev.capacity ? '／名額 ' + ev.capacity : '') + '</p>'
          + '<p><a href="' + link + '" style="display:inline-block;background:#3A6848;color:#fff;padding:12px 20px;border-radius:999px;text-decoration:none;font-weight:bold">打開平台，發提醒到群組</a></p>'
          + '<p style="color:#888;font-size:12px">在手機上點按鈕會開啟 LINE，選群組就能送出提醒卡。不想再收到這類信，可在平台「新活動 → 提醒」清空 Email。</p>',
      });
    } catch (e) { console.warn('寄送提醒失敗：' + e.message); }
  }
}

function ensureReminderTrigger() {
  if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'remindTick')) return;
  ScriptApp.newTrigger('remindTick').timeBased().everyMinutes(15).create();
}
