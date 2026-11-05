/**
 * LINE 群組一鍵報名系統 — Google Apps Script 後端
 * 貼到 Google 試算表的「擴充功能 → Apps Script」，先執行一次 setup()。
 */

const CONFIG = {
  CHANNEL_ID: '請填入 LINE Login Channel ID', // LINE Developers → 你的 LINE Login channel → Basic settings
};

const SH = { ROSTER: '名單', EVENTS: '活動', LOG: '報名紀錄', PUBLISHERS: '發布者' };

const HEADERS = {
  '名單': ['姓名', '手機', 'LINE userId', 'LINE名稱'],
  '活動': ['活動代碼', '活動名稱', '日期時間', '地點', '說明', '圖片網址', '按鈕文字', '開放報名', '建立者', '建立時間'],
  '報名紀錄': ['時間', '活動代碼', '活動名稱', '姓名', 'LINE名稱', 'LINE userId'],
  '發布者': ['LINE userId', '名稱', '備註'],
};

/** 第一次使用時手動執行：建立四個工作表與標題列 */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold');
    }
  });
  ss.getSheetByName(SH.ROSTER).getRange('B:B').setNumberFormat('@'); // 手機保留開頭的 0
  ss.getSheetByName(SH.EVENTS).getRange('C:C').setNumberFormat('@'); // 日期時間存成文字
}

function doGet() {
  return json({ ok: true, msg: '報名系統運作中' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'event') {
      out = { ok: true, event: getEvent(req.eventId) };
    } else {
      const user = verifyIdToken(req.idToken);
      switch (req.action) {
        case 'register':    out = register(user, req.eventId); break;
        case 'bind':        out = bind(user, req.eventId, req.name, req.phone); break;
        case 'me':          out = { ok: true, userId: user.userId, displayName: user.displayName, isPublisher: isPublisher(user.userId) }; break;
        case 'createEvent': requirePublisher(user); out = { ok: true, event: createEvent(user, req.event || {}) }; break;
        case 'listEvents':  requirePublisher(user); out = { ok: true, events: listEvents() }; break;
        case 'setOpen':     requirePublisher(user); out = { ok: true, event: setOpen(req.eventId, !!req.open) }; break;
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
function normPhone(v) { return String(v || '').replace(/\D/g, '').slice(-9); }
function isTrue(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function fmtDate(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') : String(v || '');
}
function rowToEvent(r) {
  return {
    id: String(r[0]), title: String(r[1]), date: fmtDate(r[2]), place: String(r[3] || ''),
    desc: String(r[4] || ''), image: String(r[5] || ''), button: String(r[6] || '我要報名'), open: isTrue(r[7]),
  };
}

/* ---------- 發布者 ---------- */
function isPublisher(uid) { return rows(SH.PUBLISHERS).some(r => String(r[0]).trim() === uid); }
function requirePublisher(user) { if (!isPublisher(user.userId)) throw new Error('你沒有發布權限'); }

/* ---------- 活動 ---------- */
function getEvent(id) {
  const r = rows(SH.EVENTS).find(r => String(r[0]) === String(id));
  return r ? rowToEvent(r) : null;
}

function createEvent(user, ev) {
  const title = String(ev.title || '').trim();
  if (!title) throw new Error('請填寫活動名稱');
  const image = /^https:\/\//.test(ev.image || '') ? ev.image.trim() : '';
  const id = 'E' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  const row = [id, title, String(ev.date || '').trim(), String(ev.place || '').trim(), String(ev.desc || '').trim(),
    image, String(ev.button || '我要報名').trim().slice(0, 20), true, user.displayName + ' (' + user.userId + ')', new Date()];
  sheet(SH.EVENTS).appendRow(row);
  return rowToEvent(row);
}

function listEvents() {
  const counts = {};
  rows(SH.LOG).forEach(r => { counts[r[1]] = (counts[r[1]] || 0) + 1; });
  return rows(SH.EVENTS).reverse().slice(0, 30).map(r => {
    const ev = rowToEvent(r);
    ev.count = counts[ev.id] || 0;
    return ev;
  });
}

function setOpen(id, open) {
  const data = sheet(SH.EVENTS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet(SH.EVENTS).getRange(i + 1, 8).setValue(open);
      data[i][7] = open;
      return rowToEvent(data[i]);
    }
  }
  throw new Error('找不到這個活動');
}

/* ---------- 報名 ---------- */
function register(user, eventId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ev = getEvent(eventId);
    if (!ev) throw new Error('找不到這個活動');
    if (!ev.open) return { ok: true, status: 'closed', event: ev };

    const rs = sheet(SH.ROSTER);
    const data = rs.getDataRange().getValues();
    let idx = -1;
    for (let i = 1; i < data.length; i++) if (String(data[i][2]).trim() === user.userId) { idx = i; break; }
    if (idx < 0) return { ok: true, status: 'need_bind', event: ev };

    const name = String(data[idx][0]);
    if (user.displayName && data[idx][3] !== user.displayName) rs.getRange(idx + 1, 4).setValue(user.displayName);

    const dup = rows(SH.LOG).some(r => String(r[1]) === ev.id && String(r[5]) === user.userId);
    if (dup) return { ok: true, status: 'already', name, event: ev };

    sheet(SH.LOG).appendRow([new Date(), ev.id, ev.title, name, user.displayName, user.userId]);
    return { ok: true, status: 'registered', name, event: ev };
  } finally {
    lock.releaseLock();
  }
}

/** 第一次報名：用姓名＋手機比對名單，把 LINE userId 綁上去 */
function bind(user, eventId, name, phone) {
  name = String(name || '').trim();
  const p = normPhone(phone);
  if (!name || p.length < 9) throw new Error('請填寫姓名和完整手機號碼');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rs = sheet(SH.ROSTER);
    const data = rs.getDataRange().getValues();
    const already = data.some((r, i) => i > 0 && String(r[2]).trim() === user.userId);
    if (!already) {
      let idx = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === name && normPhone(data[i][1]) === p) { idx = i; break; }
      }
      if (idx < 0) return { ok: true, status: 'not_found', event: getEvent(eventId) };
      const bound = String(data[idx][2]).trim();
      if (bound && bound !== user.userId) return { ok: true, status: 'bound_other', event: getEvent(eventId) };
      rs.getRange(idx + 1, 3, 1, 2).setValues([[user.userId, user.displayName]]);
    }
  } finally {
    lock.releaseLock();
  }
  return register(user, eventId);
}
