/**
 * «Смена в отеле Сауыт»: серверная часть на Google Apps Script.
 * Хранит результаты студентов в листе «students» и показывает, кто сейчас в сети.
 * Пароль преподавателя меняется в строке ниже.
 */
const TEACHER_PW = '2105';
const HEAD = ['id', 'name', 'group', 'lang', 'total', 'done', 'door', 'lastAt', 'json'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Смена в отеле «Сауыт»')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('students');
  if (!sh) { sh = ss.insertSheet('students'); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}

function checkPw_(pw) { if (String(pw) !== TEACHER_PW) throw new Error('forbidden'); }

function findRow_(sh, id) {
  const cache = CacheService.getScriptCache();
  const c = cache.get('row:' + id);
  if (c) { const r = Number(c); if (r > 1 && sh.getRange(r, 1).getValue() === id) return r; }
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0] === id) { cache.put('row:' + id, String(i + 2), 21600); return i + 2; }
  }
  return 0;
}

/** Студент сохраняет свой прогресс. */
function api_save(id, doc) {
  id = String(id || '').replace(/[^A-Za-z0-9_:\-]/g, '').slice(0, 80);
  if (!id || !doc) return false;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const json = JSON.stringify(doc);
    const row = [id, String(doc.name || '').slice(0, 60), String(doc.group || '').slice(0, 30), doc.lang || '', Number(doc.tot) || 0, Number(doc.k) || 0, doc.door || '', Number(doc.lastAt) || Date.now(), json.length > 45000 ? json.slice(0, 45000) : json];
    const r = findRow_(sh, id);
    if (r) sh.getRange(r, 1, 1, HEAD.length).setValues([row]);
    else { sh.appendRow(row); CacheService.getScriptCache().put('row:' + id, String(sh.getLastRow()), 21600); }
  } finally { lock.releaseLock(); }
  return true;
}

/** Студент сообщает, что он в сети и в каком окне (хранится в кэше, быстро). */
function api_ping(id, presence) {
  id = String(id || '').replace(/[^A-Za-z0-9_:\-]/g, '').slice(0, 80);
  if (!id) return false;
  const cache = CacheService.getScriptCache();
  cache.put('p:' + id, JSON.stringify({ presence: presence || {}, t: Date.now() }), 21600);
  let ids = JSON.parse(cache.get('ids') || '[]');
  if (ids.indexOf(id) < 0) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      ids = JSON.parse(cache.get('ids') || '[]');
      if (ids.indexOf(id) < 0) { ids.push(id); cache.put('ids', JSON.stringify(ids), 21600); }
    } finally { lock.releaseLock(); }
  }
  return true;
}

/** Преподаватель получает всех студентов и тех, кто в сети. */
function api_list(pw) {
  checkPw_(pw);
  const sh = sheet_();
  const last = sh.getLastRow();
  const docs = {};
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
    vals.forEach(function (v) {
      try { docs[v[0]] = JSON.parse(v[8]); } catch (e) { docs[v[0]] = { name: v[1], group: v[2], lang: v[3], tot: v[4], k: v[5], door: v[6], lastAt: v[7], res: {} }; }
    });
  }
  const cache = CacheService.getScriptCache();
  const ids = JSON.parse(cache.get('ids') || '[]');
  const peers = [];
  if (ids.length) {
    const all = cache.getAll(ids.map(function (i) { return 'p:' + i; }));
    ids.forEach(function (i) {
      const raw = all['p:' + i];
      if (!raw) return;
      const p = JSON.parse(raw);
      if (Date.now() - p.t < 90000) peers.push({ id: i, presence: p.presence, t: p.t });
    });
  }
  return { docs: docs, peers: peers, now: Date.now() };
}

/** Состав урока: читают все, меняет преподаватель. */
function api_cfg_get() {
  const v = PropertiesService.getScriptProperties().getProperty('cfg');
  return v ? JSON.parse(v) : null;
}
function api_cfg_set(pw, on) {
  checkPw_(pw);
  PropertiesService.getScriptProperties().setProperty('cfg', JSON.stringify(on || []));
  return true;
}

/** Удаление студентов из списка. */
function api_del(pw, ids) {
  checkPw_(pw);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const cache = CacheService.getScriptCache();
    (ids || []).forEach(function (id) {
      const r = findRow_(sh, id);
      if (r) sh.deleteRow(r);
      cache.remove('p:' + id);
    });
    cache.removeAll((ids || []).map(function (i) { return 'row:' + i; }));
    // номера строк сдвинулись: сбросить весь кэш строк
    const all = JSON.parse(cache.get('ids') || '[]');
    cache.removeAll(all.map(function (i) { return 'row:' + i; }));
  } finally { lock.releaseLock(); }
  return true;
}
