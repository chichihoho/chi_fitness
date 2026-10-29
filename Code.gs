// 身體數據與飲食日記：Google 試算表同步程式（v2：合併寫入＋每日自動備份）
// 使用前，請把下面這行的密碼改成你自己的（日記頁面要填同一組）
const PASSWORD = '請改成你自己的密碼';

const SHEET = '日記';
const CFG = '設定';
const HEAD = ['日期','體重(kg)','體脂率(%)','骨骼肌率(%)','BMI','基礎代謝率(kcal)','身體年齡','內臟脂肪','有氧(分)','重訓(分)','熱量(kcal)','蛋白質(g)','重訓部位','餐點明細','有氧項目'];
const KEYS = ['weight','fat','muscle','bmi','bmr','bage','visc','cardio','strength','kcal','protein'];

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// 用瀏覽器直接開部署網址時，會看到這行，代表部署成功
function doGet() {
  return json_({ ok: true, msg: '服務運作中' });
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  if (p.token !== PASSWORD) return json_({ ok: false, error: '密碼錯誤' });
  if (p.action === 'pull') return json_({ ok: true, data: read_() });
  if (p.action === 'push') {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try { write_(p.data); } finally { lock.releaseLock(); }
    return json_({ ok: true });
  }
  return json_({ ok: false, error: '未知動作' });
}

function read_() {
  const sh = sheet_(SHEET), cfg = sheet_(CFG);
  let profile = null;
  try { profile = JSON.parse(cfg.getRange('A1').getValue()); } catch (e) {}
  const entries = {};
  const last = sh.getLastRow();
  if (last > 1) {
    const rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
    rows.forEach(function (r) {
      let d = r[0];
      if (d instanceof Date) d = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      d = String(d);
      if (!d) return;
      const e = {};
      KEYS.forEach(function (k, i) { if (r[i + 1] !== '') e[k] = Number(r[i + 1]); });
      if (r[12]) e.groups = String(r[12]).split('、');
      if (r[13]) { try { e.meals = JSON.parse(r[13]); } catch (x) {} }
      if (r[14]) e.ctypes = String(r[14]).split('、');
      entries[d] = e;
    });
  }
  return { profile: profile, entries: entries };
}

function backup_() {
  // 每天第一次寫入前，先把「日記」分頁複製成「備份_日期」，保留最近 7 天
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  const name = '備份_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  if (ss.getSheetByName(name)) return;
  sh.copyTo(ss).setName(name);
  const olds = ss.getSheets().filter(function (s) { return /^備份_\d{8}$/.test(s.getName()); })
    .sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });
  while (olds.length > 7) ss.deleteSheet(olds.shift());
}

// 合併寫入：傳來的日期會新增或更新，沒傳來的日期一律保留；
// 只有 data.del 清單裡明確指定的日期才會被刪除。
function write_(data) {
  data = data || {};
  backup_();
  const cur = read_();
  const entries = cur.entries;
  (data.del || []).forEach(function (d) { delete entries[d]; });
  const inc = data.entries || {};
  Object.keys(inc).forEach(function (d) { entries[d] = inc[d]; });
  const profile = (data.profile && data.profile.sex) ? data.profile : (cur.profile || {});
  const sh = sheet_(SHEET), cfg = sheet_(CFG);
  cfg.getRange('A1').setValue(JSON.stringify(profile));
  sh.clear();
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  const ds = Object.keys(entries).sort();
  if (!ds.length) return;
  const rows = ds.map(function (d) {
    const e = entries[d];
    return [d].concat(
      KEYS.map(function (k) { return e[k] == null ? '' : e[k]; }),
      [(e.groups || []).join('、'), e.meals ? JSON.stringify(e.meals) : '', (e.ctypes || []).join('、')]
    );
  });
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, HEAD.length).setValues(rows);
}
