/**
 * 打卡小幫手：跨裝置同步（Google Apps Script）
 *
 * 用法：在你自己的 Google 試算表 → 擴充功能 → Apps Script，貼上整份程式，
 * 把下面的 TOKEN 換成你自己的同步密碼，再部署成「網頁應用程式」。
 * 執行身分：我；誰可以存取：所有人。部署後的網址（結尾 /exec）貼到打卡網頁的設定頁。
 *
 * 資料只存在這張試算表：
 *   「打卡紀錄」分頁：每筆打卡一列（最後一欄是完整資料，請勿手動修改）
 *   「設定」分頁：A1 更新時間、A2 設定內容（含表單網址與 Email，不要分享這張試算表）
 */
const TOKEN = 'CHANGE_ME';

const LOG_SHEET = '打卡紀錄';
const CFG_SHEET = '設定';
const LOG_HEAD = ['id', '日期', '時間', '類型', '地點', '狀態', '裝置', '已刪除', '更新時間', 'json'];
const TYPE_LABEL = { in: '上班', out: '下班', trip: '上班時間公出' };

function doGet() {
  return out_({ ok: true, msg: 'daka sync is running' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad json' }); }
  if (TOKEN === 'CHANGE_ME' || !req || req.token !== TOKEN) return out_({ ok: false, error: 'unauthorized' });
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logs = mergeLog_(sheet_(ss, LOG_SHEET, LOG_HEAD), Array.isArray(req.upsert) ? req.upsert : []);
    const config = mergeConfig_(sheet_(ss, CFG_SHEET, null), req.config);
    return out_({ ok: true, log: logs, config: config });
  } finally {
    lock.releaseLock();
  }
}

function mergeLog_(sh, incoming) {
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, LOG_HEAD.length).getValues() : [];
  const index = {};
  const all = [];
  rows.forEach(function (row, i) {
    let entry;
    try { entry = JSON.parse(row[LOG_HEAD.length - 1]); } catch (err) { return; }
    if (!entry || !entry.id) return;
    index[entry.id] = { rowNo: i + 2, entry: entry };
    all.push(entry);
  });
  incoming.forEach(function (entry) {
    if (!entry || typeof entry !== 'object' || !entry.id) return;
    const cur = index[entry.id];
    if (cur && (cur.entry.updatedAt || 0) >= (entry.updatedAt || 0)) return;
    const values = [toRow_(entry)];
    if (cur) {
      sh.getRange(cur.rowNo, 1, 1, LOG_HEAD.length).setValues(values);
      all[all.indexOf(cur.entry)] = entry;
      cur.entry = entry;
    } else {
      sh.appendRow(values[0]);
      index[entry.id] = { rowNo: sh.getLastRow(), entry: entry };
      all.push(entry);
    }
  });
  return all;
}

function toRow_(r) {
  return [
    String(r.id), String(r.date || ''), String(r.time || ''), TYPE_LABEL[r.type] || String(r.type || ''),
    String(r.place || ''), String(r.status || ''), String(r.device || ''), r.deleted ? '是' : '',
    r.updatedAt ? new Date(r.updatedAt) : '', JSON.stringify(r)
  ];
}

function mergeConfig_(sh, incoming) {
  let stored = null;
  try { stored = JSON.parse(sh.getRange('A2').getValue() || 'null'); } catch (err) { stored = null; }
  if (incoming && typeof incoming === 'object' && (incoming.updatedAt || 0) > ((stored && stored.updatedAt) || 0)) {
    sh.getRange('A1').setValue(new Date(incoming.updatedAt));
    sh.getRange('A2').setValue(JSON.stringify(incoming));
    stored = incoming;
  }
  return stored;
}

function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (head) { sh.appendRow(head); sh.setFrozenRows(1); }
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
