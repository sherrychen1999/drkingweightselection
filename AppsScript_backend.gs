/**
 * 減重特定申報資料庫 - 後端 Apps Script
 *
 * 部署步驟：
 * 1. 建立新的 Google Sheet
 * 2. 在 Sheet 內：擴充功能 → Apps Script
 * 3. 把這整份程式碼貼進去取代原本內容
 * 4. 儲存後執行一次 setupSheets() 建立工作表
 * 5. 部署 → 新增部署作業 → 類型：網頁應用程式
 *    執行身分：我，存取權：「知道連結的任何人」
 * 6. 複製產生的 Web App URL，貼到減重資料篩選器設定裡
 */

const SHEET_DETAIL = '明細';
const SHEET_LOG = '上傳紀錄';
const SHEET_SUMMARY = '區間摘要';

// === HTTP 入口 ===
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheets_(ss);

    const ts = new Date();
    const clinic   = String(payload.clinic || '').trim();
    const range    = String(payload.dateRange || '').trim();
    const filename = String(payload.filename || '').trim();
    const total    = Number(payload.total || 0);
    const kept     = Number(payload.kept || 0);
    const removed  = Number(payload.removed || 0);
    const removedDetail = String(payload.removedDetail || '');
    const rows = Array.isArray(payload.rows) ? payload.rows : [];

    if (!clinic || !range) {
      return jsonResponse_({ status: 'error', message: '院區或日期區間缺失（請檢查檔名格式）' });
    }

    // 防重複：同檔名 + 同院區 + 同區間 視為重複
    const logSheet = ss.getSheetByName(SHEET_LOG);
    const lastLogRow = logSheet.getLastRow();
    let duplicate = false;
    if (lastLogRow > 1) {
      const logData = logSheet.getRange(2, 1, lastLogRow - 1, 4).getValues();
      for (const r of logData) {
        if (String(r[1]).trim() === clinic && String(r[2]).trim() === range && String(r[3]).trim() === filename) {
          duplicate = true; break;
        }
      }
    }
    if (duplicate) {
      return jsonResponse_({ status: 'duplicate', message: '此檔案已上傳過：' + filename });
    }

    // 寫入明細
    const detailSheet = ss.getSheetByName(SHEET_DETAIL);
    if (rows.length > 0) {
      const detailRows = rows.map(r => [
        ts, clinic, range,
        r.診別 || '', r.醫師 || '', r.kcstmr || '', r.date || '', r.labeno || ''
      ]);
      detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, 8).setValues(detailRows);
    }

    // 寫入上傳紀錄
    logSheet.appendRow([ts, clinic, range, filename, total, kept, removed, removedDetail]);

    // 同步刷新區間摘要
    refreshSummary_(ss);

    return jsonResponse_({
      status: 'ok',
      message: '上傳成功',
      clinic, dateRange: range, kept, removed
    });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: '伺服器錯誤：' + (err && err.message ? err.message : err) });
  }
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.action === 'aggregated') {
    return jsonResponse_(getAggregated_());
  }
  return jsonResponse_({ status: 'alive', message: '減重資料庫運作中', time: new Date() });
}

// === 抽取週數：擷取 labeno 中第一個 (數字)+W/w ===
function extractWeeks_(s) {
  if (!s) return 0;
  const m = String(s).match(/(\d{1,2})[Ww]/);
  return m ? parseInt(m[1], 10) : 0;
}

// === 取出儀表板用的聚合資料（三層獨立去重）===
function getAggregated_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailSheet = ss.getSheetByName(SHEET_DETAIL);
  if (!detailSheet || detailSheet.getLastRow() < 2) {
    return { status: 'ok', ranges: [], byRangeClinic: {}, byRangeClinicDoc: {}, byRangeDoctor: {} };
  }
  const data = detailSheet.getRange(2, 1, detailSheet.getLastRow() - 1, 8).getValues();
  // 欄位順序：上傳時間, 院區, 日期區間, 診別, 醫師, kcstmr, date, labeno

  const byRC  = {};  // byRC[range][clinic]         = {count, weeks, persons:{kcstmr:true...}}
  const byRCD = {};  // byRCD[range][clinic][doctor]= {count, weeks, persons}
  const byRD  = {};  // byRD[range][doctor]         = {count, weeks, persons}

  function bucket(map, keys) {
    let m = map;
    for (let i = 0; i < keys.length - 1; i++) {
      if (!m[keys[i]]) m[keys[i]] = {};
      m = m[keys[i]];
    }
    const last = keys[keys.length - 1];
    if (!m[last]) m[last] = { count: 0, weeks: 0, persons: {} };
    return m[last];
  }

  // 「拿藥」判斷：labeno 含「拿」或其容錯字 → 歸入「減重拿藥」虛擬醫師
  const reNa = /[拿那哪納拎]/;

  for (const r of data) {
    const range   = String(r[2]).trim();
    const clinic  = String(r[1]).trim();
    const docRaw  = String(r[4]).trim();
    const kcstmr  = String(r[5]).trim();
    const labeno  = String(r[7]).trim();
    if (!range || !clinic || !docRaw) continue;
    const w = extractWeeks_(labeno);
    // 含「拿」→ 不掛醫師頭上，改掛「減重拿藥」
    const doctor = reNa.test(labeno) ? '減重拿藥' : docRaw;

    const a = bucket(byRC,  [range, clinic]);
    const b = bucket(byRCD, [range, clinic, doctor]);
    const c = bucket(byRD,  [range, doctor]);
    [a, b, c].forEach(x => {
      x.count += 1;
      x.weeks += w;
      if (kcstmr) x.persons[kcstmr] = true;
    });
  }

  // 把 persons 物件轉成數量（人頭數）
  function finalize(map, depth) {
    if (depth <= 0) {
      map.persons = Object.keys(map.persons || {}).length;
      return;
    }
    for (const k in map) finalize(map[k], depth - 1);
  }
  finalize(byRC,  2);  // range → clinic → data
  finalize(byRCD, 3);  // range → clinic → doctor → data
  finalize(byRD,  2);  // range → doctor → data

  // 上傳紀錄：完整列表 + 每區間院區數
  const logSheet = ss.getSheetByName(SHEET_LOG);
  const uploaded = {};
  const uploadList = [];  // 完整紀錄 [{ts, clinic, range, filename, total, kept, removed}]
  if (logSheet && logSheet.getLastRow() >= 2) {
    const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 8).getValues();
    for (const r of logData) {
      const ts = r[0];
      const clinic = String(r[1]).trim();
      const range  = String(r[2]).trim();
      const filename = String(r[3]).trim();
      if (!range || !clinic) continue;
      if (!uploaded[range]) uploaded[range] = {};
      uploaded[range][clinic] = true;
      uploadList.push({
        ts: ts instanceof Date ? ts.toISOString() : String(ts),
        clinic, range, filename,
        total: Number(r[4]) || 0,
        kept: Number(r[5]) || 0,
        removed: Number(r[6]) || 0
      });
    }
  }
  const uploadedCount = {};
  for (const r in uploaded) uploadedCount[r] = Object.keys(uploaded[r]).length;

  return {
    status: 'ok',
    ranges: Object.keys(byRCD).sort(),
    byRangeClinic: byRC,
    byRangeClinicDoc: byRCD,
    byRangeDoctor: byRD,
    uploadedClinics: uploadedCount,
    uploadList: uploadList,
    generatedAt: new Date().toISOString()
  };
}

// === 工具：統一 JSON 回應（含 CORS 標頭）===
function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureSheets_(ss) {
  if (!ss.getSheetByName(SHEET_DETAIL)) {
    const s = ss.insertSheet(SHEET_DETAIL);
    s.getRange(1, 1, 1, 8).setValues([['上傳時間','院區','日期區間','診別','醫師','kcstmr','date','labeno']]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, 8).setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
    s.setColumnWidth(1, 150); s.setColumnWidth(2, 80); s.setColumnWidth(3, 100);
  }
  if (!ss.getSheetByName(SHEET_LOG)) {
    const s = ss.insertSheet(SHEET_LOG);
    s.getRange(1, 1, 1, 8).setValues([['上傳時間','院區','日期區間','檔名','原始筆數','保留筆數','刪除筆數','刪除明細']]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, 8).setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
    s.setColumnWidth(1, 150); s.setColumnWidth(4, 220); s.setColumnWidth(8, 320);
  }
  if (!ss.getSheetByName(SHEET_SUMMARY)) {
    const s = ss.insertSheet(SHEET_SUMMARY);
    s.getRange(1, 1, 1, 5).setValues([['日期區間','院區','已上傳檔案數','保留筆數','刪除筆數']]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, 5).setBackground('#1e293b').setFontColor('#fff').setFontWeight('bold');
  }
}

function refreshSummary_(ss) {
  const logSheet = ss.getSheetByName(SHEET_LOG);
  const summarySheet = ss.getSheetByName(SHEET_SUMMARY);
  if (logSheet.getLastRow() < 2) return;
  const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 8).getValues();
  const map = {}; // key = range|clinic
  for (const r of logData) {
    const range = String(r[2]).trim();
    const clinic = String(r[1]).trim();
    if (!range || !clinic) continue;
    const key = range + '|' + clinic;
    if (!map[key]) map[key] = { range, clinic, files: 0, kept: 0, removed: 0 };
    map[key].files += 1;
    map[key].kept += Number(r[5]) || 0;
    map[key].removed += Number(r[6]) || 0;
  }
  const rows = Object.values(map).sort((a, b) => {
    if (a.range !== b.range) return a.range < b.range ? -1 : 1;
    return a.clinic < b.clinic ? -1 : 1;
  });
  summarySheet.getRange(2, 1, summarySheet.getMaxRows() - 1, 5).clearContent();
  if (rows.length) {
    summarySheet.getRange(2, 1, rows.length, 5).setValues(
      rows.map(x => [x.range, x.clinic, x.files, x.kept, x.removed])
    );
  }
}

// === 手動初始化（部署前在編輯器執行一次）===
function setupSheets() {
  ensureSheets_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getActive().toast('工作表已建立完成');
}

// === 手動重整區間摘要 ===
function refreshSummary() {
  refreshSummary_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getActive().toast('區間摘要已更新');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('減重資料庫')
    .addItem('初始化工作表', 'setupSheets')
    .addItem('重整區間摘要', 'refreshSummary')
    .addToUi();
}
