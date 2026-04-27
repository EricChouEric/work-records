// Script Properties:
//   MASTER_SHEET_ID     - Master Google Sheet ID
//   MANAGER_INVITE_CODE - Secret code for manager registration
//
// Master Sheet tabs required:
//   Users      : [工號, 姓名, 密碼, 角色, 組別]
//   Groups     : [組別, 建立時間]
//   WorkOrders : [工單號碼, 船號, 工單內容, 預估工時, 建立時間, 備註]
//   Reports    : [提交時間, 施工日期, 工號, 員工姓名, 組別, 工單號碼, 船號, 實際工時, 類別]
//   Sessions   : [token, userId, expiry]
//
// Run setupSheets() once from the GAS editor to create missing tabs.

function doGet(e) {
  if (!e || !e.parameter) return jsonResponse({ status: 'error', message: '請透過 Web App URL 呼叫此 API' });
  const action = e.parameter.action;
  try {
    switch (action) {
      case 'register':         return register(e.parameter);
      case 'registerManager':  return registerManager(e.parameter);
      case 'login':            return login(e.parameter);
      case 'logout':           return logout(e.parameter);
      case 'getGroups':        return getGroups(e.parameter);
      case 'submitReport':     return submitReport(e.parameter);
      case 'getMyReports':     return getMyReports(e.parameter);
      case 'getWorkOrders':    return getWorkOrders(e.parameter);
      case 'addWorkOrder':     return addWorkOrder(e.parameter);
      case 'deleteWorkOrder':  return deleteWorkOrder(e.parameter);
      case 'importWorkOrders': return importWorkOrders(e.parameter);
      case 'importGroups':     return importGroups(e.parameter);
      case 'testImportSheet':  return testImportSheet(e.parameter);
      case 'importFromSheet':  return importFromSheet(e.parameter);
      case 'getAllReports':    return getAllReports(e.parameter);
      case 'getAllEmployees':  return getAllEmployees(e.parameter);
      default:                 return jsonResponse({ status: 'error', message: '未知操作' });
    }
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Sheet helpers ─────────────────────────────────────────────────────────────

function getMaster() {
  const id = PropertiesService.getScriptProperties().getProperty('MASTER_SHEET_ID');
  if (!id) throw new Error('MASTER_SHEET_ID 未設定，請至 Script Properties 設定');
  return SpreadsheetApp.openById(id);
}

function usersSheet()      { return getMaster().getSheetByName('Users'); }
function groupsSheet()     { return getMaster().getSheetByName('Groups'); }
function sessionsSheet()   { return getMaster().getSheetByName('Sessions'); }
function workOrdersSheet() { return getMaster().getSheetByName('WorkOrders'); }
function reportsSheet()    { return getMaster().getSheetByName('Reports'); }

function formatTextColumns(ws, cols) {
  if (!ws || !cols.length) return;
  const maxRows = Math.max(ws.getMaxRows(), 1);
  cols.forEach(col => ws.getRange(1, col, maxRows, 1).setNumberFormat('@'));
}

function formatIdentifierColumns() {
  formatTextColumns(usersSheet(),      [1, 2, 3, 4, 5]);
  formatTextColumns(groupsSheet(),     [1, 2]);
  formatTextColumns(sessionsSheet(),   [1, 2]);
  formatTextColumns(workOrdersSheet(), [1, 2, 3, 4, 5, 6]);
  formatTextColumns(reportsSheet(),    [1, 2, 3, 4, 5, 6, 7, 8, 9]);
}

function appendTextRow(ws, values, textCols) {
  const row = ws.getLastRow() + 1;
  textCols.forEach(col => ws.getRange(row, col).setNumberFormat('@'));
  ws.getRange(row, 1, 1, values.length).setValues([values]);
}

function appendTextRows(ws, rows, textCols) {
  if (!rows.length) return;
  const startRow = ws.getLastRow() + 1;
  textCols.forEach(col => ws.getRange(startRow, col, rows.length, 1).setNumberFormat('@'));
  ws.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function setPlainTextValues(ws, row, col, values) {
  if (!values.length || !values[0].length) return;
  const range = ws.getRange(row, col, values.length, values[0].length);
  range.setNumberFormat('@');
  range.setValues(values);
}

function normalizeWorkOrdersSheetCodes_() {
  const ws = workOrdersSheet();
  if (!ws) throw new Error('找不到 WorkOrders 分頁');

  const lastRow = ws.getLastRow();
  if (lastRow < 2) {
    return { count: 0, preview: '', formats: '' };
  }

  const codeRange = ws.getRange(2, 1, lastRow - 1, 2);
  const codeRows = codeRange.getDisplayValues().map(row => [
    sheetCodeValue(normalizeWorkOrderId(row[0])),
    sheetCodeValue(normalizeShipNo(row[1]))
  ]);

  codeRange.clearFormat();
  setPlainTextValues(ws, 2, 1, codeRows);
  SpreadsheetApp.flush();

  return {
    count: codeRows.length,
    preview: codeRange.getDisplayValues().slice(0, 5).map(r => r.join('/')).join(', '),
    formats: codeRange.getNumberFormats().slice(0, 2).map(r => r.join('/')).join(', ')
  };
}

// Run once from GAS editor to initialise missing sheets
function setupSheets() {
  const ss = getMaster();
  console.log('操作的試算表：' + ss.getUrl());
  function ensure(name, headers) {
    let ws = ss.getSheetByName(name);
    if (!ws) {
      ws = ss.insertSheet(name);
      ws.appendRow(headers);
      ws.setFrozenRows(1);
      console.log('已建立分頁：' + name);
    } else {
      console.log('分頁已存在：' + name);
      const oldHeaders = ws.getRange(1, 1, 1, Math.max(headers.length, ws.getLastColumn())).getDisplayValues()[0];
      if (name === 'WorkOrders' && oldHeaders[2] === '建立時間') {
        const lastRow = ws.getLastRow();
        if (lastRow > 1) {
          const oldRows = ws.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
          const migrated = oldRows.map(row => [row[0], row[1], '', '', row[2], row[3]]);
          setPlainTextValues(ws, 2, 1, migrated);
          console.log('已升級 WorkOrders 欄位：新增工單內容與預估工時');
        }
      }
    }
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    return ws;
  }
  ensure('Users',      ['工號', '姓名', '密碼', '角色', '組別']);
  ensure('Groups',     ['組別', '建立時間']);
  ensure('Sessions',   ['token', 'userId', 'expiry']);
  ensure('WorkOrders', ['工單號碼', '船號', '工單內容', '預估工時', '建立時間', '備註']);
  ensure('Reports',    ['提交時間', '施工日期', '工號', '員工姓名', '組別', '工單號碼', '船號', '實際工時', '類別']);
  formatIdentifierColumns();
  console.log('setupSheets 完成');
}

// 用來確認 MASTER_SHEET_ID 是否設定正確
function debugConfig() {
  const id = PropertiesService.getScriptProperties().getProperty('MASTER_SHEET_ID');
  console.log('MASTER_SHEET_ID = ' + id);
  if (!id) {
    console.log('❌ 未設定！請到指令碼屬性新增 MASTER_SHEET_ID');
    return;
  }
  try {
    const ss = SpreadsheetApp.openById(id);
    console.log('✅ 成功開啟試算表：' + ss.getName());
    console.log('網址：' + ss.getUrl());
    console.log('現有分頁：' + ss.getSheets().map(s => s.getName()).join(', '));
  } catch (e) {
    console.log('❌ 無法開啟試算表：' + e.message);
  }
}

// Run once if old imports converted codes like 001/012 into 1/12.
function repairWorkOrderCodes() {
  const ws = workOrdersSheet();
  if (!ws) throw new Error('找不到 WorkOrders 分頁');

  const lastRow = ws.getLastRow();
  if (lastRow < 2) {
    console.log('WorkOrders 沒有資料需要修復');
    return;
  }

  const range = ws.getRange(2, 1, lastRow - 1, 4);
  const rows = range.getDisplayValues().map(row => [
    String(row[0] || ''),
    String(row[1] || ''),
    String(row[2] || ''),
    String(row[3] || '')
  ]);
  const result = normalizeWorkOrdersSheetCodes_();
  ws.getRange(2, 1, lastRow - 1, 2).clearFormat();
  ws.getRange(2, 3, lastRow - 1, 2).clearFormat();
  rows.forEach((row, index) => {
    const fixed = result.count ? ws.getRange(index + 2, 1, 1, 2).getDisplayValues()[0] : row;
    row[0] = fixed[0];
    row[1] = fixed[1];
  });
  setPlainTextValues(ws, 2, 1, rows);
  console.log(`repairWorkOrderCodes 完成，共修復 ${rows.length} 筆`);
}

function forceRepairWorkOrderCodes() {
  const ss = getMaster();
  const ws = ss.getSheetByName('WorkOrders');
  if (!ws) throw new Error('找不到 WorkOrders 分頁');
  const lastRow = ws.getLastRow();
  console.log('MASTER_SHEET_ID = ' + PropertiesService.getScriptProperties().getProperty('MASTER_SHEET_ID'));
  console.log('正在修復的試算表 = ' + ss.getUrl());
  if (lastRow < 2) {
    console.log('WorkOrders 沒有資料需要修復');
    return;
  }

  const otherRange = ws.getRange(2, 3, lastRow - 1, 2);
  otherRange.clearFormat();
  otherRange.setNumberFormat('@');
  const result = normalizeWorkOrdersSheetCodes_();

  console.log(`forceRepairWorkOrderCodes 完成，共修復 ${result.count} 筆。前 5 筆顯示值：${result.preview}`);
  console.log(`A/B 欄格式預覽：${result.formats}`);
}

// ── Utilities ─────────────────────────────────────────────────────────────────

function toDateStr(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

function normalizeShipNo(val) {
  const text = String(val || '').trim();
  return normalizeThreeDigitCode(text);
}

function normalizeWorkOrderId(val) {
  const text = String(val || '').trim();
  return normalizeThreeDigitCode(text);
}

function normalizeEmployeeId(val) {
  const text = String(val || '').trim();
  if (/^_\d+$/.test(text)) return text;
  return /^\d+$/.test(text) ? '_' + text : text;
}

function normalizeThreeDigitCode(val) {
  const text = String(val || '').trim();
  const raw = text.startsWith('_') ? text.slice(1) : text;
  if (!/^\d+$/.test(raw)) return text;
  const padded = raw.length < 3 ? raw.padStart(3, '0') : raw;
  return '_' + padded;
}

function sheetText(val) {
  return String(val || '').trim();
}

function sheetCodeValue(val) {
  return String(val || '').trim();
}

function parseHours(val) {
  const text = String(val || '').replace(/[^\d.]/g, '').trim();
  if (!text) return 0;
  const num = Number(text);
  return isFinite(num) ? num : 0;
}

function normalizeReportType(val) {
  const text = String(val || '').trim();
  return text || '定期';
}

function findHeaderIndex_(headers, names) {
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').replace(/\s+/g, '').trim();
    if (names.includes(header)) return i;
  }
  return -1;
}

function extractSpreadsheetId(input) {
  const text = String(input || '').replace(/\s+/g, '').trim();
  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : text;
}

function findUser(userId) {
  const data = usersSheet().getDataRange().getDisplayValues();
  const normalizedUserId = normalizeEmployeeId(userId);
  for (let i = 1; i < data.length; i++) {
    if (normalizeEmployeeId(data[i][0]) === normalizedUserId) {
      return {
        id:    normalizeEmployeeId(data[i][0]),
        name:  String(data[i][1]),
        pwd:   String(data[i][2]),
        role:  String(data[i][3]),
        group: String(data[i][4] || '')
      };
    }
  }
  return null;
}

function createSession(userId) {
  const token  = Utilities.getUuid();
  const expiry = Date.now() + 365 * 24 * 60 * 60 * 1000;
  appendTextRow(sessionsSheet(), [token, String(userId), expiry], [1, 2]);
  return token;
}

function validateToken(token) {
  if (!token) return null;
  const sheet = sessionsSheet();
  const data  = sheet.getDataRange().getValues();
  const now   = Date.now();
  for (let i = data.length - 1; i >= 1; i--) {
    if (Number(data[i][2]) < now) { sheet.deleteRow(i + 1); data.splice(i, 1); }
  }
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) return String(data[i][1]);
  }
  return null;
}

function deleteToken(token) {
  const sheet = sessionsSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) { sheet.deleteRow(i + 1); return; }
  }
}

function getGroupNames() {
  const ws = groupsSheet();
  if (!ws) return [];
  const data = ws.getDataRange().getDisplayValues();
  return data.slice(1)
    .map(r => String(r[0] || '').trim())
    .filter(Boolean)
    .sort();
}

function groupExists(groupName) {
  const group = String(groupName || '').trim();
  return !!group && getGroupNames().includes(group);
}

// ── Auth actions ──────────────────────────────────────────────────────────────

function getGroups(p) {
  return jsonResponse({ status: 'success', groups: getGroupNames() });
}

function register(p) {
  if (!p.id || !p.name || !p.pwd || !p.group)
    return jsonResponse({ status: 'error', message: '缺少必要參數（工號、姓名、組別、密碼）' });
  const employeeId = normalizeEmployeeId(p.id);
  if (!groupExists(p.group)) return jsonResponse({ status: 'error', message: '請選擇有效組別' });
  if (findUser(employeeId)) return jsonResponse({ status: 'error', message: '工號已存在' });
  appendTextRow(usersSheet(), [employeeId, String(p.name), String(p.pwd), 'employee', String(p.group)], [1, 2, 3, 4, 5]);
  return jsonResponse({ status: 'success', message: '註冊成功，請登入' });
}

function registerManager(p) {
  if (!p.id || !p.name || !p.pwd || !p.invite)
    return jsonResponse({ status: 'error', message: '缺少必要參數' });
  const inviteCode = PropertiesService.getScriptProperties().getProperty('MANAGER_INVITE_CODE');
  if (!inviteCode) return jsonResponse({ status: 'error', message: 'MANAGER_INVITE_CODE 未設定' });
  if (p.invite !== inviteCode) return jsonResponse({ status: 'error', message: '邀請碼不正確' });
  const employeeId = normalizeEmployeeId(p.id);
  if (findUser(employeeId)) return jsonResponse({ status: 'error', message: '工號已存在' });
  appendTextRow(usersSheet(), [employeeId, String(p.name), String(p.pwd), 'manager', ''], [1, 2, 3, 4, 5]);
  return jsonResponse({ status: 'success', message: '主管帳號建立成功，請登入' });
}

function login(p) {
  if (!p.id || !p.pwd) return jsonResponse({ status: 'error', message: '缺少必要參數' });
  const user = findUser(normalizeEmployeeId(p.id));
  if (!user || user.pwd !== p.pwd) return jsonResponse({ status: 'error', message: '工號或密碼錯誤' });
  const token = createSession(user.id);
  return jsonResponse({ status: 'success', token, role: user.role, name: user.name, group: user.group });
}

function logout(p) {
  if (p.token) deleteToken(p.token);
  return jsonResponse({ status: 'success' });
}

// ── Employee actions ──────────────────────────────────────────────────────────

function submitReport(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  if (!p.date || !p.workOrders || !p.shipNo) return jsonResponse({ status: 'error', message: '請填寫施工日期、工單號碼和船號' });
  const hours = parseHours(p.hours);
  if (hours <= 0) return jsonResponse({ status: 'error', message: '請填寫實際工時' });
  const reportType = normalizeReportType(p.reportType);

  const user = findUser(userId);
  if (!user) return jsonResponse({ status: 'error', message: '找不到員工資料' });

  const orders = String(p.workOrders).split(',').map(s => s.trim()).filter(Boolean);
  if (!orders.length) return jsonResponse({ status: 'error', message: '請至少填入一個工單號碼' });

  const sheet = reportsSheet();
  const now   = new Date().toISOString();
  const shipNo = normalizeShipNo(p.shipNo);
  orders.forEach(order => {
    const row = sheet.getLastRow() + 1;
    setPlainTextValues(sheet, row, 1, [[
      now,
      String(p.date),
      user.id,
      user.name,
      user.group,
      sheetCodeValue(normalizeWorkOrderId(order)),
      sheetCodeValue(shipNo),
      String(hours),
      reportType
    ]]);
  });
  return jsonResponse({ status: 'success', message: `報工已儲存（共 ${orders.length} 筆）` });
}

function getMyReports(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });

  const range = reportsSheet().getDataRange();
  const data = range.getValues();
  const displayData = range.getDisplayValues();
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const displayRow = displayData[i];
    if (normalizeEmployeeId(displayRow[2]) !== normalizeEmployeeId(userId) || !row[1]) continue;
    const dateStr = toDateStr(row[1]);
    if (p.year && !dateStr.startsWith(p.year)) continue;
    if (p.month) {
      const m = String(new Date(dateStr).getMonth() + 1).padStart(2, '0');
      if (m !== String(p.month).padStart(2, '0')) continue;
    }
    records.push({
      date:       dateStr,
      empId:      normalizeEmployeeId(displayRow[2]),
      group:      String(displayRow[4] || ''),
      workOrders: String(displayRow[5] || '').split(',').map(s => normalizeWorkOrderId(s)).filter(Boolean),
      shipNo:     normalizeShipNo(displayRow[6]),
      hours:      parseHours(displayRow[7]),
      reportType: normalizeReportType(displayRow[8])
    });
  }
  records.sort((a, b) => b.date.localeCompare(a.date));
  return jsonResponse({ status: 'success', records });
}

function getWorkOrders(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const ws = workOrdersSheet();
  if (!ws) return jsonResponse({ status: 'success', workOrders: [] });
  const data = ws.getDataRange().getDisplayValues();
  const legacyLayout = String(data[0] && data[0][2] || '') === '建立時間';
  const workOrders = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    // [工單號碼, 船號, 工單內容, 預估工時, 建立時間, 備註]
    workOrders.push({
      id:     normalizeWorkOrderId(data[i][0]),
      shipNo: normalizeShipNo(data[i][1]),
      content: legacyLayout ? '' : String(data[i][2] || ''),
      estimatedHours: legacyLayout ? '' : String(data[i][3] || ''),
      remark: legacyLayout ? String(data[i][3] || '') : String(data[i][5] || '')
    });
  }
  return jsonResponse({ status: 'success', workOrders });
}

// ── Manager: work order management ───────────────────────────────────────────

function addWorkOrder(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });
  if (!p.workOrderId) return jsonResponse({ status: 'error', message: '請輸入工單號碼' });
  if (!p.shipNo)      return jsonResponse({ status: 'error', message: '請輸入對應船號' });
  const workOrderId = normalizeWorkOrderId(p.workOrderId);
  const shipNo = normalizeShipNo(p.shipNo);
  const content = String(p.content || '');
  const estimatedHours = String(p.estimatedHours || '');

  const ws   = workOrdersSheet();
  const data = ws.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizeWorkOrderId(data[i][0]) === workOrderId)
      return jsonResponse({ status: 'error', message: '工單號碼已存在' });
  }
  // [工單號碼, 船號, 工單內容, 預估工時, 建立時間, 備註]
  const row = ws.getLastRow() + 1;
  setPlainTextValues(ws, row, 1, [[sheetCodeValue(workOrderId), sheetCodeValue(shipNo), content, estimatedHours, new Date().toISOString(), String(p.remark || '')]]);
  return jsonResponse({ status: 'success', message: '工單已新增' });
}

function deleteWorkOrder(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });
  if (!p.workOrderId) return jsonResponse({ status: 'error', message: '請指定工單號碼' });

  const ws   = workOrdersSheet();
  const data = ws.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizeWorkOrderId(data[i][0]) === normalizeWorkOrderId(p.workOrderId)) {
      ws.deleteRow(i + 1);
      return jsonResponse({ status: 'success', message: '工單已刪除' });
    }
  }
  return jsonResponse({ status: 'error', message: '找不到工單' });
}

// p.items = JSON string of [{id, shipNo, content?, estimatedHours?, remark?}, ...]
function importWorkOrders(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });
  if (!p.items) return jsonResponse({ status: 'error', message: '缺少資料' });

  let items;
  try { items = JSON.parse(p.items); } catch (e) {
    return jsonResponse({ status: 'error', message: '資料格式錯誤' });
  }
  if (!Array.isArray(items) || !items.length)
    return jsonResponse({ status: 'error', message: '無有效工單資料' });

  const ws       = workOrdersSheet();
  normalizeWorkOrdersSheetCodes_();
  const existing = ws.getDataRange().getDisplayValues();
  const existSet = new Set(existing.slice(1).map(r => normalizeWorkOrderId(r[0])));
  const now      = new Date().toISOString();
  let added = 0, skipped = 0;

  for (const item of items) {
    const id     = normalizeWorkOrderId(item.id);
    const shipNo = normalizeShipNo(item.shipNo);
    const content = String(item.content || '');
    const estimatedHours = String(item.estimatedHours || '');
    const remark = String(item.remark || '');
    if (!id) continue;
    if (existSet.has(id)) { skipped++; continue; }
    // [工單號碼, 船號, 工單內容, 預估工時, 建立時間, 備註]
    const row = ws.getLastRow() + 1;
    setPlainTextValues(ws, row, 1, [[sheetCodeValue(id), sheetCodeValue(shipNo), content, estimatedHours, now, remark]]);
    existSet.add(id);
    added++;
  }
  return jsonResponse({ status: 'success', message: `匯入完成：新增 ${added} 筆，略過重複 ${skipped} 筆` });
}

// A欄 = 組別名稱，預設第 2 列起；若指定分頁只讀該分頁，否則讀全部分頁。
function importGroups(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });
  if (!p.sheetUrl) return jsonResponse({ status: 'error', message: '請提供 Google Sheet 網址' });

  const sheetId = extractSpreadsheetId(p.sheetUrl);
  let ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (e) {
    return jsonResponse({ status: 'error', message: '無法開啟 Google Sheet，請確認網址正確且已分享給此腳本帳號' });
  }

  const col = Math.max(1, parseInt(p.col || '1', 10)) - 1;
  const startRow = Math.max(1, parseInt(p.startRow || '2', 10)) - 1;
  const sheets = p.sheetName
    ? [ss.getSheetByName(p.sheetName)].filter(Boolean)
    : ss.getSheets();
  if (!sheets.length) return jsonResponse({ status: 'error', message: '找不到指定的工作表分頁' });

  const targetWs = groupsSheet();
  if (!targetWs) return jsonResponse({ status: 'error', message: '找不到 Groups 分頁，請先執行 setupSheets' });
  const existing = getGroupNames();
  const existSet = new Set(existing);
  const now = new Date().toISOString();
  const rowsToAppend = [];
  let skipped = 0;

  for (const ws of sheets) {
    const lastRow = ws.getLastRow();
    if (lastRow <= startRow) continue;
    const data = ws.getRange(1, 1, lastRow, col + 1).getDisplayValues();
    for (let i = startRow; i < data.length; i++) {
      const group = String(data[i][col] || '').trim();
      if (!group) continue;
      if (existSet.has(group)) { skipped++; continue; }
      rowsToAppend.push([sheetText(group), now]);
      existSet.add(group);
    }
  }

  appendTextRows(targetWs, rowsToAppend, [1, 2]);
  return jsonResponse({ status: 'success', message: `組別匯入完成：新增 ${rowsToAppend.length} 筆，略過重複 ${skipped} 筆` });
}

// 預設：分頁名稱 = 船號，A欄（或 p.col 指定欄）= 工單號碼
// 若提供 p.shipCol，則從該欄讀取船號，適合同一張表包含多個船號的匯入格式。
// p.sheetUrl  - Google Sheet URL or Sheet ID
// p.col       - 1-based column index for work order IDs (default 1)
// p.shipCol   - 1-based column index for ship numbers (optional)
// p.contentCol - 1-based column index for work order content (optional)
// p.estimateCol - 1-based column index for estimated hours (optional)
// p.remarkCol - 1-based column index for remarks (optional)
// p.startRow  - 1-based first data row (default 2, skipping header)
// 若 p.sheetName 指定，只匯入該分頁；否則匯入所有分頁
function testImportSheet(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });
  if (!p.sheetUrl) return jsonResponse({ status: 'error', message: '請提供 Google Sheet 網址' });

  const sheetId = extractSpreadsheetId(p.sheetUrl);
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheets = ss.getSheets().map(s => ({
      name: s.getName(),
      rows: s.getLastRow()
    }));
    return jsonResponse({
      status: 'success',
      message: `可開啟來源 Sheet：${ss.getName()}，共 ${sheets.length} 個分頁`,
      sheetId,
      sheets
    });
  } catch (e) {
    return jsonResponse({
      status: 'error',
      message: `無法開啟來源 Sheet：${e.message}`
    });
  }
}

function importFromSheet(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });
  if (!p.sheetUrl) return jsonResponse({ status: 'error', message: '請提供 Google Sheet 網址' });

  const sheetId = extractSpreadsheetId(p.sheetUrl);

  let ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (e) {
    return jsonResponse({ status: 'error', message: '無法開啟 Google Sheet，請確認網址正確、沒有缺漏，且已分享給此腳本帳號' });
  }

  const col        = Math.max(1, parseInt(p.col        || '1', 10)) - 1;
  const shipCol    = p.shipCol ? Math.max(1, parseInt(p.shipCol, 10)) - 1 : -1;
  const contentCol = p.contentCol ? Math.max(1, parseInt(p.contentCol, 10)) - 1 : -1;
  const estimateCol= p.estimateCol ? Math.max(1, parseInt(p.estimateCol, 10)) - 1 : -1;
  const remarkCol  = p.remarkCol ? Math.max(1, parseInt(p.remarkCol, 10)) - 1 : -1;
  const startRow   = Math.max(1, parseInt(p.startRow || '2', 10)) - 1;

  // 選擇要匯入的分頁：指定單一分頁 or 全部分頁。若指定名稱不存在但只有一個分頁，回退讀取該分頁。
  let sheets = [];
  if (p.sheetName) {
    const specified = ss.getSheetByName(p.sheetName);
    sheets = specified ? [specified] : (ss.getSheets().length === 1 ? ss.getSheets() : []);
  } else {
    sheets = ss.getSheets();
  }

  if (!sheets.length) return jsonResponse({ status: 'error', message: '找不到指定的工作表分頁' });

  const targetWs = workOrdersSheet();
  normalizeWorkOrdersSheetCodes_();
  const targetLastRow = targetWs.getLastRow();
  const existing = targetLastRow > 1
    ? targetWs.getRange(2, 1, targetLastRow - 1, 1).getDisplayValues()
    : [];
  const existSet = new Set(existing.map(r => normalizeWorkOrderId(r[0])));
  const now      = new Date().toISOString();
  const rowsToAppend = [];
  let   skipped = 0;

  for (const ws of sheets) {
    const defaultShipNo = normalizeShipNo(ws.getName()); // 未指定船號欄時，分頁名稱即為船號
    const lastRow = ws.getLastRow();
    if (lastRow <= startRow) continue;

    let effectiveCol = col;
    let effectiveShipCol = shipCol;
    let effectiveContentCol = contentCol;
    let effectiveEstimateCol = estimateCol;
    let effectiveRemarkCol = remarkCol;
    let effectiveStartRow = startRow;

    const headerWidth = Math.max(col, shipCol, contentCol, estimateCol, remarkCol, 8) + 1;
    const headers = ws.getRange(1, 1, 1, headerWidth).getDisplayValues()[0];
    const detectedShipCol = findHeaderIndex_(headers, ['船號']);
    const detectedOrderCol = findHeaderIndex_(headers, ['工單', '工單號碼']);
    const detectedContentCol = findHeaderIndex_(headers, ['工單內容', '內容']);
    const detectedEstimateCol = findHeaderIndex_(headers, ['預估工時', '預估時數']);
    const detectedRemarkCol = findHeaderIndex_(headers, ['備註', '備註(選填)']);
    if (detectedOrderCol >= 0) {
      effectiveCol = detectedOrderCol;
      effectiveStartRow = Math.max(effectiveStartRow, 1);
    }
    if (detectedShipCol >= 0) effectiveShipCol = detectedShipCol;
    if (detectedContentCol >= 0) effectiveContentCol = detectedContentCol;
    if (detectedEstimateCol >= 0) effectiveEstimateCol = detectedEstimateCol;
    if (detectedRemarkCol >= 0) effectiveRemarkCol = detectedRemarkCol;

    const readWidth = Math.max(effectiveCol, effectiveShipCol, effectiveContentCol, effectiveEstimateCol, effectiveRemarkCol, 0) + 1;
    const data = ws.getRange(1, 1, lastRow, readWidth).getDisplayValues();

    for (let i = effectiveStartRow; i < data.length; i++) {
      const id = normalizeWorkOrderId(data[i][effectiveCol]);
      const shipNo = effectiveShipCol >= 0 ? normalizeShipNo(data[i][effectiveShipCol]) : defaultShipNo;
      if (!id) continue;
      if (!shipNo) continue;
      if (existSet.has(id)) { skipped++; continue; }
      const content = effectiveContentCol >= 0 ? String(data[i][effectiveContentCol] || '') : '';
      const estimatedHours = effectiveEstimateCol >= 0 ? String(data[i][effectiveEstimateCol] || '') : '';
      const remark = effectiveRemarkCol >= 0 ? String(data[i][effectiveRemarkCol] || '') : '';
      // [工單號碼, 船號, 工單內容, 預估工時, 建立時間, 備註]
      rowsToAppend.push([sheetCodeValue(id), sheetCodeValue(shipNo), content, estimatedHours, now, remark]);
      existSet.add(id);
    }
  }
  if (rowsToAppend.length) {
    const appendStartRow = targetWs.getLastRow() + 1;
    setPlainTextValues(targetWs, appendStartRow, 1, rowsToAppend);
    normalizeWorkOrdersSheetCodes_();
  }
  const added = rowsToAppend.length;
  return jsonResponse({ status: 'success', message: `匯入完成：新增 ${added} 筆，略過重複 ${skipped} 筆` });
}

// ── Manager: report queries ───────────────────────────────────────────────────

// Optional filters: startDate, endDate, group, empId, workOrder, shipNo, reportType
// Optional sortBy: date (default desc) | workOrder | group | empId | shipNo | hours | reportType
function getAllReports(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });

  const range = reportsSheet().getDataRange();
  const data = range.getValues();
  const displayData = range.getDisplayValues();
  let   records = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const displayRow = displayData[i];
    if (!row[1]) continue;
    const dateStr = toDateStr(row[1]);
    if (p.startDate  && dateStr < p.startDate)            continue;
    if (p.endDate    && dateStr > p.endDate)              continue;
    if (p.group      && String(displayRow[4]) !== p.group) continue;
    if (p.empId      && normalizeEmployeeId(displayRow[2]) !== normalizeEmployeeId(p.empId)) continue;
    if (p.workOrder) {
      const wos = String(displayRow[5] || '').split(',').map(s => normalizeWorkOrderId(s)).filter(Boolean);
      if (!wos.includes(normalizeWorkOrderId(p.workOrder))) continue;
    }
    if (p.shipNo && normalizeShipNo(displayRow[6]) !== normalizeShipNo(p.shipNo)) continue;
    if (p.reportType && normalizeReportType(displayRow[8]) !== normalizeReportType(p.reportType)) continue;
    records.push({
      date:       dateStr,
      empId:      normalizeEmployeeId(displayRow[2]),
      name:       String(displayRow[3] || ''),
      group:      String(displayRow[4] || ''),
      workOrders: String(displayRow[5] || '').split(',').map(s => normalizeWorkOrderId(s)).filter(Boolean),
      shipNo:     normalizeShipNo(displayRow[6]),
      hours:      parseHours(displayRow[7]),
      reportType: normalizeReportType(displayRow[8])
    });
  }

  const sortBy = p.sortBy || 'date';
  if      (sortBy === 'workOrder') records.sort((a, b) => (a.workOrders[0]||'').localeCompare(b.workOrders[0]||'') || b.date.localeCompare(a.date));
  else if (sortBy === 'group')     records.sort((a, b) => a.group.localeCompare(b.group)         || b.date.localeCompare(a.date));
  else if (sortBy === 'empId')     records.sort((a, b) => a.empId.localeCompare(b.empId)         || b.date.localeCompare(a.date));
  else if (sortBy === 'shipNo')    records.sort((a, b) => a.shipNo.localeCompare(b.shipNo)       || b.date.localeCompare(a.date));
  else if (sortBy === 'hours')     records.sort((a, b) => b.hours - a.hours                       || b.date.localeCompare(a.date));
  else if (sortBy === 'reportType')records.sort((a, b) => a.reportType.localeCompare(b.reportType)|| b.date.localeCompare(a.date));
  else                             records.sort((a, b) => b.date.localeCompare(a.date));

  return jsonResponse({ status: 'success', records });
}

function getAllEmployees(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });

  const data      = usersSheet().getDataRange().getDisplayValues();
  const employees = [];
  const groupSet  = new Set(getGroupNames());
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] !== 'employee') continue;
    employees.push({ id: normalizeEmployeeId(data[i][0]), name: String(data[i][1]), group: String(data[i][4] || '') });
    if (data[i][4]) groupSet.add(String(data[i][4]));
  }
  return jsonResponse({ status: 'success', employees, groups: Array.from(groupSet).sort() });
}
