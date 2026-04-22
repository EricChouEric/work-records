// Script Properties:
//   MASTER_SHEET_ID     - Master Google Sheet ID
//   MANAGER_INVITE_CODE - Secret code for manager registration
//
// Master Sheet tabs required:
//   Users      : [工號, 姓名, 密碼, 角色, 組別]
//   WorkOrders : [工單號碼, 建立時間, 備註]
//   Reports    : [提交時間, 施工日期, 工號, 員工姓名, 組別, 工單號碼]
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
      case 'submitReport':     return submitReport(e.parameter);
      case 'getMyReports':     return getMyReports(e.parameter);
      case 'getWorkOrders':    return getWorkOrders(e.parameter);
      case 'addWorkOrder':     return addWorkOrder(e.parameter);
      case 'deleteWorkOrder':  return deleteWorkOrder(e.parameter);
      case 'importWorkOrders': return importWorkOrders(e.parameter);
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
function sessionsSheet()   { return getMaster().getSheetByName('Sessions'); }
function workOrdersSheet() { return getMaster().getSheetByName('WorkOrders'); }
function reportsSheet()    { return getMaster().getSheetByName('Reports'); }

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
    }
    return ws;
  }
  ensure('Users',      ['工號', '姓名', '密碼', '角色', '組別']);
  ensure('Sessions',   ['token', 'userId', 'expiry']);
  ensure('WorkOrders', ['工單號碼', '船號', '建立時間', '備註']);
  ensure('Reports',    ['提交時間', '施工日期', '工號', '員工姓名', '組別', '工單號碼', '船號']);
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

// ── Utilities ─────────────────────────────────────────────────────────────────

function toDateStr(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

function findUser(userId) {
  const data = usersSheet().getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      return {
        id:    String(data[i][0]),
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
  sessionsSheet().appendRow([token, String(userId), expiry]);
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

// ── Auth actions ──────────────────────────────────────────────────────────────

function register(p) {
  if (!p.id || !p.name || !p.pwd || !p.group)
    return jsonResponse({ status: 'error', message: '缺少必要參數（工號、姓名、組別、密碼）' });
  if (findUser(p.id)) return jsonResponse({ status: 'error', message: '工號已存在' });
  usersSheet().appendRow([p.id, p.name, p.pwd, 'employee', p.group]);
  return jsonResponse({ status: 'success', message: '註冊成功，請登入' });
}

function registerManager(p) {
  if (!p.id || !p.name || !p.pwd || !p.invite)
    return jsonResponse({ status: 'error', message: '缺少必要參數' });
  const inviteCode = PropertiesService.getScriptProperties().getProperty('MANAGER_INVITE_CODE');
  if (!inviteCode) return jsonResponse({ status: 'error', message: 'MANAGER_INVITE_CODE 未設定' });
  if (p.invite !== inviteCode) return jsonResponse({ status: 'error', message: '邀請碼不正確' });
  if (findUser(p.id)) return jsonResponse({ status: 'error', message: '工號已存在' });
  usersSheet().appendRow([p.id, p.name, p.pwd, 'manager', '']);
  return jsonResponse({ status: 'success', message: '主管帳號建立成功，請登入' });
}

function login(p) {
  if (!p.id || !p.pwd) return jsonResponse({ status: 'error', message: '缺少必要參數' });
  const user = findUser(p.id);
  if (!user || user.pwd !== p.pwd) return jsonResponse({ status: 'error', message: '工號或密碼錯誤' });
  const token = createSession(p.id);
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

  const user = findUser(userId);
  if (!user) return jsonResponse({ status: 'error', message: '找不到員工資料' });

  const orders = String(p.workOrders).split(',').map(s => s.trim()).filter(Boolean);
  if (!orders.length) return jsonResponse({ status: 'error', message: '請至少填入一個工單號碼' });

  const sheet = reportsSheet();
  const now   = new Date().toISOString();
  orders.forEach(order => {
    sheet.appendRow([now, p.date, user.id, user.name, user.group, order, p.shipNo]);
  });
  return jsonResponse({ status: 'success', message: `報工已儲存（共 ${orders.length} 筆）` });
}

function getMyReports(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });

  const data = reportsSheet().getDataRange().getValues();
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[2]) !== userId || !row[1]) continue;
    const dateStr = toDateStr(row[1]);
    if (p.year && !dateStr.startsWith(p.year)) continue;
    if (p.month) {
      const m = String(new Date(dateStr).getMonth() + 1).padStart(2, '0');
      if (m !== String(p.month).padStart(2, '0')) continue;
    }
    records.push({
      date:       dateStr,
      empId:      String(row[2] || ''),
      group:      String(row[4] || ''),
      workOrders: String(row[5] || '').split(',').map(s => s.trim()).filter(Boolean),
      shipNo:     String(row[6] || '')
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
  const data = ws.getDataRange().getValues();
  const workOrders = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    // [工單號碼, 船號, 建立時間, 備註]
    workOrders.push({
      id:     String(data[i][0]),
      shipNo: String(data[i][1] || ''),
      remark: String(data[i][3] || '')
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

  const ws   = workOrdersSheet();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.workOrderId))
      return jsonResponse({ status: 'error', message: '工單號碼已存在' });
  }
  // [工單號碼, 船號, 建立時間, 備註]
  ws.appendRow([p.workOrderId, p.shipNo, new Date().toISOString(), p.remark || '']);
  return jsonResponse({ status: 'success', message: '工單已新增' });
}

function deleteWorkOrder(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });
  if (!p.workOrderId) return jsonResponse({ status: 'error', message: '請指定工單號碼' });

  const ws   = workOrdersSheet();
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.workOrderId)) {
      ws.deleteRow(i + 1);
      return jsonResponse({ status: 'success', message: '工單已刪除' });
    }
  }
  return jsonResponse({ status: 'error', message: '找不到工單' });
}

// p.items = JSON string of [{id, shipNo, remark?}, ...]
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
  const existing = ws.getDataRange().getValues();
  const existSet = new Set(existing.slice(1).map(r => String(r[0])));
  const now      = new Date().toISOString();
  let added = 0, skipped = 0;

  for (const item of items) {
    const id     = String(item.id     || '').trim();
    const shipNo = String(item.shipNo || '').trim();
    const remark = String(item.remark || '');
    if (!id) continue;
    if (existSet.has(id)) { skipped++; continue; }
    // [工單號碼, 船號, 建立時間, 備註]
    ws.appendRow([id, shipNo, now, remark]);
    existSet.add(id);
    added++;
  }
  return jsonResponse({ status: 'success', message: `匯入完成：新增 ${added} 筆，略過重複 ${skipped} 筆` });
}

// 分頁名稱 = 船號，A欄（或 p.col 指定欄）= 工單號碼
// p.sheetUrl  - Google Sheet URL or Sheet ID
// p.col       - 1-based column index for work order IDs (default 1)
// p.remarkCol - 1-based column index for remarks (optional)
// p.startRow  - 1-based first data row (default 2, skipping header)
// 若 p.sheetName 指定，只匯入該分頁；否則匯入所有分頁
function importFromSheet(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });
  if (!p.sheetUrl) return jsonResponse({ status: 'error', message: '請提供 Google Sheet 網址' });

  let sheetId = p.sheetUrl.trim();
  const match = sheetId.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) sheetId = match[1];

  let ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (e) {
    return jsonResponse({ status: 'error', message: '無法開啟 Google Sheet，請確認網址正確且已分享給此腳本帳號' });
  }

  const col      = Math.max(1, parseInt(p.col      || '1', 10)) - 1;
  const remarkCol= p.remarkCol ? Math.max(1, parseInt(p.remarkCol, 10)) - 1 : -1;
  const startRow = Math.max(1, parseInt(p.startRow || '2', 10)) - 1;

  // 選擇要匯入的分頁：指定單一分頁 or 全部分頁
  const sheets = p.sheetName
    ? [ss.getSheetByName(p.sheetName)].filter(Boolean)
    : ss.getSheets();

  if (!sheets.length) return jsonResponse({ status: 'error', message: '找不到指定的工作表分頁' });

  const targetWs = workOrdersSheet();
  const existing = targetWs.getDataRange().getValues();
  const existSet = new Set(existing.slice(1).map(r => String(r[0])));
  const now      = new Date().toISOString();
  let   added = 0, skipped = 0;

  for (const ws of sheets) {
    const shipNo = ws.getName(); // 分頁名稱即為船號
    const data   = ws.getDataRange().getValues();

    for (let i = startRow; i < data.length; i++) {
      const id = String(data[i][col] || '').trim();
      if (!id) continue;
      if (existSet.has(id)) { skipped++; continue; }
      const remark = remarkCol >= 0 ? String(data[i][remarkCol] || '') : '';
      // [工單號碼, 船號, 建立時間, 備註]
      targetWs.appendRow([id, shipNo, now, remark]);
      existSet.add(id);
      added++;
    }
  }
  return jsonResponse({ status: 'success', message: `匯入完成：新增 ${added} 筆，略過重複 ${skipped} 筆` });
}

// ── Manager: report queries ───────────────────────────────────────────────────

// Optional filters: startDate, endDate, group, empId, workOrder, shipNo
// Optional sortBy: date (default desc) | workOrder | group | empId | shipNo
function getAllReports(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });

  const data    = reportsSheet().getDataRange().getValues();
  let   records = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    const dateStr = toDateStr(row[1]);
    if (p.startDate  && dateStr < p.startDate)            continue;
    if (p.endDate    && dateStr > p.endDate)              continue;
    if (p.group      && String(row[4]) !== p.group)       continue;
    if (p.empId      && String(row[2]) !== p.empId)       continue;
    if (p.workOrder) {
      const wos = String(row[5] || '').split(',').map(s => s.trim());
      if (!wos.includes(p.workOrder)) continue;
    }
    if (p.shipNo && String(row[6]) !== p.shipNo) continue;
    records.push({
      date:       dateStr,
      empId:      String(row[2] || ''),
      name:       String(row[3] || ''),
      group:      String(row[4] || ''),
      workOrders: String(row[5] || '').split(',').map(s => s.trim()).filter(Boolean),
      shipNo:     String(row[6] || '')
    });
  }

  const sortBy = p.sortBy || 'date';
  if      (sortBy === 'workOrder') records.sort((a, b) => (a.workOrders[0]||'').localeCompare(b.workOrders[0]||'') || b.date.localeCompare(a.date));
  else if (sortBy === 'group')     records.sort((a, b) => a.group.localeCompare(b.group)         || b.date.localeCompare(a.date));
  else if (sortBy === 'empId')     records.sort((a, b) => a.empId.localeCompare(b.empId)         || b.date.localeCompare(a.date));
  else if (sortBy === 'shipNo')    records.sort((a, b) => a.shipNo.localeCompare(b.shipNo)       || b.date.localeCompare(a.date));
  else                             records.sort((a, b) => b.date.localeCompare(a.date));

  return jsonResponse({ status: 'success', records });
}

function getAllEmployees(p) {
  const userId = validateToken(p.token);
  if (!userId) return jsonResponse({ status: 'error', message: '登入逾時，請重新登入' });
  const user = findUser(userId);
  if (!user || user.role !== 'manager') return jsonResponse({ status: 'error', message: '無權限' });

  const data      = usersSheet().getDataRange().getValues();
  const employees = [];
  const groupSet  = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] !== 'employee') continue;
    employees.push({ id: String(data[i][0]), name: String(data[i][1]), group: String(data[i][4] || '') });
    if (data[i][4]) groupSet.add(String(data[i][4]));
  }
  return jsonResponse({ status: 'success', employees, groups: Array.from(groupSet).sort() });
}
