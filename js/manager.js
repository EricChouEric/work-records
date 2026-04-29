let session;
let allEmployees = [];
let allGroups    = [];
let allWorkOrders = [];

function normalizeNumericInputValue(value) {
  return String(value || '').trim().replace(/^_(?=\d+$)/, '');
}

async function init() {
  session = requireAuth('manager');
  if (!session) return;

  document.getElementById('userName').textContent = session.name;

  setupNumericInputNormalizers();
  setupTabs();
  await Promise.all([loadEmployeeFilters(), loadWorkOrderList(), loadGroups()]);
}

function setupNumericInputNormalizers() {
  ['newWoId', 'newWoShip'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('blur', () => {
      el.value = normalizeNumericInputValue(el.value);
    });
  });
}

function setupTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.add('hidden'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.panel).classList.remove('hidden');
    });
  });
}

// ── Work Order Management ────────────────────────────────────────────────────

async function loadWorkOrderList() {
  const wrap = document.getElementById('woListWrap');
  wrap.innerHTML = '<div class="loading">載入中...</div>';
  try {
    const result = await callAPI({ action: 'getWorkOrders', token: session.token });
    if (result.status !== 'success') {
      wrap.innerHTML = `<div class="empty">${result.message}</div>`;
      return;
    }
    allWorkOrders = result.workOrders;
    renderWorkOrderList(allWorkOrders);
    refreshWorkOrderFilter();
  } catch (e) {
    wrap.innerHTML = '<div class="empty">載入失敗</div>';
  }
}

function renderWorkOrderList(workOrders) {
  const wrap = document.getElementById('woListWrap');
  if (!workOrders.length) {
    wrap.innerHTML = '<div class="empty">尚無工單，請新增或匯入</div>';
    return;
  }
  wrap.innerHTML = `
    <div class="table-wrap">
      <table>
        <thead>
          <tr><th>#</th><th>工單號碼</th><th>船號</th><th>工單內容</th><th>預估工時</th><th>備註</th><th></th></tr>
        </thead>
        <tbody>
          ${workOrders.map((wo, i) => `
            <tr>
              <td style="color:var(--text-muted)">${i + 1}</td>
              <td><strong>${escHtml(wo.id)}</strong></td>
              <td><span class="badge badge-blue">${escHtml(wo.shipNo)}</span></td>
              <td>${escHtml(wo.content || '')}</td>
              <td>${escHtml(wo.estimatedHours || '')}</td>
              <td>${escHtml(wo.remark)}</td>
              <td><button class="btn btn-sm" style="color:#dc2626;border:1px solid #fecaca;background:#fef2f2"
                onclick="deleteWorkOrder('${escAttr(wo.id)}')">刪除</button></td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
}

async function addWorkOrder() {
  const idInput = document.getElementById('newWoId');
  const shipInput = document.getElementById('newWoShip');
  const id     = normalizeNumericInputValue(idInput.value);
  const shipNo = normalizeNumericInputValue(shipInput.value);
  const content = document.getElementById('newWoContent').value.trim();
  const estimatedHours = document.getElementById('newWoEstimate').value.trim();
  const remark = document.getElementById('newWoRemark').value.trim();
  const msgEl  = document.getElementById('addWoMsg');
  idInput.value = id;
  shipInput.value = shipNo;
  if (!id)     { showMsg(msgEl, 'error', '請輸入工單號碼'); return; }
  if (!shipNo) { showMsg(msgEl, 'error', '請輸入對應船號'); return; }

  try {
    const result = await callAPI({ action: 'addWorkOrder', token: session.token, workOrderId: id, shipNo, content, estimatedHours, remark });
    if (result.status === 'success') {
      showMsg(msgEl, 'success', result.message);
      document.getElementById('newWoId').value    = '';
      document.getElementById('newWoShip').value  = '';
      document.getElementById('newWoContent').value = '';
      document.getElementById('newWoEstimate').value = '';
      document.getElementById('newWoRemark').value = '';
      await loadWorkOrderList();
    } else {
      showMsg(msgEl, 'error', result.message);
    }
  } catch (e) {
    showMsg(msgEl, 'error', '連線失敗');
  }
}

async function deleteWorkOrder(workOrderId) {
  if (!confirm(`確定要刪除工單「${workOrderId}」嗎？`)) return;
  try {
    const result = await callAPI({ action: 'deleteWorkOrder', token: session.token, workOrderId });
    if (result.status === 'success') {
      await loadWorkOrderList();
    } else {
      alert(result.message);
    }
  } catch (e) {
    alert('連線失敗');
  }
}

async function importFromSheet() {
  const sheetInput = document.getElementById('importSheetUrl').value;
  const sheetUrl   = extractSpreadsheetId(sheetInput);
  const sheetName  = document.getElementById('importSheetName').value.trim();
  const startRow   = document.getElementById('importStartRow').value;
  const shipCol    = document.getElementById('importShipCol').value.trim();
  const col        = document.getElementById('importCol').value;
  const contentCol = document.getElementById('importContentCol').value.trim();
  const estimateCol= document.getElementById('importEstimateCol').value.trim();
  const remarkCol  = document.getElementById('importRemarkCol').value.trim();
  const msgEl      = document.getElementById('importMsg');

  if (!sheetUrl) { showMsg(msgEl, 'error', '請輸入 Google Sheet 網址'); return; }

  const btn = document.querySelector('[onclick="importFromSheet()"]');
  btn.disabled    = true;
  btn.textContent = '匯入中...';
  showMsg(msgEl, '', '');

  try {
    const version = await callAPI({ action: 'getVersion' });
    if (version.status !== 'success' || !['20260428-no-underscore-v3', '20260429-delete-group-v4', '20260429-per-order-hours-v5', '20260430-hour-format-v6'].includes(version.version)) {
      showMsg(msgEl, 'error', 'GAS Web App 尚未更新到新版匯入程式，請先重新部署 Apps Script。');
      return;
    }

    const params = { action: 'importFromSheet', token: session.token, sheetUrl, col, startRow };
    if (sheetName) params.sheetName = sheetName;
    if (shipCol) params.shipCol = shipCol;
    if (contentCol) params.contentCol = contentCol;
    if (estimateCol) params.estimateCol = estimateCol;
    if (remarkCol) params.remarkCol = remarkCol;

    const result = await callAPI(params);
    if (result.status === 'success') {
      showMsg(msgEl, 'success', result.message);
      await loadWorkOrderList();
    } else {
      showMsg(msgEl, 'error', result.message);
    }
  } catch (e) {
    showMsg(msgEl, 'error', `連線失敗：${e.message || e}`);
  } finally {
    btn.disabled    = false;
    btn.textContent = '從 Google Sheet 匯入';
  }
}

// ── Group Management ─────────────────────────────────────────────────────────

async function loadGroups() {
  const wrap = document.getElementById('groupListWrap');
  if (!wrap) return;
  wrap.innerHTML = '<div class="loading">載入中...</div>';
  try {
    const result = await callAPI({ action: 'getGroups' });
    if (result.status !== 'success') {
      wrap.innerHTML = `<div class="empty">${result.message}</div>`;
      return;
    }
    const groups = result.groups || [];
    if (!groups.length) {
      wrap.innerHTML = '<div class="empty">尚無組別，請先匯入</div>';
      return;
    }
    wrap.innerHTML = `
      <div class="table-wrap">
        <table>
          <thead><tr><th>#</th><th>組別</th><th></th></tr></thead>
          <tbody>
            ${groups.map((g, i) => `
              <tr>
                <td style="color:var(--text-muted)">${i + 1}</td>
                <td><span class="badge badge-green">${escHtml(g)}</span></td>
                <td><button class="btn btn-sm" style="color:#dc2626;border:1px solid #fecaca;background:#fef2f2"
                  onclick="deleteGroup('${escAttr(g)}')">刪除</button></td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>`;
  } catch (e) {
    wrap.innerHTML = '<div class="empty">載入失敗</div>';
  }
}

async function deleteGroup(groupName) {
  if (!confirm(`確定要刪除組別「${groupName}」嗎？`)) return;
  try {
    const result = await callAPI({ action: 'deleteGroup', token: session.token, group: groupName });
    if (result.status === 'success') {
      await loadGroups();
      await loadEmployeeFilters();
    } else {
      alert(result.message);
    }
  } catch (e) {
    alert('連線失敗');
  }
}

async function importGroupsFromSheet() {
  const sheetUrl  = extractSpreadsheetId(document.getElementById('importGroupSheetUrl').value);
  const sheetName = document.getElementById('importGroupSheetName').value.trim();
  const col       = document.getElementById('importGroupCol').value;
  const msgEl     = document.getElementById('importGroupMsg');

  if (!sheetUrl) { showMsg(msgEl, 'error', '請輸入 Google Sheet 網址'); return; }

  const btn = document.querySelector('[onclick="importGroupsFromSheet()"]');
  btn.disabled = true;
  btn.textContent = '匯入中...';
  showMsg(msgEl, '', '');

  try {
    const params = { action: 'importGroups', token: session.token, sheetUrl, col };
    if (sheetName) params.sheetName = sheetName;
    const result = await callAPI(params);
    if (result.status === 'success') {
      showMsg(msgEl, 'success', result.message);
      await loadGroups();
      await loadEmployeeFilters();
    } else {
      showMsg(msgEl, 'error', result.message);
    }
  } catch (e) {
    showMsg(msgEl, 'error', `連線失敗：${e.message || e}`);
  } finally {
    btn.disabled = false;
    btn.textContent = '匯入組別';
  }
}

// ── Records ──────────────────────────────────────────────────────────────────

async function loadEmployeeFilters() {
  try {
    const result = await callAPI({ action: 'getAllEmployees', token: session.token });
    if (result.status !== 'success') return;
    allEmployees = result.employees;
    allGroups    = result.groups;

    ['rGroup', 'sGroup'].forEach(id => {
      const sel = document.getElementById(id);
      while (sel.options.length > 1) sel.remove(1);
      allGroups.forEach(g => sel.appendChild(new Option(g, g)));
    });

    const empSel = document.getElementById('rEmpId');
    while (empSel.options.length > 1) empSel.remove(1);
    allEmployees.forEach(emp => {
      empSel.appendChild(new Option(`${emp.id} ${emp.name}`, emp.id));
    });
  } catch (e) {}
}

function refreshWorkOrderFilter() {
  const woSel = document.getElementById('rWorkOrder');
  while (woSel.options.length > 1) woSel.remove(1);
  allWorkOrders.forEach(wo => woSel.appendChild(new Option(wo.id, wo.id)));

  const shipSel = document.getElementById('rShipNo');
  while (shipSel.options.length > 1) shipSel.remove(1);
  const ships = Array.from(new Set(allWorkOrders.map(wo => wo.shipNo).filter(Boolean))).sort();
  ships.forEach(shipNo => shipSel.appendChild(new Option(shipNo, shipNo)));
}

async function loadRecords() {
  const tbody  = document.getElementById('recordsTbody');
  const countEl = document.getElementById('recordsCount');
  const btn    = document.getElementById('rQueryBtn');

  tbody.innerHTML  = '<tr><td colspan="7" class="loading">載入中...</td></tr>';
  countEl.textContent = '';
  btn.disabled = true;

  const params = {
    action:    'getAllReports',
    token:     session.token,
    startDate: document.getElementById('rStartDate').value  || undefined,
    endDate:   document.getElementById('rEndDate').value    || undefined,
    group:     document.getElementById('rGroup').value      || undefined,
    empId:     document.getElementById('rEmpId').value      || undefined,
    workOrder: document.getElementById('rWorkOrder').value  || undefined,
    shipNo:    document.getElementById('rShipNo').value || undefined,
    reportType: document.getElementById('rReportType').value || undefined,
    sortBy:    document.getElementById('rSortBy').value
  };

  try {
    const result = await callAPI(params);
    if (result.status !== 'success') {
      if (result.message.includes('逾時')) return doLogout();
      tbody.innerHTML = `<tr><td colspan="7" class="empty">${result.message}</td></tr>`;
      return;
    }
    const records = result.records;
    countEl.textContent = `共 ${records.length} 筆`;

    if (!records.length) {
      tbody.innerHTML = '<tr><td colspan="7" class="empty">此條件無記錄</td></tr>';
      return;
    }
    tbody.innerHTML = records.map(r => `
      <tr>
        <td>${r.date}</td>
        <td>${escHtml(r.empId)}</td>
        <td><span class="badge badge-green">${escHtml(r.group)}</span></td>
        <td>${escHtml(r.shipNo)}</td>
        <td>${(r.workOrders||[]).map(w => `<span class="badge badge-blue">${escHtml(w)}</span>`).join(' ')}</td>
        <td>${escHtml(r.hours)}</td>
        <td>${escHtml(r.reportType)}</td>
      </tr>`).join('');
  } catch (e) {
    tbody.innerHTML = '<tr><td colspan="7" class="empty">載入失敗</td></tr>';
  } finally {
    btn.disabled = false;
  }
}

// ── Statistics ────────────────────────────────────────────────────────────────

async function loadStats() {
  const params = {
    action:    'getAllReports',
    token:     session.token,
    startDate: document.getElementById('sStartDate').value || undefined,
    endDate:   document.getElementById('sEndDate').value   || undefined,
    group:     document.getElementById('sGroup').value     || undefined,
    reportType: document.getElementById('sReportType').value || undefined
  };

  ['statsByWorkOrder', 'statsByShip', 'statsByGroup', 'statsByEmp', 'statsTimeline'].forEach(id => {
    document.getElementById(id).innerHTML = '<div class="loading">載入中...</div>';
  });

  try {
    const result = await callAPI(params);
    if (result.status !== 'success') {
      ['statsByWorkOrder', 'statsByGroup', 'statsByEmp', 'statsTimeline'].forEach(id => {
        document.getElementById(id).innerHTML = `<div class="empty">${result.message}</div>`;
      });
      return;
    }

    const records = result.records;
    if (!records.length) {
      ['statsByWorkOrder', 'statsByShip', 'statsByGroup', 'statsByEmp', 'statsTimeline'].forEach(id => {
        document.getElementById(id).innerHTML = '<div class="empty">此條件無記錄</div>';
      });
      return;
    }

    renderCountTable('statsByWorkOrder', sumByWorkOrders(records), '工單號碼', 'badge-blue');
    renderCountTable('statsByShip',      sumBy(records, 'shipNo'),    '船號',     '');
    renderCountTable('statsByGroup',     sumBy(records, 'group'),     '組別',     'badge-green');
    renderCountTable('statsByEmp',       sumByEmp(records),           '員工',     '');
    renderTimeline('statsTimeline', records);
  } catch (e) {
    ['statsByWorkOrder', 'statsByShip', 'statsByGroup', 'statsByEmp', 'statsTimeline'].forEach(id => {
      document.getElementById(id).innerHTML = '<div class="empty">載入失敗</div>';
    });
  }
}

function countBy(records, field) {
  const map = {};
  records.forEach(r => { const k = r[field]; map[k] = (map[k] || 0) + 1; });
  return Object.entries(map).sort((a, b) => b[1] - a[1]);
}

function sumBy(records, field) {
  const map = {};
  records.forEach(r => {
    const k = r[field] || '未設定';
    map[k] = (map[k] || 0) + Number(r.hours || 0);
  });
  return Object.entries(map).sort((a, b) => b[1] - a[1]);
}

function sumByWorkOrders(records) {
  const map = {};
  records.forEach(r => {
    (r.workOrders || []).forEach(w => { map[w] = (map[w] || 0) + Number(r.hours || 0); });
  });
  return Object.entries(map).sort((a, b) => b[1] - a[1]);
}

function sumByEmp(records) {
  const map = {};
  records.forEach(r => {
    const k = r.empId;
    if (!map[k]) map[k] = { label: `${r.empId} ${r.name}`, hours: 0 };
    map[k].hours += Number(r.hours || 0);
  });
  return Object.values(map).sort((a, b) => b.hours - a.hours).map(v => [v.label, v.hours]);
}

function renderCountTable(elId, entries, label, badgeClass) {
  const el = document.getElementById(elId);
  if (!entries.length) { el.innerHTML = '<div class="empty">無資料</div>'; return; }
  el.innerHTML = `
    <div class="table-wrap">
      <table>
        <thead><tr><th>${label}</th><th style="text-align:right">總工時</th></tr></thead>
        <tbody>
          ${entries.map(([k, v]) => `
            <tr>
              <td>${badgeClass ? `<span class="badge ${badgeClass}">${escHtml(k)}</span>` : escHtml(k)}</td>
              <td style="text-align:right;font-weight:600">${formatHours(v)}</td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
}

function renderTimeline(elId, records) {
  const el    = document.getElementById(elId);
  const byDay = {};
  records.forEach(r => { byDay[r.date] = (byDay[r.date] || 0) + Number(r.hours || 0); });
  const days  = Object.entries(byDay).sort((a, b) => a[0].localeCompare(b[0]));
  const max   = Math.max(...days.map(d => d[1]));
  const CHART_H = 132; // px — bar area height (160 total minus label space)

  const cols = days.map(([date, count]) => {
    const h   = Math.max(4, Math.round((count / max) * CHART_H));
    const label = date.slice(5); // MM-DD
    return `
      <div class="tl-col">
        <span class="tl-count">${formatHours(count)}</span>
        <div class="tl-bar" style="height:${h}px"></div>
        <span class="tl-date">${label}</span>
      </div>`;
  }).join('');

  el.innerHTML = `<div class="timeline-wrap"><div class="timeline-chart">${cols}</div></div>`;
  enableDragScroll(el.querySelector('.timeline-wrap'));
}

function enableDragScroll(el) {
  let startX, scrollLeft, dragging = false;
  el.addEventListener('pointerdown', e => {
    dragging  = true;
    startX    = e.pageX - el.offsetLeft;
    scrollLeft = el.scrollLeft;
    el.setPointerCapture(e.pointerId);
  });
  el.addEventListener('pointermove', e => {
    if (!dragging) return;
    const dx = e.pageX - el.offsetLeft - startX;
    el.scrollLeft = scrollLeft - dx;
  });
  el.addEventListener('pointerup',     () => { dragging = false; });
  el.addEventListener('pointercancel', () => { dragging = false; });
}

// ── Utilities ─────────────────────────────────────────────────────────────────

function showMsg(el, type, text) {
  el.className = type === 'success' ? 'msg msg-success' : type === 'error' ? 'msg msg-error' : 'msg hidden';
  el.textContent = text;
}

function formatHours(value) {
  const num = Number(value || 0);
  return Number.isInteger(num) ? String(num) : num.toFixed(1).replace(/\.0$/, '');
}

function extractSpreadsheetId(input) {
  const text = String(input || '').replace(/\s+/g, '').trim();
  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : text;
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escAttr(str) {
  return String(str).replace(/'/g, '&#39;').replace(/"/g, '&quot;');
}

window.addEventListener('DOMContentLoaded', init);
