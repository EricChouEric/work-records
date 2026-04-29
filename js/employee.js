let session;
let workOrdersByShip = {}; // { shipNo: [{id, remark}] }

function normalizeNumericInputValue(value) {
  return String(value || '').trim().replace(/^_(?=\d+$)/, '');
}

async function init() {
  session = requireAuth('employee');
  if (!session) return;

  document.getElementById('userGroup').textContent  = session.group || '';
  document.getElementById('infoEmpId').textContent  = session.empId;
  document.getElementById('infoGroup').textContent  = session.group || '（未設定）';

  const now = new Date();
  document.getElementById('date').valueAsDate = now;

  const yearInput  = document.getElementById('filterYear');
  const monthInput = document.getElementById('filterMonth');
  yearInput.value  = String(now.getFullYear());
  monthInput.value = String(now.getMonth() + 1).padStart(2, '0');

  setupNumericInputNormalizers();
  setupTabs();
  await loadWorkOrders();

  document.getElementById('reportForm').addEventListener('submit', handleSubmit);
  document.getElementById('filterApply').addEventListener('click', loadRecords);

  await loadRecords();
}

function setupNumericInputNormalizers() {
  document.addEventListener('blur', e => {
    if (e.target.matches('#shipManual, .manual-wo-input, .extra-wo-input')) {
      e.target.value = normalizeNumericInputValue(e.target.value);
    }
  }, true);
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

async function loadWorkOrders() {
  const shipSel = document.getElementById('shipSelect');
  const msgEl = document.getElementById('reportMsg');
  try {
    const result = await callAPI({ action: 'getWorkOrders', token: session.token });

    if (result.status !== 'success') {
      if (result.message && result.message.includes('逾時')) return doLogout();
      shipSel.innerHTML = '<option value="">工單載入失敗</option>';
      msgEl.className = 'msg msg-error';
      msgEl.textContent = result.message || '工單載入失敗';
      return;
    }

    workOrdersByShip = {};
    result.workOrders.forEach(wo => {
      const key = wo.shipNo || '（未分配船號）';
      if (!workOrdersByShip[key]) workOrdersByShip[key] = [];
      workOrdersByShip[key].push(wo);
    });

    const ships = Object.keys(workOrdersByShip).sort();
    shipSel.innerHTML = '<option value="">-- 請選擇船號 --</option>';
    ships.forEach(s => shipSel.appendChild(new Option(s, s)));
    shipSel.appendChild(new Option('▶ 手動輸入臨時船號...', '__manual__'));

    shipSel.addEventListener('change', onShipChange);
  } catch (e) {
    shipSel.innerHTML = '<option value="">載入失敗</option>';
    shipSel.appendChild(new Option('▶ 手動輸入臨時船號...', '__manual__'));
    shipSel.addEventListener('change', onShipChange);
    msgEl.className = 'msg msg-error';
    msgEl.textContent = `工單載入失敗：${e.message || e}`;
  }
}

function onShipChange() {
  const val = document.getElementById('shipSelect').value;
  const shipManualGroup   = document.getElementById('shipManualGroup');
  const workOrderSection  = document.getElementById('workOrderSection');
  const woManualSection   = document.getElementById('workOrderManualSection');

  if (val === '__manual__') {
    shipManualGroup.style.display  = '';
    workOrderSection.style.display = 'none';
    woManualSection.style.display  = '';
  } else if (val) {
    shipManualGroup.style.display  = 'none';
    workOrderSection.style.display = '';
    woManualSection.style.display  = 'none';
    renderCheckboxes(workOrdersByShip[val] || []);
  } else {
    shipManualGroup.style.display  = 'none';
    workOrderSection.style.display = 'none';
    woManualSection.style.display  = 'none';
  }
}

function renderCheckboxes(orders) {
  const wrap = document.getElementById('workOrderCheckboxes');
  if (!orders.length) {
    wrap.innerHTML = '<div class="empty" style="padding:.75rem 0">此船號尚無工單</div>';
    return;
  }
  wrap.innerHTML = orders.map(wo => {
    const details = [wo.content, wo.remark]
      .filter(Boolean)
      .join(' / ');
    const label = details ? `${escHtml(wo.id)}（${escHtml(details)}）` : escHtml(wo.id);
    return `
      <label class="check-item">
        <input type="checkbox" class="wo-check" value="${escAttr(wo.id)}">
        <span>${label}</span>
      </label>`;
  }).join('');
}

function getSubmitValues() {
  const shipSel = document.getElementById('shipSelect');

  if (shipSel.value === '__manual__') {
    const shipInput = document.getElementById('shipManual');
    const shipNo = normalizeNumericInputValue(shipInput.value);
    shipInput.value = shipNo;
    const orders = [...document.querySelectorAll('.manual-wo-input')]
      .map(el => {
        el.value = normalizeNumericInputValue(el.value);
        return el.value;
      }).filter(Boolean);
    return { shipNo, workOrders: orders.join(',') };
  }

  const shipNo  = shipSel.value;
  const checked = [...document.querySelectorAll('.wo-check:checked')].map(c => c.value);
  const extras  = [...document.querySelectorAll('.extra-wo-input')]
    .map(el => {
      el.value = normalizeNumericInputValue(el.value);
      return el.value;
    }).filter(Boolean);
  return { shipNo, workOrders: [...checked, ...extras].join(',') };
}

function addManualWoField() {
  const list = document.getElementById('manualWoList');
  const row  = document.createElement('div');
  row.className = 'manual-wo-row';
  row.innerHTML = `<input type="text" class="manual-wo-input" placeholder="請輸入工單號碼">
    <button type="button" class="btn-remove" onclick="this.parentElement.remove()">✕</button>`;
  list.appendChild(row);
}

function addExtraWoField() {
  const list = document.getElementById('extraWoList');
  const row  = document.createElement('div');
  row.className = 'manual-wo-row';
  row.innerHTML = `<input type="text" class="extra-wo-input" placeholder="臨時工單號碼">
    <button type="button" class="btn-remove" onclick="this.parentElement.remove()">✕</button>`;
  list.appendChild(row);
}

function resetManualLists() {
  document.getElementById('manualWoList').innerHTML =
    '<div class="manual-wo-row"><input type="text" class="manual-wo-input" placeholder="請輸入工單號碼"></div>';
  document.getElementById('extraWoList').innerHTML = '';
}

async function handleSubmit(e) {
  e.preventDefault();
  const btn   = document.getElementById('submitBtn');
  const msgEl = document.getElementById('reportMsg');

  const { shipNo, workOrders } = getSubmitValues();
  const hours = Number(document.getElementById('hours').value);
  const reportType = document.getElementById('reportType').value;

  if (!shipNo) {
    msgEl.className   = 'msg msg-error';
    msgEl.textContent = '請選擇或輸入船號';
    return;
  }
  if (!workOrders) {
    msgEl.className   = 'msg msg-error';
    msgEl.textContent = '請至少勾選一個工單號碼';
    return;
  }
  if (!hours || hours <= 0) {
    msgEl.className   = 'msg msg-error';
    msgEl.textContent = '請填寫實際工時';
    return;
  }

  btn.disabled    = true;
  btn.textContent = '送出中...';
  msgEl.className = 'msg hidden';

  try {
    const result = await callAPI({
      action:     'submitReport',
      token:      session.token,
      date:       document.getElementById('date').value,
      shipNo,
      workOrders,
      hours,
      reportType
    });

    if (result.status === 'success') {
      msgEl.className   = 'msg msg-success';
      msgEl.textContent = result.message || '報工已儲存';
      e.target.reset();
      document.getElementById('date').valueAsDate = new Date();
      document.getElementById('shipManualGroup').style.display       = 'none';
      document.getElementById('workOrderSection').style.display      = 'none';
      document.getElementById('workOrderManualSection').style.display = 'none';
      resetManualLists();
      await loadRecords();
    } else {
      if (result.message.includes('逾時')) return doLogout();
      msgEl.className   = 'msg msg-error';
      msgEl.textContent = result.message;
    }
  } catch (err) {
    msgEl.className   = 'msg msg-error';
    msgEl.textContent = err.message || '連線失敗，請稍後再試';
  } finally {
    btn.disabled    = false;
    btn.textContent = '送出報工';
  }
}

async function loadRecords() {
  const year  = document.getElementById('filterYear').value;
  const month = document.getElementById('filterMonth').value;
  const tbody = document.getElementById('recordsTbody');

  tbody.innerHTML = '<tr><td colspan="6" class="loading">載入中...</td></tr>';
  try {
    const result = await callAPI({ action: 'getMyReports', token: session.token, year, month });
    if (result.status !== 'success') {
      if (result.message.includes('逾時')) return doLogout();
      tbody.innerHTML = `<tr><td colspan="6" class="empty">${result.message}</td></tr>`;
      return;
    }
    if (!result.records.length) {
      tbody.innerHTML = '<tr><td colspan="6" class="empty">此期間無記錄</td></tr>';
      return;
    }
    tbody.innerHTML = result.records.map(r => `
      <tr>
        <td>${r.date}</td>
        <td>${escHtml(r.shipNo)}</td>
        <td>${r.workOrders.map(w => `<span class="badge badge-blue">${escHtml(w)}</span>`).join(' ')}</td>
        <td>${escHtml(r.hours)}</td>
        <td>${escHtml(r.reportType)}</td>
        <td>${escHtml(r.group)}</td>
      </tr>`).join('');
  } catch (e) {
    tbody.innerHTML = '<tr><td colspan="6" class="empty">載入失敗</td></tr>';
  }
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
