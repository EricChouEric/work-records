// ── 部署後請將此 URL 改為您的 Google Apps Script Web App URL ──────────────────
const API_URL = 'YOUR_GOOGLE_APPS_SCRIPT_URL';
// ─────────────────────────────────────────────────────────────────────────────

async function hashPassword(password) {
  const buf = new TextEncoder().encode(password);
  const hash = await crypto.subtle.digest('SHA-256', buf);
  return Array.from(new Uint8Array(hash))
    .map(b => b.toString(16).padStart(2, '0'))
    .join('');
}

async function callAPI(params) {
  if (API_URL === 'YOUR_GOOGLE_APPS_SCRIPT_URL') {
    throw new Error('請先在 js/api.js 設定 API_URL（Google Apps Script Web App URL）');
  }
  const url = new URL(API_URL);
  Object.entries(params).forEach(([k, v]) => {
    if (v !== undefined && v !== null && v !== '') url.searchParams.set(k, v);
  });
  const resp = await fetch(url.toString(), { redirect: 'follow' });
  if (!resp.ok) throw new Error('API 請求失敗 (' + resp.status + ')');
  return resp.json();
}
