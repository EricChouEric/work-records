function getSession() {
  const token = localStorage.getItem('token');
  if (!token) return null;
  return {
    token,
    role:  localStorage.getItem('role'),
    name:  localStorage.getItem('name'),
    empId: localStorage.getItem('empId'),
    group: localStorage.getItem('group') || ''
  };
}

function requireAuth(role) {
  const session = getSession();
  if (!session) { window.location.href = 'index.html'; return null; }
  if (role && session.role !== role) {
    window.location.href = session.role === 'manager' ? 'manager.html' : 'employee.html';
    return null;
  }
  return session;
}

async function doLogout() {
  const token = localStorage.getItem('token');
  if (token) {
    try { await callAPI({ action: 'logout', token }); } catch (e) {}
  }
  localStorage.clear();
  window.location.href = 'index.html';
}
