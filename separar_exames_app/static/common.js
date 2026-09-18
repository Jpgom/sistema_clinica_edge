window.edgeBasePath = function() {
  if (window.EDGE_BASE_PATH !== undefined) return window.EDGE_BASE_PATH || '';
  const marker = '/separar-exames';
  if (window.location.pathname === marker || window.location.pathname.startsWith(marker + '/')) return marker;
  return '';
};

window.edgeUrl = function(url) {
  if (!url) return url;
  const value = String(url);
  if (/^(https?:|blob:|data:|mailto:|tel:)/i.test(value)) return value;
  const base = window.edgeBasePath();
  if (!base) return value;
  if (value === '/') return base + '/';
  if (value.startsWith(base + '/') || value === base) return value;
  if (value.startsWith('/')) return base + value;
  return value;
};

window.edgeToast = function(message, kind='ok') {
  const el = document.getElementById('toast');
  if (!el) return;
  el.textContent = message;
  el.className = `toast show ${kind}`;
  clearTimeout(window.__edgeToastTimer);
  window.__edgeToastTimer = setTimeout(() => el.className = 'toast', 3200);
};

window.edgeJson = async function(url, options={}) {
  const r = await fetch(window.edgeUrl(url), options);
  let data = {};
  try { data = await r.json(); } catch (_) {}
  if (!r.ok) throw new Error(data.error || `Erro ${r.status}`);
  return data;
};
