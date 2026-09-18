window.edgeToast = function(message, kind='ok') {
  const el = document.getElementById('toast');
  if (!el) return;
  el.textContent = message;
  el.className = `toast show ${kind}`;
  clearTimeout(window.__edgeToastTimer);
  window.__edgeToastTimer = setTimeout(() => el.className = 'toast', 3200);
};

window.edgeJson = async function(url, options={}) {
  const r = await fetch(url, options);
  let data = {};
  try { data = await r.json(); } catch (_) {}
  if (!r.ok) throw new Error(data.error || `Erro ${r.status}`);
  return data;
};
