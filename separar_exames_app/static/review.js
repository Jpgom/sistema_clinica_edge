(() => {
  const jobId = window.EDGE_JOB_ID;
  let items = [], employees = [], selected = null, busy = false;
  const saveBtn = document.getElementById('saveReview');
  const ignoreBtn = document.getElementById('ignoreReview');

  async function load() {
    try {
      const data = await edgeJson(`/api/jobs/${jobId}/review`);
      items = data.items;
      employees = data.employees;
      if (!items.length) {
        selected = null;
        document.getElementById('reviewWorkspace').classList.add('hidden');
        document.getElementById('reviewEmpty').classList.remove('hidden');
        return;
      }
      document.getElementById('reviewEmpty').classList.add('hidden');
      document.getElementById('reviewWorkspace').classList.remove('hidden');
      document.getElementById('reviewItems').innerHTML = items.map((x, i) =>
        `<button class="review-item ${i === 0 ? 'active' : ''}" data-id="${x.id}"><span>${x.page_number}</span><div><strong>${esc(x.employee_name || 'Não identificado')}</strong><small>${esc(x.exam_type || 'Tipo incerto')} • ${esc(x.source_pdf)}</small></div></button>`
      ).join('');
      document.querySelectorAll('.review-item').forEach(b => b.onclick = () => select(b.dataset.id));
      select(items[0].id);
    } catch (e) { edgeToast(e.message, 'error'); }
  }

  function select(id) {
    selected = items.find(x => x.id === id);
    if (!selected) return;
    document.querySelectorAll('.review-item').forEach(b => b.classList.toggle('active', b.dataset.id === id));
    document.getElementById('reviewImage').src = edgeUrl(`/api/jobs/${jobId}/page/${id}.png?t=${Date.now()}`);
    document.getElementById('reviewSource').textContent = `${selected.source_pdf} • página ${selected.page_number}`;
    document.getElementById('reviewEmployee').textContent = selected.employee_name || 'Funcionário não confirmado';
    document.getElementById('reviewReason').textContent = selected.reason || '';
    const sel = document.getElementById('reviewEmployeeSelect');
    sel.innerHTML = employees.map(e => `<option value="${e.row_id}">${esc(e.name)} — ${esc(e.company || '')}</option>`).join('');
    if (selected.employee_row_id) sel.value = selected.employee_row_id;
    const exam = document.getElementById('reviewExam');
    if ([...exam.options].some(o => o.value === selected.exam_type)) exam.value = selected.exam_type;
  }

  function setBusy(value) {
    busy = value;
    saveBtn.disabled = value;
    ignoreBtn.disabled = value;
    saveBtn.textContent = value ? 'Salvando...' : 'Confirmar e salvar';
  }

  async function resolve(action) {
    if (!selected || busy) return;
    const payload = {
      action,
      employee_row_id: +document.getElementById('reviewEmployeeSelect').value,
      exam_type: document.getElementById('reviewExam').value
    };
    setBusy(true);
    try {
      const data = await edgeJson(`/api/jobs/${jobId}/review/${selected.id}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      edgeToast(data.already_resolved ? 'Esta página já havia sido revisada.' : (action === 'IGNORAR' ? 'Página ignorada.' : 'Página salva e adicionada à extração.'));
      await load();
    } catch (e) {
      edgeToast(e.message, 'error');
    } finally {
      setBusy(false);
    }
  }

  saveBtn.onclick = () => resolve('SALVAR');
  ignoreBtn.onclick = () => resolve('IGNORAR');
  const esc = s => String(s || '').replace(/[&<>'"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c]));
  load();
})();
