(() => {
  const q = document.getElementById('archiveQuery');
  const period = document.getElementById('periodFilter');
  const company = document.getElementById('companyFilter');
  const exam = document.getElementById('examFilter');
  const receipt = document.getElementById('receiptFilter');
  const rows = document.getElementById('archiveRows');
  const total = document.getElementById('archiveTotal');
  const companies = document.getElementById('archiveCompanies');
  const employees = document.getElementById('archiveEmployees');
  const hint = document.getElementById('archiveTableHint');
  const prev = document.getElementById('prevPage');
  const next = document.getElementById('nextPage');
  const pageInfo = document.getElementById('pageInfo');
  const selectAll = document.getElementById('selectAllDocs');
  const downloadSelected = document.getElementById('downloadSelected');
  const downloadFiltered = document.getElementById('downloadFiltered');
  const deleteSelected = document.getElementById('deleteSelected');
  const deleteFiltered = document.getElementById('deleteFiltered');
  let currentPage = 1;
  let pages = 1;
  let debounce = null;

  function esc(s='') { return String(s).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }
  function parsePeriod() {
    if (!period.value) return {};
    const [year, month] = period.value.split('-').map(Number);
    return {year, month};
  }
  function params(includePage=true) {
    const p = new URLSearchParams();
    if (q.value.trim()) p.set('q', q.value.trim());
    const pm = parsePeriod();
    if (pm.year) p.set('year', pm.year);
    if (pm.month) p.set('month', pm.month);
    if (company.value) p.set('company', company.value);
    if (exam.value) p.set('exam_type', exam.value);
    if (receipt.value) p.set('receipt_filter', receipt.value);
    if (includePage) { p.set('page', currentPage); p.set('page_size', 100); }
    return p;
  }
  function filterPayload() {
    const p = params(false);
    const out = {use_filters: true};
    for (const [k, v] of p.entries()) out[k] = v;
    return out;
  }
  async function loadMeta() {
    const pm = parsePeriod();
    const qp = new URLSearchParams();
    if (pm.year) qp.set('year', pm.year);
    if (pm.month) qp.set('month', pm.month);
    const data = await edgeJson('/api/archive/meta?' + qp.toString());
    const oldCompany = company.value;
    const oldExam = exam.value;
    period.innerHTML = '<option value="">Todos os períodos</option>' + data.periods.map(x => `<option value="${x.year}-${String(x.month).padStart(2,'0')}">${String(x.month).padStart(2,'0')}/${x.year} · ${x.count} arquivo(s)</option>`).join('');
    if (pm.year) period.value = `${pm.year}-${String(pm.month).padStart(2,'0')}`;
    company.innerHTML = '<option value="">Todas as empresas</option>' + data.companies.map(x => `<option value="${esc(x.key)}">${esc(x.name)}${x.document ? ` · ${esc(x.document_kind)} ${esc(x.document)}` : ''} · ${x.count}</option>`).join('');
    if ([...company.options].some(o => o.value === oldCompany)) company.value = oldCompany;
    const quick = new Set(['ASO','AUDIOMETRIA','ESPIROMETRIA','ACUIDADE VISUAL']);
    const types = [...new Set([...data.exam_types, ...quick])].sort();
    exam.innerHTML = '<option value="">Todos os exames</option>' + types.map(x => `<option value="${esc(x)}">${esc(x)}</option>`).join('');
    if ([...exam.options].some(o => o.value === oldExam)) exam.value = oldExam;
  }
  async function load() {
    rows.innerHTML = '<tr><td colspan="7" class="table-empty">Carregando arquivo...</td></tr>';
    try {
      const data = await edgeJson('/api/archive?' + params().toString());
      total.textContent = data.total;
      companies.textContent = data.companies;
      employees.textContent = data.employees;
      pages = data.pages || 1;
      currentPage = Math.min(currentPage, pages);
      pageInfo.textContent = `Página ${currentPage} de ${pages}`;
      prev.disabled = currentPage <= 1;
      next.disabled = currentPage >= pages;
      hint.textContent = data.total ? `${data.total} documento(s) correspondente(s) aos filtros.` : 'Nenhum documento encontrado.';
      rows.innerHTML = data.items.length ? data.items.map(x => `
        <tr>
          <td class="check-col"><input class="doc-check" type="checkbox" value="${x.id}"></td>
          <td><strong>${esc(x.employee_name || '—')}</strong><small>${x.employee_cpf ? `CPF ${esc(x.employee_cpf)}` : ''}</small></td>
          <td><span class="receipt-pill ${String(x.receipt||'').toUpperCase()==='A PRAZO' ? 'on-credit' : ''}">${esc(x.receipt || '—')}</span></td>
          <td><strong>${esc(x.company_name || 'Empresa não identificada')}</strong><small>${x.company_document ? `${esc(x.company_document_kind)} ${esc(x.company_document)}` : 'CPF/CNPJ não informado'}</small></td>
          <td><span class="exam-pill">${esc(x.exam_type || 'OUTROS')}</span>${x.exam_subtype ? `<small>${esc(x.exam_subtype)}</small>` : ''}</td>
          <td><strong>${esc(x.competency)}</strong><small>${esc(x.original_filename)}</small></td>
          <td><div class="archive-row-actions"><a class="btn tiny ghost" target="_blank" href="${edgeUrl(`/arquivo/documento/${x.id}/visualizar`)}">Visualizar</a><a class="btn tiny primary" href="${edgeUrl(`/arquivo/documento/${x.id}/baixar`)}">Baixar</a><button class="btn tiny danger ghost doc-delete" type="button" data-id="${x.id}" data-name="${esc(x.employee_name || x.original_filename || 'documento')}">Apagar</button></div></td>
        </tr>`).join('') : '<tr><td colspan="7" class="table-empty">Nenhum exame arquivado com estes filtros.</td></tr>';
      selectAll.checked = false;
      updateSelected();
      document.querySelectorAll('.doc-check').forEach(c => c.addEventListener('change', updateSelected));
      document.querySelectorAll('.doc-delete').forEach(b => b.addEventListener('click', () => deleteOne(b.dataset.id, b.dataset.name)));
    } catch (e) {
      rows.innerHTML = `<tr><td colspan="7" class="table-empty">${esc(e.message)}</td></tr>`;
      edgeToast(e.message, 'error');
    }
  }
  function updateSelected() {
    const n = document.querySelectorAll('.doc-check:checked').length;
    downloadSelected.disabled = n === 0;
    deleteSelected.disabled = n === 0;
    downloadSelected.textContent = n ? `Baixar selecionados (${n})` : 'Baixar selecionados';
    deleteSelected.textContent = n ? `Apagar selecionados (${n})` : 'Apagar selecionados';
  }
  function selectedIds() { return [...document.querySelectorAll('.doc-check:checked')].map(x => Number(x.value)).filter(Boolean); }
  function downloadUrl(ids=[]) {
    const p = params(false);
    if (ids.length) p.set('ids', ids.join(','));
    return edgeUrl('/arquivo/baixar.zip?' + p.toString());
  }
  async function refreshAfterDelete(count) {
    edgeToast(`${count} documento(s) apagado(s) do arquivo.`);
    await loadMeta();
    await load();
  }
  async function deleteOne(id, name) {
    if (!confirm(`Apagar definitivamente o exame arquivado de ${name}?\n\nO PDF também será removido do armazenamento permanente.`)) return;
    try {
      const data = await edgeJson(`/api/archive/${id}`, {method: 'DELETE'});
      await refreshAfterDelete(data.deleted_count || 0);
    } catch (e) { edgeToast(e.message, 'error'); }
  }
  async function deleteIds(ids) {
    if (!ids.length) return;
    if (!confirm(`Apagar definitivamente ${ids.length} documento(s) selecionado(s)?\n\nOs PDFs também serão removidos do armazenamento permanente.`)) return;
    try {
      const data = await edgeJson('/api/archive', {
        method: 'DELETE', headers: {'Content-Type':'application/json'}, body: JSON.stringify({ids})
      });
      await refreshAfterDelete(data.deleted_count || 0);
    } catch (e) { edgeToast(e.message, 'error'); }
  }
  async function deleteCurrentFilters() {
    const n = Number(total.textContent || 0);
    if (!n) return edgeToast('Nenhum documento para apagar.', 'error');
    const hasAnyFilter = q.value.trim() || period.value || company.value || exam.value || receipt.value;
    const warning = hasAnyFilter
      ? `Apagar definitivamente os ${n} documento(s) que correspondem aos filtros atuais?`
      : `ATENÇÃO: nenhum filtro está aplicado. Isso apagará TODOS os ${n} documentos arquivados. Continuar?`;
    if (!confirm(`${warning}\n\nOs PDFs também serão removidos do armazenamento permanente.`)) return;
    if (!hasAnyFilter && !confirm('Confirma novamente a exclusão de TODO O ARQUIVO DE EXAMES?')) return;
    try {
      const data = await edgeJson('/api/archive', {
        method: 'DELETE', headers: {'Content-Type':'application/json'}, body: JSON.stringify(filterPayload())
      });
      currentPage = 1;
      await refreshAfterDelete(data.deleted_count || 0);
    } catch (e) { edgeToast(e.message, 'error'); }
  }

  period.addEventListener('change', async () => { currentPage=1; await loadMeta(); await load(); });
  company.addEventListener('change', () => { currentPage=1; load(); });
  exam.addEventListener('change', () => { currentPage=1; syncQuick(); load(); });
  receipt.addEventListener('change', () => { currentPage=1; load(); });
  q.addEventListener('input', () => { clearTimeout(debounce); debounce=setTimeout(() => {currentPage=1; load();}, 260); });
  document.getElementById('clearFilters').addEventListener('click', async () => { q.value=''; period.value=''; company.value=''; exam.value=''; receipt.value=''; currentPage=1; await loadMeta(); syncQuick(); await load(); });
  document.querySelectorAll('.quick-chip').forEach(b => b.addEventListener('click', () => { exam.value=b.dataset.exam || ''; currentPage=1; syncQuick(); load(); }));
  function syncQuick() { document.querySelectorAll('.quick-chip').forEach(b => b.classList.toggle('active', (b.dataset.exam||'') === exam.value)); }
  selectAll.addEventListener('change', () => { document.querySelectorAll('.doc-check').forEach(c => c.checked=selectAll.checked); updateSelected(); });
  downloadSelected.addEventListener('click', () => { const ids=selectedIds(); if(ids.length) window.location.href=downloadUrl(ids); });
  downloadFiltered.addEventListener('click', () => { if (+total.textContent === 0) return edgeToast('Nenhum documento para baixar.', 'error'); window.location.href=downloadUrl(); });
  deleteSelected.addEventListener('click', () => deleteIds(selectedIds()));
  deleteFiltered.addEventListener('click', deleteCurrentFilters);
  prev.addEventListener('click', () => { if(currentPage>1){currentPage--; load();} });
  next.addEventListener('click', () => { if(currentPage<pages){currentPage++; load();} });

  (async () => { try { await loadMeta(); syncQuick(); await load(); } catch(e){ edgeToast(e.message,'error'); } })();
})();
