(function () {
  'use strict';
  const table = document.getElementById('companiesTable');
  if (!table) return;
  const rows = Array.from(table.querySelectorAll('[data-company-row]'));
  const checks = Array.from(table.querySelectorAll('.company-select'));
  const all = document.getElementById('companySelectAll'), search = document.getElementById('companyListSearch');
  const review = document.getElementById('companyReviewSelection');
  const normalize = value => String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[.\/-]/g, '').trim();
  function update() {
    const selected = checks.filter(check => check.checked), visible = checks.filter(check => !check.closest('tr').hidden);
    const hiddenSelected = selected.filter(check => check.closest('tr').hidden).length;
    all.checked = visible.length > 0 && visible.every(check => check.checked);
    all.indeterminate = visible.some(check => check.checked) && !all.checked;
    all.disabled = !visible.length;
    review.disabled = !selected.length;
    review.textContent = selected.length ? `Revisar exclusão de ${selected.length} selecionada(s)` : 'Revisar exclusão das selecionadas';
    document.getElementById('companySelectionStatus').textContent = selected.length ? `${selected.length} empresa(s) selecionada(s)${hiddenSelected ? ` · ${hiddenSelected} fora da busca atual` : ''}.` : 'Nenhuma empresa selecionada.';
    document.getElementById('companyFilterStatus').textContent = `${visible.length} de ${rows.length} empresa(s) visível(is). Selecionar todas considera somente as linhas visíveis; a seleção fora da busca é mantida.`;
  }
  search.addEventListener('input', () => {
    const query = normalize(search.value);
    rows.forEach(row => row.hidden = !normalize(row.textContent).includes(query));
    update();
  });
  all.addEventListener('change', () => { checks.filter(check => !check.closest('tr').hidden).forEach(check => check.checked = all.checked); update(); });
  checks.forEach(check => check.addEventListener('change', update));
  document.getElementById('companyClearSelection').addEventListener('click', () => { checks.forEach(check => check.checked = false); update(); });
  document.getElementById('companySelectionForm').addEventListener('submit', event => {
    if (!checks.some(check => check.checked)) { event.preventDefault(); document.getElementById('companySelectionStatus').textContent = 'Selecione ao menos uma empresa para revisar a exclusão.'; }
  });
  update();
})();
