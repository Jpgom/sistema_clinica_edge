(function () {
  'use strict';
  const table = document.getElementById('paymentsTable');
  if (!table) return;
  const rows = Array.from(table.querySelectorAll('tbody tr[data-search-row]'));
  const checks = Array.from(table.querySelectorAll('.rowcheck'));
  const selectAll = document.getElementById('checkAll');
  const search = document.getElementById('paymentSearch'), state = document.getElementById('paymentState');
  const normalize = value => String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  function updateSelection() {
    const visible = checks.filter(check => !check.closest('tr').hidden);
    const selected = checks.filter(check => check.checked);
    const hiddenSelected = selected.filter(check => check.closest('tr').hidden).length;
    selectAll.checked = visible.length > 0 && visible.every(check => check.checked);
    selectAll.indeterminate = visible.some(check => check.checked) && !selectAll.checked;
    selectAll.disabled = !visible.length;
    document.getElementById('selectionStatus').textContent = selected.length
      ? `${selected.length} empresa(s) selecionada(s)${hiddenSelected ? ` · ${hiddenSelected} fora do filtro atual` : ''}.`
      : 'Nenhuma empresa selecionada.';
    document.querySelectorAll('[data-selection-action]').forEach(button => button.disabled = selected.length === 0);
  }
  function applyFilters() {
    const query = normalize(search.value).trim(), filter = state.value;
    rows.forEach(row => {
      const code = row.dataset.paymentState;
      const stateMatches = !filter || (filter === 'OTHER' ? !['PAYMENT', 'COMPLETE'].includes(code) : code === filter);
      row.hidden = !stateMatches || !normalize(row.textContent).includes(query);
    });
    document.getElementById('paymentFilterStatus').textContent = `${rows.filter(row => !row.hidden).length} de ${rows.length} empresa(s) exibida(s). A seleção fora do filtro é mantida até você limpá-la.`;
    updateSelection();
  }
  selectAll.addEventListener('change', () => { checks.filter(check => !check.closest('tr').hidden).forEach(check => check.checked = selectAll.checked); updateSelection(); });
  checks.forEach(check => check.addEventListener('change', updateSelection));
  document.getElementById('clearPaymentSelection').addEventListener('click', () => { checks.forEach(check => check.checked = false); updateSelection(); });
  search.addEventListener('input', applyFilters); state.addEventListener('change', applyFilters);
  document.getElementById('paymentsForm').addEventListener('submit', event => {
    const selected = checks.filter(check => check.checked).length;
    if (selected && !confirm(`Aplicar "${event.submitter?.textContent.trim() || 'alteração de pagamento'}" a ${selected} empresa(s) selecionada(s)?`)) event.preventDefault();
  });
  applyFilters();
})();
