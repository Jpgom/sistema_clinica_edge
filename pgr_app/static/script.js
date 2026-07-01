
const PGR_BASE_URL = (window.PGR_BASE_URL || '').replace(/\/$/, '');
function pgrUrl(path) {
  const cleanPath = String(path || '');
  return `${PGR_BASE_URL}${cleanPath.startsWith('/') ? cleanPath : `/${cleanPath}`}`;
}
function isVisible(el) {
  return !el || el.offsetParent !== null || el.style.display !== 'none';
}

function toggleChecks(name, checked) {
  document.querySelectorAll(`input[name="${name}"]`).forEach((box) => {
    const item = box.closest('.compact-row,.list-check,.sector-option,.sector-picker,.risk-card');
    if (!item || item.style.display !== 'none') box.checked = checked;
  });
}
function toggleAll(checked) { toggleChecks('risk_ids', checked); }

function attachSearch(inputId, selector) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener('input', () => {
    const term = input.value.trim().toLowerCase();
    document.querySelectorAll(selector).forEach((item) => {
      const text = item.dataset.search || item.textContent.toLowerCase();
      item.style.display = text.includes(term) ? '' : 'none';
    });
  });
}

attachSearch('riskSearch', '#riskList .compact-row');
attachSearch('riskGroupSearch', '#riskGroupList .risk-group-row');
attachSearch('riskGroupFormSearch', '#riskGroupFormList .risk-group-risk-option');
attachSearch('companySearch', '#companyList .compact-row');
attachSearch('sectorGroupSearch', '#sectorGroupList .group-box');
attachSearch('pgrSectorSearch', '#pgrSectorList .sector-option');
attachSearch('riskPickerSearch', '.risk-option');
attachSearch('examPickerSearch', '.exam-option');

// Máscaras automáticas para datas.
// Campos de vigência da empresa aceitam os dois formatos:
// - MM/AAAA (ex.: 06/2027)
// - DD/MM/AAAA (ex.: 01/06/2027)
function maskFullDateValue(raw) {
  const digits = String(raw || '').replace(/\D/g, '').slice(0, 8);
  if (digits.length <= 2) return digits;
  if (digits.length <= 4) return `${digits.slice(0, 2)}/${digits.slice(2)}`;
  return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
}

function maskFlexibleCompanyDateValue(raw) {
  const digits = String(raw || '').replace(/\D/g, '').slice(0, 8);
  if (digits.length <= 2) return digits;

  // Até 6 números, tratar como MM/AAAA. Assim 062027 vira 06/2027.
  if (digits.length <= 6) return `${digits.slice(0, 2)}/${digits.slice(2)}`;

  // Com 7 ou 8 números, tratar como DD/MM/AAAA. Assim 01062027 vira 01/06/2027.
  return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
}

function attachDateMasks() {
  const fullDateSelectors = [
    '#dataCriacaoLaudo',
    '#dataDaRevisao',
    'input[name="data_avaliacao"]',
    'input[name="data_criacao_laudo"]',
    'input[name="data_da_revisao"]'
  ];
  const flexibleCompanyDateSelectors = [
    'input[name="data_atual"]',
    'input[name="data_final"]'
  ];

  document.querySelectorAll(fullDateSelectors.join(',')).forEach((input) => {
    input.setAttribute('inputmode', 'numeric');
    input.setAttribute('maxlength', '10');
    input.addEventListener('input', () => {
      const before = input.value;
      const masked = maskFullDateValue(before);
      if (before !== masked) input.value = masked;
    });
  });

  document.querySelectorAll(flexibleCompanyDateSelectors.join(',')).forEach((input) => {
    input.setAttribute('inputmode', 'numeric');
    input.setAttribute('maxlength', '10');
    input.setAttribute('placeholder', input.getAttribute('placeholder') || 'Ex.: 06/2027 ou 01/06/2027');
    input.addEventListener('input', () => {
      const before = input.value;
      const masked = maskFlexibleCompanyDateValue(before);
      if (before !== masked) input.value = masked;
    });
  });
}
attachDateMasks();

// Edição em massa de riscos
(function initBulkRiskEdit(){
  const form = document.getElementById('bulkRiskEditForm');
  const countEl = document.getElementById('bulkRiskCount');
  if (!form) return;

  function riskBoxes() {
    return Array.from(document.querySelectorAll('input[name="bulk_risk_ids"]'));
  }
  function selectedRiskBoxes() {
    return riskBoxes().filter((box) => box.checked);
  }
  function selectedApplyFields() {
    return Array.from(form.querySelectorAll('input[name="bulk_apply_fields"]:checked'));
  }
  function fieldLabel(box) {
    const label = box.closest('.bulk-field');
    return label ? (label.querySelector('span') || label).textContent.replace(/\s+/g, ' ').replace(/^Aplicar\s+/i, '').trim() : box.value;
  }
  function updateCount() {
    const total = selectedRiskBoxes().length;
    if (countEl) countEl.textContent = `${total} risco(s) selecionado(s)`;
  }
  riskBoxes().forEach((box) => box.addEventListener('change', updateCount));

  const selectVisible = document.getElementById('selectVisibleRisksBtn');
  if (selectVisible) {
    selectVisible.addEventListener('click', () => {
      riskBoxes().forEach((box) => {
        const row = box.closest('.compact-row');
        if (!row || row.style.display !== 'none') box.checked = true;
      });
      updateCount();
    });
  }

  const clear = document.getElementById('clearBulkRisksBtn');
  if (clear) {
    clear.addEventListener('click', () => {
      riskBoxes().forEach((box) => { box.checked = false; });
      updateCount();
    });
  }

  form.addEventListener('submit', (event) => {
    const selectedRisks = selectedRiskBoxes();
    const fields = selectedApplyFields();
    if (!selectedRisks.length) {
      event.preventDefault();
      alert('Selecione pelo menos um risco na lista.');
      return;
    }
    if (!fields.length) {
      event.preventDefault();
      alert('Marque pelo menos um campo com a opção Aplicar.');
      return;
    }

    const emptyFields = fields.filter((box) => {
      const field = form.querySelector(`[name="bulk_${CSS.escape(box.value)}"]`);
      return !field || !String(field.value || '').trim();
    });
    if (emptyFields.length) {
      event.preventDefault();
      alert(`Preencha os campos marcados ou desmarque-os:\n- ${emptyFields.map(fieldLabel).join('\n- ')}`);
      return;
    }

    const riskNames = selectedRisks.slice(0, 8).map((box) => box.dataset.bulkRiskName || box.value);
    const more = selectedRisks.length > riskNames.length ? `\n...e mais ${selectedRisks.length - riskNames.length} risco(s)` : '';
    const message = [
      `Você vai alterar ${selectedRisks.length} risco(s).`,
      '',
      'Campos que serão alterados:',
      `- ${fields.map(fieldLabel).join('\n- ')}`,
      '',
      'Riscos selecionados:',
      `- ${riskNames.join('\n- ')}${more}`,
      '',
      'Confirmar alterações em massa?'
    ].join('\n');
    if (!confirm(message)) event.preventDefault();
  });

  updateCount();
})();


function selectedPgrSectorIds() {
  return new Set(Array.from(document.querySelectorAll('input[name="pgr_sector_ids"]:checked')).map((box) => box.value));
}

function syncSectorPickers() {
  const select = document.getElementById('groupFilter');
  const group = select ? select.value : '';
  const selected = selectedPgrSectorIds();
  document.querySelectorAll('.sector-picker').forEach((el) => {
    const matchesGroup = !group || el.dataset.group === group;
    const matchesSector = selected.has(el.dataset.sector);
    const show = matchesGroup && matchesSector;
    el.style.display = show ? '' : 'none';
    if (!show) el.removeAttribute('open');
  });
}

function filterGroupSectors() {
  const select = document.getElementById('groupFilter');
  if (!select) return;
  const group = select.value;
  document.querySelectorAll('.sector-option').forEach((el) => {
    const show = !group || el.dataset.group === group;
    el.style.display = show ? '' : 'none';
  });
  syncSectorPickers();
}
const groupFilter = document.getElementById('groupFilter');
if (groupFilter) groupFilter.addEventListener('change', filterGroupSectors);
document.querySelectorAll('input[name="pgr_sector_ids"]').forEach((box) => {
  box.addEventListener('change', syncSectorPickers);
});
syncSectorPickers();

function cargoRowTemplate() {
  const div = document.createElement('div');
  div.className = 'cargo-row';
  div.innerHTML = `
    <label><span>Cargo</span><input name="cargo_nome[]" required placeholder="Ex.: Auxiliar Administrativo"></label>
    <label><span>CBO</span><input name="cargo_cbo[]" required placeholder="Ex.: 4110-05"></label>
    <label><span>Nº funcionários</span><input name="cargo_nfunc[]" required placeholder="Ex.: 3"></label>
    <label class="span-3"><span>Descrição da atividade</span><textarea name="cargo_descricao[]" required rows="2" placeholder="Descreva as atividades principais do cargo."></textarea></label>
    <div class="cargo-actions"><button class="btn small ghost" type="button" onclick="removeCargoRow(this)">Remover</button></div>`;
  return div;
}
function addCargoRow() {
  const list = document.getElementById('cargoList');
  if (list) list.appendChild(cargoRowTemplate());
}
function removeCargoRow(button) {
  const list = document.getElementById('cargoList');
  const rows = list ? list.querySelectorAll('.cargo-row') : [];
  if (rows.length <= 1) { alert('Mantenha pelo menos um cargo no setor.'); return; }
  button.closest('.cargo-row').remove();
}

function togglePgrRisks(checked) {
  document.querySelectorAll('.pgr-risk-check').forEach((box) => { box.checked = checked; });
}
function togglePgrExams(checked) {
  document.querySelectorAll('.pgr-exam-check').forEach((box) => { box.checked = checked; });
}
function syncRevisionField() {
  const ajuste = document.getElementById('ajustePsicossocial');
  const field = document.getElementById('revisionField');
  const input = document.getElementById('dataDaRevisao');
  if (!ajuste || !field || !input) return;
  const active = ajuste.checked;
  field.style.display = active ? '' : 'none';
  input.required = active;
  if (!active) input.value = '';
}
const ajusteBox = document.getElementById('ajustePsicossocial');
if (ajusteBox) {
  ajusteBox.addEventListener('change', syncRevisionField);
  syncRevisionField();
}


function compactGenerationPayload(form) {
  if (!form || form.dataset.payloadCompacted === '1') return;
  const selectedSectors = new Set(
    Array.from(form.querySelectorAll('input[name="pgr_sector_ids"]:checked, input[name="sector_ids"]:checked')).map((input) => input.value)
  );
  form.querySelectorAll('[data-sector]').forEach((node) => {
    const sectorId = node.getAttribute('data-sector');
    if (!sectorId || selectedSectors.has(sectorId)) return;
    node.querySelectorAll('input, select, textarea').forEach((field) => {
      field.dataset.prunedDisabled = '1';
      field.disabled = true;
    });
  });
  form.querySelectorAll('input[type="text"], input:not([type]), textarea').forEach((field) => {
    if (field.disabled || field.required) return;
    const name = field.getAttribute('name') || '';
    if (!name.startsWith('aet_')) return;
    if (!String(field.value || '').trim()) {
      field.dataset.prunedDisabled = '1';
      field.disabled = true;
    }
  });
  form.dataset.payloadCompacted = '1';
}

function restoreCompactedPayload(form) {
  if (!form) return;
  form.querySelectorAll('[data-pruned-disabled="1"]').forEach((field) => {
    field.disabled = false;
    delete field.dataset.prunedDisabled;
  });
  delete form.dataset.payloadCompacted;
}


function compactApplyImportedPayload(form) {
  if (!form || form.dataset.applyPayloadCompacted === '1') return;
  const allowedNames = new Set(['company_id', 'imported_template_id', 'imported_template_group_id']);
  form.querySelectorAll('input, select, textarea').forEach((field) => {
    const name = field.getAttribute('name') || '';
    if (!allowedNames.has(name)) {
      field.dataset.applyPrunedDisabled = '1';
      field.disabled = true;
    }
  });
  form.dataset.applyPayloadCompacted = '1';
}

const pgrForm = document.getElementById('pgrSelectionForm');
if (pgrForm) {
  pgrForm.addEventListener('submit', (event) => {
    const submitter = event.submitter;
    const action = submitter ? (submitter.getAttribute('formaction') || '') : '';
    if (action.includes('/aplicar-modelo-importado')) {
      const company = document.getElementById('companySelect');
      const model = document.getElementById('importedTemplateSelect');
      if (company && !company.value) {
        event.preventDefault();
        company.focus();
        alert('Selecione a empresa antes de aplicar o modelo importado.');
        return;
      }
      if (model && !model.value) {
        event.preventDefault();
        model.focus();
        alert('Selecione um modelo importado antes de aplicar.');
        return;
      }
      compactApplyImportedPayload(pgrForm);
      return;
    }
    const isComplete = action.includes('gerar-pgr-completo') || action.includes('gerar-pcmso-completo') || action.includes('gerar-ltcat-completo') || action.includes('gerar-pgr-aet-psicossocial') || action.includes('gerar-aet-completa') || action.includes('gerar-pacote-empresa');
    if (!isComplete) return;

    const company = document.getElementById('companySelect');
    const dataCriacao = document.getElementById('dataCriacaoLaudo');
    const dataRevisao = document.getElementById('dataDaRevisao');
    const ajuste = document.getElementById('ajustePsicossocial');
    if (company && !company.value) {
      event.preventDefault();
      company.focus();
      alert('Selecione uma empresa antes de gerar o laudo. Suas marcações foram mantidas.');
      return;
    }
    if (dataCriacao && !dataCriacao.value.trim()) {
      event.preventDefault();
      dataCriacao.focus();
      alert('Preencha a Data de criação do laudo antes de gerar. Suas marcações foram mantidas.');
      return;
    }
    if (ajuste && ajuste.checked && dataRevisao && !dataRevisao.value.trim()) {
      event.preventDefault();
      dataRevisao.focus();
      alert('Preencha a Data da revisão psicossocial antes de gerar. Suas marcações foram mantidas.');
      return;
    }
    if (action.includes('gerar-pgr-aet-psicossocial')) {
      const psychFile = pgrForm.querySelector('input[name="psicossocial_pdf"]');
      if (psychFile && !psychFile.value) {
        event.preventDefault();
        psychFile.focus();
        alert('Envie o Relatório Psicossocial em PDF. Suas marcações foram mantidas. A AET será gerada automaticamente se você não enviar uma AET externa.');
        return;
      }
    }
    compactGenerationPayload(pgrForm);
    if (submitter && action.includes('gerar-pgr-aet-psicossocial') && !submitter.dataset.ajaxDownload) {
      submitter.disabled = true;
      submitter.textContent = 'Juntando arquivos... aguarde';
    }
  });
}

// Seleção compacta de riscos na geração dos laudos
function riskLabelOriginalIndex(label) {
  const raw = label.getAttribute('data-index') || '0';
  const parsed = parseInt(raw, 10);
  return Number.isNaN(parsed) ? 0 : parsed;
}

function riskIdsFromGroupBox(groupBox) {
  return (groupBox.dataset.riskIds || '')
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function riskInputInSector(sectorId, riskId) {
  const list = document.querySelector(`[data-risk-list="${sectorId}"]`);
  if (!list) return null;
  return Array.from(list.querySelectorAll('.pgr-risk-check')).find((box) => box.value === riskId) || null;
}

function applySectorRiskGroup(groupBox) {
  const sectorId = groupBox.dataset.sector;
  const checked = groupBox.checked;
  riskIdsFromGroupBox(groupBox).forEach((riskId) => {
    const input = riskInputInSector(sectorId, riskId);
    if (input) input.checked = checked;
  });
}

function applyCheckedRiskGroupsOnLoad() {
  document.querySelectorAll('.sector-risk-group-check:checked').forEach((groupBox) => {
    applySectorRiskGroup(groupBox);
  });
}

function syncSectorRiskGroupStates() {
  document.querySelectorAll('.sector-risk-group-check').forEach((groupBox) => {
    const sectorId = groupBox.dataset.sector;
    const ids = riskIdsFromGroupBox(groupBox);
    if (!sectorId || ids.length === 0) {
      groupBox.checked = false;
      groupBox.indeterminate = false;
      return;
    }
    const states = ids.map((riskId) => {
      const input = riskInputInSector(sectorId, riskId);
      return !!(input && input.checked);
    });
    const allSelected = states.length > 0 && states.every(Boolean);
    const someSelected = states.some(Boolean);
    groupBox.checked = allSelected;
    groupBox.indeterminate = someSelected && !allSelected;
  });
}

function updateRiskSelectionLists() {
  let totalSelected = 0;
  document.querySelectorAll('.risk-list-mode').forEach((list) => {
    const labels = Array.from(list.querySelectorAll('.risk-list-row'));
    let selectedCount = 0;

    labels.forEach((label) => {
      const input = label.querySelector('.pgr-risk-check');
      const isChecked = !!(input && input.checked);
      label.classList.toggle('is-selected', isChecked);
      if (isChecked) selectedCount += 1;
    });

    totalSelected += selectedCount;

    labels
      .sort((a, b) => {
        const aChecked = a.classList.contains('is-selected') ? 1 : 0;
        const bChecked = b.classList.contains('is-selected') ? 1 : 0;
        if (aChecked !== bChecked) return bChecked - aChecked;
        return riskLabelOriginalIndex(a) - riskLabelOriginalIndex(b);
      })
      .forEach((label) => list.appendChild(label));

    const picker = list.closest('.sector-picker');
    if (picker) {
      const countText = `${selectedCount} ${selectedCount === 1 ? 'risco selecionado' : 'riscos selecionados'}`;
      picker.querySelectorAll('[data-risk-count], [data-risk-count-badge]').forEach((el) => {
        el.textContent = countText;
      });
    }
  });

  const total = document.getElementById('totalRiskSelected');
  if (total) {
    total.textContent = `${totalSelected} ${totalSelected === 1 ? 'risco selecionado' : 'riscos selecionados'}`;
  }

  syncSectorRiskGroupStates();
}

applyCheckedRiskGroupsOnLoad();

document.querySelectorAll('.pgr-risk-check').forEach((box) => {
  box.addEventListener('change', updateRiskSelectionLists);
});
document.querySelectorAll('.sector-risk-group-check').forEach((box) => {
  box.addEventListener('change', () => {
    applySectorRiskGroup(box);
    updateRiskSelectionLists();
  });
});
updateRiskSelectionLists();

// V20 - cadastro rápido de risco dentro da tela Gerar laudos, sem recarregar a página
function openModalElement(modal) {
  if (!modal) return;
  modal.classList.add('is-open');
  modal.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
  const firstInput = modal.querySelector('input[name="risco"]');
  if (firstInput) setTimeout(() => firstInput.focus(), 50);
}

function closeModalElement(modal) {
  if (!modal) return;
  modal.classList.remove('is-open');
  modal.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('modal-open');
}

function showQuickRiskErrors(errors) {
  const box = document.getElementById('quickRiskErrors');
  if (!box) return;
  const list = Array.isArray(errors) ? errors : [String(errors || 'Erro ao salvar o risco.')];
  box.innerHTML = list.map((error) => `<div>${error}</div>`).join('');
  box.style.display = list.length ? '' : 'none';
}

function clearQuickRiskErrors() {
  const box = document.getElementById('quickRiskErrors');
  if (!box) return;
  box.innerHTML = '';
  box.style.display = 'none';
}

function createRiskOptionLabel(risk, sectorId, checked) {
  const label = document.createElement('label');
  label.className = `tiny-check risk-option risk-list-row${checked ? ' is-selected' : ''}`;
  label.setAttribute('data-index', '-1');
  label.dataset.search = `${risk.risco || ''} ${risk.tipo_risco || ''} ${risk.grau_nivel_risco || ''}`.toLowerCase();
  label.dataset.dynamicRiskId = risk.id;

  const input = document.createElement('input');
  input.className = 'pgr-risk-check';
  input.type = 'checkbox';
  input.name = `sector_risk_ids_${sectorId}`;
  input.value = risk.id;
  input.checked = !!checked;

  const title = document.createElement('span');
  title.textContent = risk.risco || '';

  const meta = document.createElement('em');
  meta.textContent = `${risk.tipo_risco || ''} · ${risk.grau_nivel_risco || ''}`;

  label.appendChild(input);
  label.appendChild(title);
  label.appendChild(meta);
  return label;
}

function addRiskToGenerationLists(risk, autoSelect) {
  const selectedSectors = selectedPgrSectorIds();
  const searchInput = document.getElementById('riskPickerSearch');
  const term = searchInput ? searchInput.value.trim().toLowerCase() : '';

  document.querySelectorAll('.risk-list-mode').forEach((list) => {
    const sectorId = list.dataset.riskList;
    if (!sectorId) return;
    if (list.querySelector(`input.pgr-risk-check[value="${CSS.escape(risk.id)}"]`)) return;
    const checked = !!autoSelect && selectedSectors.has(sectorId);
    const label = createRiskOptionLabel(risk, sectorId, checked);
    if (term && !label.dataset.search.includes(term)) label.style.display = 'none';
    list.prepend(label);
  });

  updateRiskSelectionLists();
}

const quickRiskModal = document.getElementById('quickRiskModal');
const openQuickRiskModalBtn = document.getElementById('openQuickRiskModal');
const closeQuickRiskModalBtn = document.getElementById('closeQuickRiskModal');
const cancelQuickRiskModalBtn = document.getElementById('cancelQuickRiskModal');
const quickRiskForm = document.getElementById('quickRiskForm');

if (openQuickRiskModalBtn && quickRiskModal) {
  openQuickRiskModalBtn.addEventListener('click', () => {
    clearQuickRiskErrors();
    openModalElement(quickRiskModal);
  });
}
if (closeQuickRiskModalBtn && quickRiskModal) closeQuickRiskModalBtn.addEventListener('click', () => closeModalElement(quickRiskModal));
if (cancelQuickRiskModalBtn && quickRiskModal) cancelQuickRiskModalBtn.addEventListener('click', () => closeModalElement(quickRiskModal));
if (quickRiskModal) {
  quickRiskModal.addEventListener('click', (event) => {
    if (event.target === quickRiskModal) closeModalElement(quickRiskModal);
  });
}
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && quickRiskModal && quickRiskModal.classList.contains('is-open')) {
    closeModalElement(quickRiskModal);
  }
});

document.addEventListener('change', (event) => {
  if (event.target && event.target.classList && event.target.classList.contains('pgr-risk-check')) {
    updateRiskSelectionLists();
  }
});

if (quickRiskForm) {
  quickRiskForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearQuickRiskErrors();

    if (!quickRiskForm.reportValidity()) return;

    const button = document.getElementById('saveQuickRiskBtn');
    const originalText = button ? button.textContent : '';
    if (button) {
      button.disabled = true;
      button.textContent = 'Salvando...';
    }

    try {
      const response = await fetch(pgrUrl('/api/risco/novo'), {
        method: 'POST',
        body: new FormData(quickRiskForm),
        headers: { 'X-Requested-With': 'fetch' },
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok || !payload.ok) {
        showQuickRiskErrors(payload.errors || ['Não foi possível salvar o risco. Confira os campos e tente novamente.']);
        return;
      }

      const autoSelect = !!document.getElementById('quickRiskAutoSelect')?.checked;
      addRiskToGenerationLists(payload.risk, autoSelect);
      quickRiskForm.reset();
      const fontes = quickRiskForm.querySelector('[name="fontes_circunstancias"]');
      const jornada = quickRiskForm.querySelector('[name="ltcat_periodicidade_jornada"]');
      const insal = quickRiskForm.querySelector('[name="ltcat_insalubridade"]');
      const grauInsal = quickRiskForm.querySelector('[name="ltcat_grau_insalubridade"]');
      const aposentadoria = quickRiskForm.querySelector('[name="ltcat_aposentadoria_especial"]');
      const autoSelectBox = document.getElementById('quickRiskAutoSelect');
      if (fontes) fontes.value = 'Durante o processo de trabalho.';
      if (jornada) jornada.value = 'Mensal (<= 4 horas < 10% jornada)';
      if (insal) insal.value = 'Não';
      if (grauInsal) grauInsal.value = 'Não aplicável';
      if (aposentadoria) aposentadoria.value = 'Não';
      if (autoSelectBox) autoSelectBox.checked = true;
      closeModalElement(quickRiskModal);
      alert('Risco cadastrado sem recarregar a página. Ele já está disponível na lista de riscos.');
    } catch (error) {
      showQuickRiskErrors(['Erro de conexão ao salvar o risco. Tente novamente.']);
    } finally {
      if (button) {
        button.disabled = false;
        button.textContent = originalText;
      }
    }
  });
}


// Configurações salvas da tela Gerar laudos
function readReportProfiles() {
  const node = document.getElementById('reportProfilesData');
  if (!node) return [];
  try { return JSON.parse(node.textContent || '[]'); } catch (_) { return []; }
}
const reportProfiles = readReportProfiles();
const reportProfilesById = Object.fromEntries(reportProfiles.map((profile) => [profile.id, profile]));

function filterProfileOptions() {
  const companySelect = document.getElementById('companySelect');
  const profileSelect = document.getElementById('profileSelect');
  if (!companySelect || !profileSelect) return;
  const companyId = companySelect.value;
  let visibleCount = 0;
  Array.from(profileSelect.options).forEach((option) => {
    if (!option.value) return;
    const show = !!companyId && option.dataset.company === companyId;
    option.disabled = !show;
    option.hidden = !show;
    if (show) visibleCount += 1;
  });
  const current = profileSelect.selectedOptions[0];
  if (current && current.value && current.disabled) profileSelect.value = '';
  profileSelect.title = visibleCount ? `${visibleCount} configuração(ões) salvas para esta empresa` : 'Nenhuma configuração salva para esta empresa';
}

function setCheckboxesByName(name, values) {
  const set = new Set(values || []);
  document.querySelectorAll(`input[name="${name}"]`).forEach((box) => { box.checked = set.has(box.value); });
}

function setFieldValueByName(name, value) {
  const field = document.querySelector(`[name="${CSS.escape(name)}"]`);
  if (!field) return;
  field.value = value || '';
}

function loadAetState(state) {
  const aet = state.aet || {};
  const general = aet.general || {};
  const bySector = aet.by_sector || {};
  setFieldValueByName('aet_tipo', general.tipo_aet || '');
  setFieldValueByName('aet_responsavel_tecnico', general.responsavel_tecnico || '');
  setCheckboxesByName('aet_metodologia', general.metodologia || []);
  setFieldValueByName('aet_objetivo_complementar', general.objetivo_complementar || '');
  setFieldValueByName('aet_criterios_analise', general.criterios_analise || '');
  setFieldValueByName('aet_conclusao_geral', general.conclusao_geral_manual || '');
  Object.entries(bySector).forEach(([sectorId, data]) => {
    setCheckboxesByName(`aet_postura_${sectorId}`, data.postura_predominante || []);
    ['exigencia_fisica','exigencia_cognitiva','ritmo_trabalho','pausas','mobiliario','ambiente','organizacao','equipamentos','queixas','observacoes','recomendacoes','prioridade','prazo','responsavel','conclusao_setor'].forEach((key) => {
      setFieldValueByName(`aet_${key}_${sectorId}`, data[key] || '');
    });
  });
}

function clearGenerationSelections() {
  document.querySelectorAll('input[name="pgr_sector_ids"], .pgr-risk-check, .pgr-exam-check, .sector-risk-group-check').forEach((box) => {
    box.checked = false;
    box.indeterminate = false;
  });
}

function loadReportProfile(profileId) {
  const profile = reportProfilesById[profileId];
  if (!profile) return;
  const state = profile.state || {};
  const companySelect = document.getElementById('companySelect');
  if (companySelect) companySelect.value = profile.company_id || state.company_id || '';
  const dataCriacao = document.getElementById('dataCriacaoLaudo');
  if (dataCriacao) dataCriacao.value = state.data_criacao_laudo || profile.data_criacao_laudo || '';
  const ajuste = document.getElementById('ajustePsicossocial');
  if (ajuste) ajuste.checked = (state.ajuste_psicossocial || profile.ajuste_psicossocial) === '1';
  const dataRevisao = document.getElementById('dataDaRevisao');
  if (dataRevisao) dataRevisao.value = state.data_da_revisao || profile.data_da_revisao || '';
  const profileName = document.getElementById('profileName');
  if (profileName) profileName.value = profile.nome || '';

  clearGenerationSelections();
  setCheckboxesByName('pgr_sector_ids', state.selected_sector_ids || []);
  Object.entries(state.selected_risk_ids_by_sector || {}).forEach(([sectorId, riskIds]) => setCheckboxesByName(`sector_risk_ids_${sectorId}`, riskIds));
  Object.entries(state.selected_risk_group_ids_by_sector || {}).forEach(([sectorId, groupIds]) => setCheckboxesByName(`sector_risk_group_ids_${sectorId}`, groupIds));
  Object.entries(state.selected_exam_ids_by_sector || {}).forEach(([sectorId, examIds]) => setCheckboxesByName(`sector_exam_ids_${sectorId}`, examIds));
  loadAetState(state);

  filterProfileOptions();
  syncRevisionField();
  syncSectorPickers();
  applyCheckedRiskGroupsOnLoad();
  updateRiskSelectionLists();
  alert('Configuração carregada. Confira os dados e gere o laudo desejado.');
}

const companySelectForProfiles = document.getElementById('companySelect');
if (companySelectForProfiles) companySelectForProfiles.addEventListener('change', filterProfileOptions);
const loadProfileBtn = document.getElementById('loadProfileBtn');
if (loadProfileBtn) {
  loadProfileBtn.addEventListener('click', () => {
    const profileSelect = document.getElementById('profileSelect');
    if (!profileSelect || !profileSelect.value) {
      alert('Selecione uma configuração salva para carregar.');
      return;
    }
    loadReportProfile(profileSelect.value);
  });
}
filterProfileOptions();

// Preenche o nome da empresa na aba de juntar arquivos avulsos
const joinCompanySelect = document.getElementById('joinCompanySelect');
if (joinCompanySelect) {
  joinCompanySelect.addEventListener('change', () => {
    const input = document.getElementById('joinCompanyName');
    const option = joinCompanySelect.selectedOptions[0];
    if (input && option && option.dataset.name) input.value = option.dataset.name;
  });
}

function getDownloadFilename(disposition, fallback) {
  if (!disposition) return fallback;
  const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match && utf8Match[1]) return decodeURIComponent(utf8Match[1].replace(/"/g, ''));
  const match = disposition.match(/filename="?([^";]+)"?/i);
  return match && match[1] ? match[1] : fallback;
}

function setAjaxStatus(form, message, type) {
  const status = form.querySelector('[data-ajax-status]') || document.querySelector('[data-ajax-status]');
  if (!status) return;
  status.textContent = message || '';
  status.classList.remove('success', 'error', 'loading');
  if (type) status.classList.add(type);
}

function resetFormButtons(form) {
  form.querySelectorAll('button[type="submit"]').forEach((button) => {
    if (button.dataset.originalText) button.textContent = button.dataset.originalText;
    button.disabled = false;
  });
}

async function handleAjaxDownload(event, form, submitter, action) {
  event.preventDefault();
  const activeButton = submitter || form.querySelector('button[type="submit"]');
  const originalText = activeButton ? activeButton.textContent : '';
  if (activeButton) {
    activeButton.dataset.originalText = originalText;
    activeButton.disabled = true;
    activeButton.textContent = activeButton.dataset.processingText || 'Processando... aguarde';
  }
  setAjaxStatus(form, 'Processando arquivo. Aguarde o download iniciar...', 'loading');

  try {
    const response = await fetch(action || form.action, {
      method: (form.method || 'POST').toUpperCase(),
      body: new FormData(form),
      headers: { 'X-Requested-With': 'XMLHttpRequest' },
    });

    if (!response.ok) {
      let errorMessage = 'Não foi possível gerar o arquivo.';
      const contentType = response.headers.get('content-type') || '';
      if (contentType.includes('application/json')) {
        const data = await response.json();
        errorMessage = data.error || data.message || errorMessage;
      } else {
        const text = await response.text();
        if (text) errorMessage = text.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 600) || errorMessage;
      }
      throw new Error(errorMessage);
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = getDownloadFilename(response.headers.get('content-disposition'), 'documento_gerado.docx');
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(url);

    form.querySelectorAll('input[type="file"]').forEach((input) => { input.value = ''; });
    setAjaxStatus(form, 'Arquivo gerado. Você já pode selecionar novos arquivos e juntar outro documento sem recarregar a página.', 'success');
  } catch (error) {
    setAjaxStatus(form, error.message || 'Erro ao processar arquivo.', 'error');
    alert(error.message || 'Erro ao processar arquivo.');
  } finally {
    restoreCompactedPayload(form);
    resetFormButtons(form);
  }
}

// Download por AJAX: mantém a página aberta após juntar arquivos e permite gerar outros sem recarregar.
document.querySelectorAll('form').forEach((form) => {
  form.addEventListener('submit', (event) => {
    if (event.defaultPrevented) return;
    const submitter = event.submitter;
    const action = submitter?.getAttribute('formaction') || form.getAttribute('action') || window.location.href;
    const shouldAjaxDownload = form.dataset.ajaxDownload === 'true' || submitter?.dataset.ajaxDownload === 'true' || action.includes('gerar-pgr-aet-psicossocial') || action.includes('gerar-aet-completa') || action.includes('gerar-pacote-empresa');
    if (!shouldAjaxDownload) return;
    if (!form.checkValidity()) return;
    handleAjaxDownload(event, form, submitter, action);
  });
});

// V26 - geração em etapas tipo wizard com barra de progresso
(function initWizardSteps(){
  const form = document.getElementById('pgrSelectionForm');
  if (!form) return;
  const stepNodes = Array.from(form.querySelectorAll('[data-wizard-step]'));
  const navs = Array.from(form.querySelectorAll('[data-step-nav]'));
  if (!stepNodes.length || !navs.length) return;
  let activeStep = Number(sessionStorage.getItem('sstWizardActiveStep') || '1');
  const maxStep = navs.reduce((max, nav) => Math.max(max, Number(nav.dataset.stepNav || '1')), 1);
  if (!activeStep || activeStep < 1 || activeStep > maxStep) activeStep = 1;

  function showStep(step){
    activeStep = step;
    sessionStorage.setItem('sstWizardActiveStep', String(step));
    stepNodes.forEach((node) => {
      const nodeStep = Number(node.dataset.wizardStep || '1');
      node.classList.toggle('wizard-hidden', nodeStep !== step);
    });
    navs.forEach((nav) => {
      const navStep = Number(nav.dataset.stepNav || '1');
      nav.classList.toggle('is-active', navStep === step);
      nav.classList.toggle('is-complete', navStep < step);
    });
    window.scrollTo({top:0, behavior:'smooth'});
  }

  function addStepControls(){
    for (let step=1; step<=maxStep; step++) {
      const panels = stepNodes.filter((node) => Number(node.dataset.wizardStep || '0') === step && node.classList.contains('step-panel'));
      const panel = panels[panels.length-1];
      const content = panel ? panel.querySelector('.step-content') : null;
      if (!content || content.querySelector('.wizard-nav-controls')) continue;
      const controls = document.createElement('div');
      controls.className = 'wizard-nav-controls';
      const back = document.createElement('button');
      back.type = 'button';
      back.className = 'btn ghost';
      back.textContent = step === 1 ? 'Início' : 'Voltar';
      back.disabled = step === 1;
      back.addEventListener('click', () => showStep(Math.max(1, step-1)));
      const next = document.createElement('button');
      next.type = 'button';
      next.className = 'btn primary';
      next.textContent = step === maxStep ? 'Conferir e gerar' : 'Próximo';
      next.addEventListener('click', () => showStep(Math.min(maxStep, step+1)));
      controls.appendChild(back);
      controls.appendChild(next);
      content.appendChild(controls);
    }
  }

  navs.forEach((nav) => nav.addEventListener('click', () => showStep(Number(nav.dataset.stepNav || '1'))));
  addStepControls();
  showStep(activeStep);
})();

// V26 - motor de regras para sugerir exames com base nos riscos selecionados
const suggestExamsBtn = document.getElementById('suggestExamsBtn');
if (suggestExamsBtn && pgrForm) {
  suggestExamsBtn.addEventListener('click', async () => {
    suggestExamsBtn.disabled = true;
    const originalText = suggestExamsBtn.textContent;
    suggestExamsBtn.textContent = 'Sugerindo...';
    try {
      const fd = new FormData();
      selectedPgrSectorIds().forEach((sectorId) => {
        fd.append('pgr_sector_ids', sectorId);
        document.querySelectorAll(`input[name="sector_risk_ids_${CSS.escape(sectorId)}"]:checked`).forEach((riskBox) => {
          fd.append(`sector_risk_ids_${sectorId}`, riskBox.value);
        });
      });
      const response = await fetch(pgrUrl('/api/sugerir-exames'), {
        method: 'POST',
        body: fd,
        headers: { 'X-Requested-With': 'XMLHttpRequest' },
      });
      const payload = await response.json();
      if (!response.ok || !payload.ok) throw new Error(payload.error || 'Não foi possível sugerir exames.');
      let total = 0;
      Object.entries(payload.suggestions || {}).forEach(([sectorId, examIds]) => {
        (examIds || []).forEach((examId) => {
          const input = pgrForm.querySelector(`input[name="sector_exam_ids_${CSS.escape(sectorId)}"][value="${CSS.escape(examId)}"]`);
          if (input && !input.checked) {
            input.checked = true;
            total += 1;
            const label = input.closest('.exam-option');
            if (label) {
              label.classList.add('suggested-exam-highlight');
              setTimeout(() => label.classList.remove('suggested-exam-highlight'), 2500);
            }
          }
        });
      });
      alert(total ? `${total} exame(s) sugerido(s) e marcado(s) pelas regras técnicas.` : 'Nenhum exame novo foi sugerido. Confira se os exames estão cadastrados e se há riscos selecionados.');
    } catch (error) {
      alert(error.message || 'Erro ao sugerir exames.');
    } finally {
      suggestExamsBtn.disabled = false;
      suggestExamsBtn.textContent = originalText;
    }
  });
}

// V28 - pré-preenchimento técnico do formulário AET por CNAE/atividade
(function initAetCnaePresets(){
  const presetEl = document.getElementById('aetPresetsData');
  const companiesEl = document.getElementById('companiesData');
  const button = document.getElementById('applyAetPresetBtn');
  const status = document.getElementById('aetPresetStatus');
  const companySelect = document.getElementById('companySelect');
  const form = document.getElementById('pgrSelectionForm');
  if (!presetEl || !companiesEl || !button || !companySelect || !form) return;

  let presets = [];
  let companies = [];
  try { presets = JSON.parse(presetEl.textContent || '[]'); } catch (_) { presets = []; }
  try { companies = JSON.parse(companiesEl.textContent || '[]'); } catch (_) { companies = []; }

  function norm(value){
    return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  }
  function selectedCompany(){
    const id = companySelect.value;
    return companies.find((item) => item.id === id) || null;
  }
  function findPreset(){
    const company = selectedCompany();
    if (!company) return null;
    const haystack = norm([company.cnae1, company.cnae2, company.descricao1, company.descricao2, company.nome].join(' '));
    return presets.find((preset) => (preset.keywords || []).some((kw) => haystack.includes(norm(kw)))) || null;
  }
  function updateStatus(){
    const preset = findPreset();
    if (!status) return;
    status.textContent = preset ? `Sugestão: ${preset.label}` : 'Sem preset específico';
    status.classList.toggle('is-ready', !!preset);
  }
  function setField(selector, value, onlyIfEmpty=false){
    const field = form.querySelector(selector);
    if (!field || value === undefined || value === null || value === '') return;
    if (onlyIfEmpty && field.value) return;
    field.value = value;
    field.dispatchEvent(new Event('change', {bubbles:true}));
  }
  function setGeneral(preset){
    const general = preset.general || {};
    Object.entries(general).forEach(([key, value]) => {
      setField(`[data-aet-general="${CSS.escape(key)}"]`, value);
    });
    if (general.tipo_documento) setField('[name="aet_tipo"]', general.tipo_documento);
  }
  function setSectorSelect(sectorEl, key, value){
    if (value === undefined || value === null || value === '') return;
    const field = sectorEl.querySelector(`[data-aet-sector="${CSS.escape(key)}"]`);
    if (field) {
      field.value = value;
      field.dispatchEvent(new Event('change', {bubbles:true}));
    }
  }
  function setSectorTextFieldByName(sectorEl, suffix, value){
    if (!value) return;
    const field = sectorEl.querySelector(`input[name^="aet_${suffix}_"], textarea[name^="aet_${suffix}_"]`);
    if (field) field.value = value;
  }
  function setSectorChecks(sectorEl, attr, values){
    const wanted = new Set((values || []).map(String));
    if (!wanted.size) return;
    sectorEl.querySelectorAll(`input[data-aet-sector-list="${attr}"]`).forEach((box) => {
      box.checked = wanted.has(box.value);
    });
  }
  function setPostures(sectorEl, values){
    const wanted = new Set((values || []).map(String));
    if (!wanted.size) return;
    sectorEl.querySelectorAll('input[name^="aet_postura_"]').forEach((box) => { box.checked = wanted.has(box.value); });
  }
  function applyPreset(){
    const preset = findPreset();
    if (!preset) {
      alert('Não encontrei um preset específico para o CNAE/atividade dessa empresa. Você ainda pode preencher manualmente.');
      return;
    }
    setGeneral(preset);
    const sectorPreset = preset.sector || {};
    document.querySelectorAll('.aet-sector-picker').forEach((sectorEl) => {
      if (sectorEl.style.display === 'none') return;
      setPostures(sectorEl, sectorPreset.postura || []);
      ['tipo_atividade','exigencia_fisica','exigencia_cognitiva','prioridade'].forEach((key) => setSectorSelect(sectorEl, key, sectorPreset[key]));
      ['ritmo_trabalho','pausas','mobiliario','ambiente','organizacao','equipamentos','prazo','responsavel'].forEach((key) => setSectorSelect(sectorEl, key, sectorPreset[key]));
      setSectorChecks(sectorEl, 'fatores', sectorPreset.fatores || []);
      setSectorChecks(sectorEl, 'medidas', sectorPreset.medidas || []);
    });
    if (status) status.textContent = `Aplicado: ${preset.label}`;
    alert(`Sugestões aplicadas para: ${preset.label}. Revise os campos antes de gerar a AET.`);
  }
  companySelect.addEventListener('change', updateStatus);
  button.addEventListener('click', applyPreset);
  updateStatus();
})();

// Tela de carregamento para aplicação de modelo importado.
// Evita a sensação de travamento quando o sistema precisa criar muitos setores, cargos e vínculos de riscos.
(function(){
  function isApplyImportedSubmitter(submitter){
    if (!submitter) return false;
    const action = submitter.getAttribute('formaction') || submitter.formAction || '';
    return String(action).includes('/aplicar-modelo-importado');
  }
  function ensureOverlay(){
    let overlay = document.getElementById('importApplyLoadingOverlay');
    if (overlay) return overlay;
    overlay = document.createElement('div');
    overlay.id = 'importApplyLoadingOverlay';
    overlay.className = 'loading-page-overlay';
    overlay.innerHTML = `
      <div class="loading-page-card" role="status" aria-live="polite">
        <div class="loading-badge">Importação inteligente</div>
        <h2>Aplicando modelo na empresa...</h2>
        <p class="loading-text" id="importApplyLoadingText">Preparando dados do laudo antigo.</p>
        <div class="progress-track"><div class="progress-bar" id="importApplyProgressBar"></div></div>
        <ol class="loading-steps">
          <li data-loading-step="1" class="is-active">Conferindo empresa selecionada</li>
          <li data-loading-step="2">Criando setores e cargos</li>
          <li data-loading-step="3">Cadastrando riscos ainda inexistentes</li>
          <li data-loading-step="4">Vinculando riscos por setor</li>
          <li data-loading-step="5">Preparando seleção para geração dos laudos</li>
        </ol>
        <small>Não feche esta aba. Em modelos grandes, essa etapa pode levar alguns instantes.</small>
      </div>`;
    document.body.appendChild(overlay);
    return overlay;
  }
  function showApplyLoading(){
    const overlay = ensureOverlay();
    overlay.classList.add('is-visible');
    const bar = overlay.querySelector('#importApplyProgressBar');
    const text = overlay.querySelector('#importApplyLoadingText');
    const steps = Array.from(overlay.querySelectorAll('[data-loading-step]'));
    const messages = [
      ['Conferindo empresa selecionada e modelo importado.', 12],
      ['Lendo setores e cargos do modelo.', 28],
      ['Cadastrando setores/cargos e evitando duplicidades.', 45],
      ['Cadastrando riscos e preservando vínculo por setor.', 65],
      ['Aplicando exames e preparando a tela de geração.', 82],
      ['Finalizando. Você será redirecionado automaticamente.', 94]
    ];
    let i = 0;
    function tick(){
      const current = messages[Math.min(i, messages.length - 1)];
      if (text) text.textContent = current[0];
      if (bar) bar.style.width = `${current[1]}%`;
      steps.forEach((step, idx) => {
        step.classList.toggle('is-active', idx === Math.min(i, steps.length - 1));
        step.classList.toggle('is-done', idx < Math.min(i, steps.length - 1));
      });
      if (i < messages.length - 1) {
        i += 1;
        window.setTimeout(tick, i < 3 ? 900 : 1400);
      } else {
        // Mantém uma animação leve até a resposta do servidor chegar.
        let pulse = 94;
        const id = window.setInterval(() => {
          if (!overlay.classList.contains('is-visible')) return window.clearInterval(id);
          pulse = pulse >= 98 ? 94 : pulse + 1;
          if (bar) bar.style.width = `${pulse}%`;
        }, 900);
      }
    }
    tick();
  }

  // Garante que o botão funcione mesmo com campos obrigatórios em outras etapas do wizard.
  // Para aplicar modelo importado, só precisamos de empresa + modelo; data do laudo/exames/AET
  // serão preenchidos depois, então desativamos a validação nativa apenas neste envio.
  const applyBtn = document.getElementById('applyImportedTemplateBtn');
  if (applyBtn) {
    applyBtn.addEventListener('click', function(){
      const form = document.getElementById('pgrSelectionForm');
      if (!form) return;
      form.noValidate = true;
      applyBtn.setAttribute('formnovalidate', 'formnovalidate');
    });
  }

  document.addEventListener('submit', function(ev){
    const form = ev.target;
    if (!form || !form.matches || !form.matches('#pgrSelectionForm')) return;
    if (isApplyImportedSubmitter(ev.submitter)) {
      const company = document.getElementById('companySelect');
      const model = document.getElementById('importedTemplateSelect');
      if ((company && !company.value) || (model && !model.value)) return;
      compactApplyImportedPayload(form);
      showApplyLoading();
      if (ev.submitter) {
        ev.submitter.disabled = true;
        ev.submitter.textContent = 'Aplicando modelo...';
      }
    }
  }, true);
})();

// V42 - busca simples em catálogos/telas compactas.
(function initGenericCardFilter(){
  document.querySelectorAll('[data-filter-target]').forEach((input) => {
    const selector = input.getAttribute('data-filter-target');
    if (!selector) return;
    const norm = (v) => String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    input.addEventListener('input', () => {
      const query = norm(input.value);
      document.querySelectorAll(selector).forEach((item) => {
        const text = norm(item.getAttribute('data-search') || item.textContent || '');
        item.style.display = !query || text.includes(query) ? '' : 'none';
      });
    });
  });
})();
