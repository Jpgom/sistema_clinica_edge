(() => {
  'use strict';

  const root = document.getElementById('receiptApp');
  const form = document.getElementById('receiptForm');
  // Perfis sem permissão de edição recebem apenas o aviso em HTML; não fazemos POST.
  if (!root || !form) return;

  const $ = id => document.getElementById(id);
  const rows = [...document.querySelectorAll('.receipt-item')];
  const image = $('receiptPreviewImage');
  const stage = $('receiptPreviewStage');
  const overlay = $('receiptFieldOverlay');
  const scroll = $('receiptPreviewScroll');
  const feedback = $('receiptFeedback');
  const warningsBox = $('receiptWarnings');
  const fieldSelect = $('calibrationField');
  const calibrationPanel = $('calibrationPanel');
  const scopeInputs = [...document.querySelectorAll('input[name="receiptScope"]')];
  const csrf = document.querySelector('meta[name="csrf-token"]')?.content || '';
  const endpoints = {
    preview: root.dataset.previewUrl,
    pdf: root.dataset.pdfUrl,
    config: root.dataset.configUrl
  };
  const names = {
    cliente: 'Cliente', data: 'Data', quantidade: 'Quantidade', unidade: 'Unidade',
    descricao: 'Descrição', preco_unitario: 'Preço unitário', total_item: 'Total do item',
    total_geral: 'Total geral'
  };
  const scopeDescriptions = {
    cell: 'Move somente o campo selecionado nesta linha.',
    line: 'Move todos os dados desta linha juntos.',
    field: 'Move este tipo de campo em todas as linhas.',
    general: 'Move todos os dados do recibo juntos.'
  };
  const zoomLevels = [1, 1.5, 2, 3];
  const copy = value => JSON.parse(JSON.stringify(value));

  let savedConfig;
  try { savedConfig = JSON.parse($('receiptInitialConfig').textContent); }
  catch (_) { savedConfig = {}; }
  let draftConfig = copy(savedConfig);
  let layout = null;
  let selectedId = null;
  let calibrating = false;
  let zoomIndex = 0;
  let previewTimer = null;
  let previewController = null;
  let previewSequence = 0;
  let visibleRows = 1;
  let dragState = null;
  let scopeTouchedByUser = false;

  const paperWidth = Number(savedConfig.papel_largura_mm) || 144.018;
  const paperHeight = Number(savedConfig.papel_altura_mm) || 105.41;

  function setFeedback(message, error = false) {
    feedback.textContent = message;
    feedback.classList.toggle('is-error', error);
  }

  function normalizedDecimal(value) {
    let raw = String(value ?? '').trim().replace(/\s/g, '');
    if (raw.startsWith('R$')) raw = raw.slice(2);
    if (raw.length > 32) return null;
    if (!raw) return '0';
    if (raw.includes(',')) {
      if (!/^[+-]?(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?$/.test(raw)) return null;
      return raw.replace(/\./g, '').replace(',', '.');
    }
    return /^[+-]?\d+(?:\.\d+)?$/.test(raw) ? raw : null;
  }

  function parseDecimal(value) {
    const normalized = normalizedDecimal(value);
    return normalized === null ? NaN : Number(normalized);
  }

  function decimalParts(value) {
    const normalized = normalizedDecimal(value);
    if (normalized === null || normalized.startsWith('-')) return null;
    const unsigned = normalized.startsWith('+') ? normalized.slice(1) : normalized;
    const [whole, fraction = ''] = unsigned.split('.');
    const units = BigInt(whole + fraction);
    const scale = 10n ** BigInt(fraction.length);
    if (units > 999999999n * scale) return null;
    return { units, scale };
  }

  function moneyFromCents(cents) {
    const integer = (cents / 100n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return `${integer},${(cents % 100n).toString().padStart(2, '0')}`;
  }

  function offsetText(value) { return String(Math.round(Number(value || 0) * 1000) / 1000).replace('.', ','); }
  function roundOffset(value) { return Math.round(value * 1000) / 1000; }

  function validDate(value) {
    const match = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(value || '').trim());
    if (!match) return false;
    const day = Number(match[1]);
    const month = Number(match[2]);
    const year = Number(match[3]);
    const test = new Date(year, month - 1, day);
    return year >= 1900 && test.getFullYear() === year && test.getMonth() === month - 1 && test.getDate() === day;
  }

  function validateForPdf() {
    const dateInput = $('receiptDate');
    const isDateValid = validDate(dateInput.value);
    dateInput.setAttribute('aria-invalid', String(!isDateValid));
    if (!isDateValid) {
      setFeedback('Informe uma data válida no formato DD/MM/AAAA.', true);
      dateInput.focus();
      return false;
    }
    if (!updateTotals()) {
      setFeedback('Corrija os valores destacados antes de gerar o PDF.', true);
      form.querySelector('[aria-invalid="true"]')?.focus();
      return false;
    }
    const payload = readPayload();
    const hasItem = payload.itens.some(item => item.descricao || parseDecimal(item.quantidade) || parseDecimal(item.preco_unitario));
    if (!payload.cliente && !hasItem) {
      setFeedback('Preencha o cliente ou pelo menos um item antes de gerar o PDF.', true);
      $('receiptClient').focus();
      return false;
    }
    return true;
  }

  function updateTotals() {
    let totalCents = 0n;
    let valid = true;
    for (const row of rows) {
      const qtyInput = row.querySelector('[data-field="quantidade"]');
      const priceInput = row.querySelector('[data-field="preco_unitario"]');
      const qty = decimalParts(qtyInput.value);
      const price = decimalParts(priceInput.value);
      const qtyValid = qty !== null;
      const priceValid = price !== null;
      qtyInput.setAttribute('aria-invalid', String(!qtyValid));
      priceInput.setAttribute('aria-invalid', String(!priceValid));
      if (qtyValid && priceValid) {
        const numerator = qty.units * price.units * 100n;
        const denominator = qty.scale * price.scale;
        const lineCents = (numerator * 2n + denominator) / (denominator * 2n);
        row.querySelector('output').textContent = moneyFromCents(lineCents);
        totalCents += lineCents;
      } else {
        row.querySelector('output').textContent = '—';
        valid = false;
      }
    }
    $('receiptTotal').textContent = `R$ ${moneyFromCents(totalCents)}`;
    $('generateReceiptPdf').disabled = !valid;
    $('generateReceiptProof').disabled = !valid;
    return valid;
  }

  function readPayload() {
    return {
      cliente: $('receiptClient').value.trim(),
      data: $('receiptDate').value.trim(),
      itens: rows.map(row => ({
        quantidade: row.querySelector('[data-field="quantidade"]').value.trim(),
        unidade: row.querySelector('[data-field="unidade"]').value.trim(),
        descricao: row.querySelector('[data-field="descricao"]').value.trim(),
        preco_unitario: row.querySelector('[data-field="preco_unitario"]').value.trim()
      }))
    };
  }

  function previewPayload() {
    if (!$('receiptExampleMode').checked) return readPayload();
    const date = $('receiptDate').value.trim() || new Date().toLocaleDateString('pt-BR');
    return {
      cliente: 'CLIENTE EXEMPLO', data: date,
      itens: Array.from({ length: 10 }, (_, index) => ({
        quantidade: '1', unidade: 'UN', descricao: `ITEM EXEMPLO ${index + 1}`,
        preco_unitario: '10,00'
      }))
    };
  }

  async function jsonPost(url, body, signal) {
    const response = await fetch(url, {
      method: 'POST', credentials: 'same-origin', signal,
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json', 'X-CSRF-Token': csrf },
      body: JSON.stringify(body)
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(result.error || result.message || `Falha na requisição (${response.status}).`);
    return result;
  }

  function schedulePreview(delay = 220) {
    clearTimeout(previewTimer);
    previewTimer = window.setTimeout(loadPreview, delay);
  }

  function hideOutdatedPreview(message) {
    ++previewSequence;
    previewController?.abort();
    layout = null;
    selectedId = null;
    image.hidden = true;
    overlay.replaceChildren();
    renderFieldSelect();
    $('receiptPreviewLoading').textContent = message;
    $('receiptPreviewLoading').hidden = false;
  }

  async function loadPreview() {
    if (!updateTotals()) {
      hideOutdatedPreview('Corrija os valores para ver a prévia atualizada.');
      setFeedback('Corrija os valores destacados para atualizar a prévia.', true);
      return;
    }
    const payload = previewPayload();
    if (payload.data && !validDate(payload.data)) {
      hideOutdatedPreview('Corrija a data para ver a prévia atualizada.');
      setFeedback('Termine de preencher a data para atualizar a prévia.');
      return;
    }
    const sequence = ++previewSequence;
    previewController?.abort();
    previewController = new AbortController();
    if (!image.src) setFeedback('Carregando prévia…');
    try {
      const result = await jsonPost(endpoints.preview,
        { payload, config: draftConfig }, previewController.signal);
      if (sequence !== previewSequence) return;
      if (!result.image || !result.layout) throw new Error('A prévia veio incompleta.');
      layout = result.layout;
      image.onload = () => { resizeStage(); renderOverlay(); };
      image.src = result.image;
      image.hidden = false;
      $('receiptPreviewLoading').hidden = true;
      renderFieldSelect();
      renderOverlay();
      renderWarnings(layout.warnings, layout.print_blockers);
      if (calibrating) refreshCalibration();
      setFeedback($('receiptExampleMode').checked ? 'Exemplo de alinhamento exibido.' : 'Prévia atualizada.');
    } catch (error) {
      if (error.name === 'AbortError' || sequence !== previewSequence) return;
      hideOutdatedPreview('Não foi possível atualizar a prévia.');
      setFeedback(error.message || 'Não foi possível atualizar a prévia.', true);
    }
  }

  function issueStrings(items) {
    return Array.isArray(items) ? items.map(item => typeof item === 'string' ? item : (item.message || String(item))) : [];
  }

  function fillIssueList(target, items) {
    target.replaceChildren();
    for (const issue of items) {
      const li = document.createElement('li');
      li.textContent = issue;
      target.append(li);
    }
  }

  function showPrintIssues(warnings, blockers) {
    const blocking = issueStrings(blockers);
    const all = [...new Set([...blocking, ...issueStrings(warnings)])];
    const container = $('receiptPrintIssues');
    fillIssueList($('receiptPrintIssuesList'), all);
    container.hidden = all.length === 0;
    container.classList.toggle('has-blockers', blocking.length > 0);
    $('receiptPrintIssuesTitle').textContent = blocking.length
      ? 'Ajuste estes campos antes de imprimir'
      : 'Confira estes pontos antes de imprimir';
  }

  function renderWarnings(warnings, blockers) {
    if ($('receiptExampleMode').checked) {
      const items = [...new Set([...issueStrings(blockers), ...issueStrings(warnings)])];
      fillIssueList(warningsBox, items);
      warningsBox.hidden = items.length === 0;
    } else {
      warningsBox.hidden = true;
      showPrintIssues(warnings, blockers);
    }
  }

  function resizeStage() {
    const nativeWidth = image.naturalWidth || 817;
    const available = Math.max(180, scroll.clientWidth - 2);
    const fit = Math.min(1, available / nativeWidth);
    stage.style.width = `${Math.round(nativeWidth * fit * zoomLevels[zoomIndex])}px`;
    $('zoomLabel').textContent = `${Math.round(zoomLevels[zoomIndex] * 100)}%`;
    $('zoomOut').disabled = zoomIndex === 0;
    $('zoomIn').disabled = zoomIndex === zoomLevels.length - 1;
  }

  function fieldLabel(field) {
    const base = names[field.key] || field.key || field.id;
    return field.line ? `${base} · linha ${field.line}` : base;
  }

  function getFields() { return Array.isArray(layout?.fields) ? layout.fields : []; }
  function selectedField() { return getFields().find(field => field.id === selectedId) || null; }

  function affectedByScope(field, selected, scope) {
    if (scope === 'general') return true;
    if (scope === 'field') return field.key === selected.key;
    if (scope === 'line') return field.line === selected.line;
    return field.id === selected.id;
  }

  function startMarkerDrag(event, field, marker) {
    if (!calibrating || !event.isPrimary) return;
    event.preventDefault();
    selectField(field.id, false);
    const target = scopeTarget();
    if (!target) return;
    marker.setPointerCapture(event.pointerId);
    dragState = {
      pointerId: event.pointerId, marker, field, target,
      scope: scopeValue(), startX: event.clientX, startY: event.clientY,
      offsetX: Number(target.x) || 0, offsetY: Number(target.y) || 0,
      moved: false
    };
  }

  function moveMarkerDrag(event) {
    const drag = dragState;
    if (!drag || drag.pointerId !== event.pointerId) return;
    const rect = stage.getBoundingClientRect();
    const deltaPxX = event.clientX - drag.startX;
    const deltaPxY = event.clientY - drag.startY;
    drag.moved ||= Math.abs(deltaPxX) + Math.abs(deltaPxY) > 3;
    drag.target.x = roundOffset(drag.offsetX + deltaPxX / rect.width * paperWidth);
    drag.target.y = roundOffset(drag.offsetY + deltaPxY / rect.height * paperHeight);
    $('receiptPrintIssues').hidden = true;
    const selected = drag.field;
    for (const button of overlay.querySelectorAll('.receipt-overlay-field')) {
      const field = getFields().find(item => item.id === button.dataset.fieldId);
      if (field && affectedByScope(field, selected, drag.scope)) {
        button.style.transform = `translate(${deltaPxX}px, ${deltaPxY}px)`;
      }
    }
    refreshCalibration();
    schedulePreview(180);
  }

  function endMarkerDrag(event, cancelled = false) {
    const drag = dragState;
    if (!drag || drag.pointerId !== event.pointerId) return;
    if (drag.marker.hasPointerCapture(event.pointerId)) drag.marker.releasePointerCapture(event.pointerId);
    if (cancelled) {
      drag.target.x = drag.offsetX;
      drag.target.y = drag.offsetY;
    }
    if (drag.moved) drag.marker.dataset.dragged = '1';
    dragState = null;
    refreshCalibration();
    if (cancelled) renderOverlay();
    schedulePreview(0);
  }

  function renderFieldSelect() {
    const fields = getFields();
    const oldId = selectedId;
    fieldSelect.replaceChildren(new Option('Selecione um campo', ''));
    for (const field of fields) fieldSelect.add(new Option(fieldLabel(field), field.id));
    selectedId = fields.some(field => field.id === oldId) ? oldId : (fields[0]?.id || null);
    fieldSelect.value = selectedId || '';
  }

  function renderOverlay() {
    if (dragState) return;
    overlay.replaceChildren();
    for (const field of getFields()) {
      const marker = document.createElement('button');
      marker.type = 'button';
      marker.className = 'receipt-overlay-field';
      marker.dataset.fieldId = field.id;
      marker.tabIndex = calibrating ? 0 : -1;
      marker.setAttribute('aria-label', `Ajustar ${fieldLabel(field)}`);
      marker.title = fieldLabel(field);
      const x = Number(field.x_mm) || 0;
      const y = Number(field.y_mm) || 0;
      const font = Number(field.font_size) || 8;
      const estimatedWidth = Math.max(6, Math.min(65, String(field.text || '').length * font * .18));
      const height = Math.max(4, font * .52);
      const align = field.align || 'left';
      const left = x - (align === 'right' ? estimatedWidth : align === 'center' ? estimatedWidth / 2 : 0);
      marker.style.left = `${left / paperWidth * 100}%`;
      marker.style.top = `${(y - height * .85) / paperHeight * 100}%`;
      marker.style.width = `${estimatedWidth / paperWidth * 100}%`;
      marker.style.height = `${height / paperHeight * 100}%`;
      marker.classList.toggle('is-selected', calibrating && field.id === selectedId);
      marker.addEventListener('click', () => {
        if (marker.dataset.dragged === '1') { delete marker.dataset.dragged; return; }
        selectField(field.id);
      });
      marker.addEventListener('pointerdown', event => startMarkerDrag(event, field, marker));
      marker.addEventListener('pointermove', moveMarkerDrag);
      marker.addEventListener('pointerup', event => endMarkerDrag(event));
      marker.addEventListener('pointercancel', event => endMarkerDrag(event, true));
      overlay.append(marker);
    }
  }

  function selectField(id, rerender = true) {
    const changed = selectedId !== id;
    selectedId = id;
    fieldSelect.value = id;
    refreshCalibration(changed);
    if (rerender) renderOverlay();
    else overlay.querySelectorAll('.receipt-overlay-field').forEach(button => {
      button.classList.toggle('is-selected', button.dataset.fieldId === id);
    });
  }

  function scopeValue() { return scopeInputs.find(input => input.checked)?.value || 'field'; }

  function scopeTarget() {
    const field = selectedField();
    if (!field) return null;
    const scope = scopeValue();
    if (scope === 'general') return {
      get x() { return Number(draftConfig.offset_x_mm) || 0; },
      set x(value) { draftConfig.offset_x_mm = value; },
      get y() { return Number(draftConfig.offset_y_mm) || 0; },
      set y(value) { draftConfig.offset_y_mm = value; }
    };
    if (scope === 'field') {
      draftConfig.campo_offsets_mm ||= {};
      draftConfig.campo_offsets_mm[field.key] ||= { x: 0, y: 0 };
      return draftConfig.campo_offsets_mm[field.key];
    }
    const index = Number(field.line) - 1;
    if (index < 0 || index >= 10) return null;
    if (scope === 'line') {
      draftConfig.linha_offsets_mm ||= [];
      draftConfig.linha_offsets_mm[index] ||= { x: 0, y: 0 };
      return draftConfig.linha_offsets_mm[index];
    }
    draftConfig.celula_offsets_mm ||= [];
    draftConfig.celula_offsets_mm[index] ||= {};
    draftConfig.celula_offsets_mm[index][field.key] ||= { x: 0, y: 0 };
    return draftConfig.celula_offsets_mm[index][field.key];
  }

  function refreshCalibration(selectionChanged = false) {
    const field = selectedField();
    const hasLine = field && Number.isInteger(Number(field.line)) && Number(field.line) >= 1;
    for (const input of scopeInputs) {
      input.disabled = !hasLine && (input.value === 'line' || input.value === 'cell');
    }
    if (scopeInputs.find(input => input.checked)?.disabled || (selectionChanged && !scopeTouchedByUser)) {
      const wanted = hasLine ? 'cell' : 'field';
      const choice = scopeInputs.find(input => input.value === wanted);
      if (choice) choice.checked = true;
    }
    $('scopeHelp').textContent = field ? scopeDescriptions[scopeValue()] : 'Preencha o recibo ou ative o exemplo para selecionar um campo.';
    const target = scopeTarget();
    $('calibrationX').value = target ? offsetText(target.x) : '';
    $('calibrationY').value = target ? offsetText(target.y) : '';
    $('calibrationX').disabled = !target;
    $('calibrationY').disabled = !target;
    document.querySelectorAll('[data-move]').forEach(button => { button.disabled = !target; });
    $('resetReceiptTarget').disabled = !target;
    $('selectedPosition').textContent = field
      ? `Posição final de ${fieldLabel(field)}: X ${offsetText(field.x_mm)} mm · Y ${offsetText(field.y_mm)} mm.`
      : '';
    updateDirtyState();
  }

  function updateDirtyState() {
    const dirty = JSON.stringify(draftConfig) !== JSON.stringify(savedConfig);
    $('saveReceiptCalibration').disabled = !dirty;
    $('discardReceiptCalibration').disabled = !dirty;
    return dirty;
  }

  function moveSelected(dx, dy) {
    const target = scopeTarget();
    if (!target) return;
    target.x = roundOffset((Number(target.x) || 0) + dx);
    target.y = roundOffset((Number(target.y) || 0) + dy);
    $('receiptPrintIssues').hidden = true;
    refreshCalibration();
    schedulePreview(80);
  }

  function commitOffsetInputs() {
    const target = scopeTarget();
    if (!target) return;
    const x = parseDecimal($('calibrationX').value);
    const y = parseDecimal($('calibrationY').value);
    if (!Number.isFinite(x) || !Number.isFinite(y) || Math.abs(x) > 100 || Math.abs(y) > 100) {
      setFeedback('Informe ajustes X e Y entre −100 e +100 mm.', true);
      refreshCalibration();
      return;
    }
    target.x = roundOffset(x);
    target.y = roundOffset(y);
    $('receiptPrintIssues').hidden = true;
    refreshCalibration();
    schedulePreview(80);
  }

  async function saveCalibration() {
    if (!updateDirtyState()) return;
    const button = $('saveReceiptCalibration');
    button.disabled = true;
    setFeedback('Salvando calibração…');
    try {
      const result = await jsonPost(endpoints.config, { config: draftConfig });
      savedConfig = copy(result.config || draftConfig);
      draftConfig = copy(savedConfig);
      refreshCalibration();
      schedulePreview(0);
      setFeedback('Calibração salva.');
    } catch (error) {
      setFeedback(error.message || 'Não foi possível salvar a calibração.', true);
      updateDirtyState();
    }
  }

  function discardCalibration() {
    if (!updateDirtyState()) return;
    if (!window.confirm('Descartar os ajustes de alinhamento que ainda não foram salvos?')) return;
    draftConfig = copy(savedConfig);
    refreshCalibration();
    schedulePreview(0);
    setFeedback('Ajustes descartados.');
  }

  async function downloadPdf(withModel = false) {
    if (!validateForPdf()) return;
    const button = withModel ? $('generateReceiptProof') : $('generateReceiptPdf');
    button.disabled = true;
    setFeedback('Conferindo posicionamento…');
    try {
      if (!withModel) {
        const check = await jsonPost(endpoints.preview, { payload: readPayload(), config: draftConfig });
        const blockers = issueStrings(check.layout?.print_blockers);
        const otherWarnings = issueStrings(check.layout?.warnings).filter(issue => !blockers.includes(issue));
        showPrintIssues(otherWarnings, blockers);
        if (blockers.length) {
          setFeedback('Há campos fora da área de impressão. Ajuste o alinhamento antes de gerar o PDF.', true);
          return;
        }
        if (otherWarnings.length && !window.confirm(`Confira antes de imprimir:\n\n${otherWarnings.join('\n')}\n\nDeseja gerar o PDF mesmo assim?`)) {
          setFeedback('Revise os avisos e gere o PDF quando estiver pronto.');
          return;
        }
      }
      setFeedback('Gerando PDF…');
      const response = await fetch(endpoints.pdf, {
        method: 'POST', credentials: 'same-origin',
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/pdf', 'X-CSRF-Token': csrf },
        body: JSON.stringify({ payload: readPayload(), config: draftConfig, with_model: withModel })
      });
      if (!response.ok) {
        const error = await response.json().catch(() => ({}));
        throw new Error(error.error || `Não foi possível gerar o PDF (${response.status}).`);
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = withModel ? 'recibo_conferencia_com_modelo.pdf' : 'recibo_dados.pdf';
      document.body.append(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 30000);
      setFeedback(withModel ? 'PDF de conferência gerado.' : 'PDF gerado. Imprima em tamanho real (100%).');
    } catch (error) {
      setFeedback(error.message || 'Não foi possível gerar o PDF.', true);
    } finally {
      button.disabled = false;
    }
  }

  function addItem() {
    if (visibleRows >= rows.length) return;
    rows[visibleRows].hidden = false;
    rows[visibleRows].querySelector('[data-field="quantidade"]').focus();
    visibleRows++;
    $('addReceiptItem').disabled = visibleRows >= rows.length;
  }

  function clearForm() {
    if (!window.confirm('Limpar todos os dados preenchidos no recibo?')) return;
    form.reset();
    $('receiptDate').value = new Date().toLocaleDateString('pt-BR');
    visibleRows = 1;
    rows.forEach((row, index) => { row.hidden = index >= visibleRows; });
    $('addReceiptItem').disabled = false;
    updateTotals();
    schedulePreview(0);
    $('receiptClient').focus();
  }

  function setExampleMode(enabled) {
    $('receiptExampleMode').checked = enabled;
    $('receiptExampleNote').hidden = !enabled;
    $('previewModeLabel').textContent = enabled ? 'Exemplo para ajuste' : 'Prévia dos dados preenchidos';
    selectedId = null;
    schedulePreview(0);
  }

  form.addEventListener('input', event => {
    if (event.target.matches('input')) {
      if (event.target.id === 'receiptDate') event.target.removeAttribute('aria-invalid');
      $('receiptPrintIssues').hidden = true;
      updateTotals();
      schedulePreview();
    }
  });
  form.addEventListener('submit', event => event.preventDefault());
  $('receiptDate').addEventListener('blur', event => {
    event.target.setAttribute('aria-invalid', String(!validDate(event.target.value)));
  });
  $('addReceiptItem').addEventListener('click', addItem);
  $('clearReceiptForm').addEventListener('click', clearForm);
  $('generateReceiptPdf').addEventListener('click', () => downloadPdf(false));
  $('generateReceiptProof').addEventListener('click', () => downloadPdf(true));

  $('toggleCalibration').addEventListener('click', () => {
    calibrating = !calibrating;
    calibrationPanel.hidden = !calibrating;
    stage.classList.toggle('is-calibrating', calibrating);
    $('toggleCalibration').setAttribute('aria-expanded', String(calibrating));
    $('toggleCalibration').textContent = calibrating ? 'Fechar ajustes' : 'Ajustar alinhamento';
    if (calibrating) {
      scopeTouchedByUser = false;
      const payload = readPayload();
      const hasContent = payload.cliente || payload.itens.some(item => item.descricao || item.quantidade || item.preco_unitario);
      if (!hasContent) setExampleMode(true);
      else schedulePreview(0);
      refreshCalibration();
      fieldSelect.focus();
    } else {
      setExampleMode(false);
    }
    renderOverlay();
  });
  fieldSelect.addEventListener('change', () => selectField(fieldSelect.value));
  scopeInputs.forEach(input => input.addEventListener('change', () => {
    scopeTouchedByUser = true;
    refreshCalibration();
  }));
  document.querySelectorAll('[data-move]').forEach(button => button.addEventListener('click', () => {
    const step = Number($('calibrationStep').value) || .5;
    const direction = button.dataset.move;
    moveSelected(direction === 'left' ? -step : direction === 'right' ? step : 0,
      direction === 'up' ? -step : direction === 'down' ? step : 0);
  }));
  $('calibrationX').addEventListener('change', commitOffsetInputs);
  $('calibrationY').addEventListener('change', commitOffsetInputs);
  $('saveReceiptCalibration').addEventListener('click', saveCalibration);
  $('discardReceiptCalibration').addEventListener('click', discardCalibration);
  $('resetReceiptTarget').addEventListener('click', () => {
    const target = scopeTarget();
    if (!target) return;
    target.x = 0;
    target.y = 0;
    $('receiptPrintIssues').hidden = true;
    refreshCalibration();
    schedulePreview(0);
  });
  $('receiptExampleMode').addEventListener('change', event => setExampleMode(event.target.checked));
  stage.addEventListener('keydown', event => {
    if (!calibrating || !selectedField() || !['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) return;
    event.preventDefault();
    const step = event.shiftKey ? 1 : (Number($('calibrationStep').value) || .5);
    moveSelected(event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0,
      event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0);
  });
  $('zoomOut').addEventListener('click', () => { zoomIndex = Math.max(0, zoomIndex - 1); resizeStage(); });
  $('zoomIn').addEventListener('click', () => { zoomIndex = Math.min(zoomLevels.length - 1, zoomIndex + 1); resizeStage(); });
  if ('ResizeObserver' in window) new ResizeObserver(resizeStage).observe(scroll);
  window.addEventListener('beforeunload', event => {
    if (!updateDirtyState()) return;
    event.preventDefault();
    event.returnValue = '';
  });

  $('receiptDate').value = new Date().toLocaleDateString('pt-BR');
  updateTotals();
  resizeStage();
  schedulePreview(0);
})();
