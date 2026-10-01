/* Operações de formulário e acompanhamento das filas de cobrança. */
(function (global) {
  'use strict';

  const JOB_KEY = 'edge_cobrancas_active_send_job';
  const STATUS_LABELS = {
    QUEUED: 'Na fila', RUNNING: 'Enviando cobranças',
    COMPLETED: 'Envio concluído', COMPLETED_WITH_ERRORS: 'Concluído com ocorrências',
    FAILED: 'Envio interrompido', INTERRUPTED: 'Envio interrompido', CANCELLED: 'Envio cancelado'
  };
  const TERMINAL = new Set(['COMPLETED', 'COMPLETED_WITH_ERRORS', 'FAILED', 'INTERRUPTED', 'CANCELLED']);

  function count(value) { return Math.max(0, Number(value) || 0); }
  function progressView(data) {
    const job = data.job || {}, counts = data.counts || {};
    const hasCounts = !!data.counts;
    const total = count(job.total_groups), processed = count(job.processed_groups);
    const sent = hasCounts ? count(counts.sent) : count(job.sent_groups);
    const tested = count(counts.test_sent), errors = hasCounts ? count(counts.error) : count(job.error_groups);
    const partial = count(counts.partial), uncertain = count(counts.uncertain), skipped = count(counts.skipped);
    const alerts = errors + partial + uncertain;
    const percent = Math.min(100, Math.max(0, Number(data.percent) || 0));
    const done = data.done === true || TERMINAL.has(job.status);
    const pieces = [`${sent} enviados`];
    if (tested) pieces.push(`${tested} testes`);
    if (errors) pieces.push(`${errors} erros`);
    if (skipped) pieces.push(`${skipped} ignorados`);
    if (partial) pieces.push(`${partial} parciais`);
    if (uncertain) pieces.push(`${uncertain} a conferir`);
    return {
      total, processed, sent, tested, errors, partial, uncertain, skipped, alerts, percent, done,
      status: STATUS_LABELS[job.status] || job.status || 'Preparando envio',
      current: (done ? job.message : job.current_label || job.message) || (done ? 'Processamento finalizado.' : 'Aguardando a próxima empresa...'),
      numbers: pieces.join(' · '),
      title: done ? (alerts ? 'Confira as ocorrências do envio' : ['FAILED', 'INTERRUPTED', 'CANCELLED'].includes(job.status) ? STATUS_LABELS[job.status] : tested && !sent ? 'Testes concluídos' : 'Cobranças processadas') : 'Processando cobranças',
      retryable: done && !!data.retryable && !!data.retry_url
    };
  }

  function normalizedText(value) {
    return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  }

  if (typeof module !== 'undefined' && module.exports) module.exports = { progressView, normalizedText };
  if (!global.document) return;

  const document = global.document;
  const csrf = document.querySelector('meta[name="csrf-token"]')?.content || '';
  let currentJob = null, jobTimer = null, starting = false, lastJobDone = false;
  let failures = 0, pollVersion = 0, polling = false, returnFocus = null;
  const byId = id => document.getElementById(id);
  const setText = (id, value) => { const element = byId(id); if (element) element.textContent = value; };
  const toggle = (id, visible) => byId(id)?.classList.toggle('hidden', !visible);

  function persistJob(id) {
    try { id ? global.localStorage.setItem(JOB_KEY, id) : global.localStorage.removeItem(JOB_KEY); } catch (_) { /* Navegadores podem restringir o armazenamento. */ }
  }

  function lockSendButtons(locked) {
    document.querySelectorAll('[data-billing-send]').forEach(button => {
      if (locked && !button.disabled) { button.dataset.billingLocked = '1'; button.disabled = true; }
      else if (!locked && button.dataset.billingLocked === '1') { button.disabled = false; delete button.dataset.billingLocked; }
    });
  }

  function showSendDock() { if (currentJob || lastJobDone || starting) toggle('sendProgressDock', true); }
  function hideSendProgress() {
    toggle('sendProgressModal', false);
    document.body.classList.remove('progress-dialog-open');
    showSendDock();
    if (returnFocus?.isConnected) returnFocus.focus();
  }
  function showSendProgress() {
    returnFocus = document.activeElement;
    toggle('sendProgressModal', true);
    toggle('sendProgressDock', false);
    document.body.classList.add('progress-dialog-open');
    byId('sendProgressClose')?.focus();
  }
  function dismissSendDock() {
    if (!lastJobDone) return;
    clearTimeout(jobTimer); pollVersion++;
    toggle('sendProgressDock', false);
    currentJob = null; persistJob(null); lastJobDone = false;
  }

  function resetProgress() {
    clearTimeout(jobTimer); pollVersion++; polling = false; failures = 0; lastJobDone = false;
    ['sendRetry', 'sendDockRetry', 'sendRefresh', 'sendDockDismiss', 'sendConnectionRetry', 'sendProgressAttention'].forEach(id => toggle(id, false));
    ['sendProgressSent', 'sendProgressErrors', 'sendProgressTest'].forEach(id => setText(id, '0'));
    setText('sendProgressProcessed', '0 / 0');
    setText('sendProgressPercent', '0%'); setText('sendDockPercent', '0%');
    setText('sendProgressTitle', 'Processando cobranças'); setText('sendDockTitle', 'Processando cobranças');
    setText('sendProgressCurrent', 'Aguardando a criação da fila...'); setText('sendDockCurrent', 'Preparando...');
    setText('sendDockNumbers', '0 enviados'); setText('sendProgressEvents', ''); setText('sendDockLabel', 'PREPARANDO ENVIO');
    ['sendProgressBar', 'sendDockBar'].forEach(id => { if (byId(id)) byId(id).style.width = '0%'; });
    byId('sendProgressTrack')?.setAttribute('aria-valuenow', '0');
  }

  async function requestJSON(url, options = {}) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 20000);
    try {
      const response = await global.fetch(url, { credentials: 'same-origin', cache: 'no-store', ...options, signal: controller.signal });
      let data = null;
      try { data = await response.json(); } catch (_) { /* Expiração de sessão pode devolver uma página HTML. */ }
      if (!response.ok || response.redirected || !data?.ok) {
        const error = new Error(data?.error || (response.status === 404 ? 'Esta fila não está mais disponível.' : response.status === 401 || response.status === 403 || response.redirected ? 'Sua sessão expirou. Atualize a página para entrar novamente.' : `Não foi possível obter uma resposta válida (HTTP ${response.status}).`));
        error.status = response.redirected ? 401 : response.status;
        throw error;
      }
      return data;
    } catch (error) {
      if (error.name === 'AbortError') throw new Error('O servidor demorou a responder.');
      throw error;
    } finally { clearTimeout(timeout); }
  }

  async function startBillingSend(url, confirmText) {
    if (starting) return;
    if (currentJob && !lastJobDone) { showSendProgress(); return; }
    if (confirmText && !global.confirm(confirmText)) return;
    starting = true; currentJob = null; persistJob(null); lockSendButtons(true); resetProgress();
    showSendProgress(); setText('sendProgressState', 'Criando fila de envio...');
    const body = new FormData(); body.append('_csrf_token', csrf);
    try {
      const data = await requestJSON(url, { method: 'POST', body, headers: { 'X-CSRF-Token': csrf } });
      if (!data.job_id) throw new Error('O servidor não informou a fila criada.');
      currentJob = String(data.job_id); persistJob(currentJob);
      await pollJob();
    } catch (error) {
      lastJobDone = true;
      const refused = [400, 401, 403, 404, 409, 422].includes(error.status);
      setText('sendProgressState', refused ? 'Envio não iniciado' : 'Início do envio não confirmado');
      const detail = refused ? error.message : `${error.message} Atualize a competência e confira se há uma fila em andamento antes de tentar novamente.`;
      setText('sendProgressCurrent', detail); setText('sendDockCurrent', detail);
      setText('sendDockLabel', 'CONFIRA O ENVIO'); setText('sendDockTitle', 'Não foi possível confirmar o início');
      toggle('sendRefresh', true); toggle('sendDockDismiss', true); lockSendButtons(false);
    } finally { starting = false; }
  }

  async function retrySendErrors(url) {
    if (!url) return;
    await startBillingSend(url, 'Reenviar somente as empresas que falharam? Envio parcial ou sem confirmação exige conferência no histórico.');
  }

  async function reconcileEvent(event, action, buttons) {
    const partial = ['PARTIAL', 'TEST_PARTIAL'].includes(event.status);
    const prompt = action === 'delivered'
      ? partial ? 'Você conferiu o Gmail e confirma que a cobrança foi recebida pela empresa principal?' : 'Você conferiu os e-mails enviados no Gmail e confirma que esta cobrança foi enviada?'
      : partial ? 'Você conferiu os destinatários e quer liberar o reenvio completo? Os destinatários que já receberam poderão receber outra cópia.' : 'Você conferiu o Gmail e confirma que esta cobrança não foi enviada? Essa confirmação libera uma nova tentativa.';
    if (!global.confirm(prompt)) return;
    const body = new FormData(); body.append('_csrf_token', csrf); body.append('action', action);
    if (action === 'delivered' && event.mode_unknown) {
      const mode = String(global.prompt('Esta fila antiga não registrou o modo de envio. Após conferir o Gmail, digite TESTE se o e-mail foi para o endereço de teste, ou REAL se foi para a empresa.') || '').trim().toUpperCase();
      if (!['TESTE', 'REAL'].includes(mode)) return;
      body.append('original_test_mode', mode === 'TESTE' ? '1' : '0');
    }
    buttons.forEach(button => button.disabled = true);
    try {
      await requestJSON(event.reconciliation_url, { method: 'POST', body, headers: { 'X-CSRF-Token': csrf } });
      await pollJob();
    } catch (error) {
      setText('sendProgressAttention', `${error.message || 'Não foi possível salvar a conferência.'} Confira o resultado atualizado antes de tentar novamente.`);
      toggle('sendProgressAttention', true);
      buttons.forEach(button => button.disabled = false);
    }
  }

  function trackJob(id) {
    if (starting) return;
    if (currentJob && !lastJobDone && currentJob !== id) { showSendProgress(); return; }
    resetProgress(); currentJob = id; persistJob(id); showSendProgress(); pollJob();
  }

  function renderJobProgress(data) {
    const view = progressView(data);
    lastJobDone = view.done;
    setText('sendProgressState', view.status); setText('sendProgressTitle', view.title);
    setText('sendProgressPercent', `${view.percent}%`); setText('sendDockPercent', `${view.percent}%`);
    setText('sendProgressProcessed', `${view.processed} / ${view.total}`);
    setText('sendProgressSent', view.sent); setText('sendProgressTest', view.tested); setText('sendProgressErrors', view.errors);
    setText('sendProgressCurrent', view.current); setText('sendDockCurrent', view.current);
    setText('sendDockNumbers', view.numbers); setText('sendDockTitle', view.title);
    setText('sendDockLabel', view.done ? view.alerts ? 'CONCLUÍDO COM OCORRÊNCIAS' : 'ENVIO CONCLUÍDO' : 'ENVIO EM ANDAMENTO');
    ['sendProgressBar', 'sendDockBar'].forEach(id => { if (byId(id)) byId(id).style.width = `${view.percent}%`; });
    byId('sendProgressTrack')?.setAttribute('aria-valuenow', String(view.percent));
    byId('sendProgressTrack')?.setAttribute('aria-valuetext', `${view.processed} de ${view.total} cobranças processadas`);
    const events = byId('sendProgressEvents');
    if (events) {
      events.replaceChildren();
      const eventList = [...(data.reconciliation_events || []), ...(data.events || [])];
      const seen = new Set();
      eventList.forEach(event => {
        if (event.group_id && seen.has(event.group_id)) return;
        if (event.group_id) seen.add(event.group_id);
        const row = document.createElement('div');
        row.className = `progress-event ${['success', 'warning', 'error'].includes(event.level) ? event.level : 'info'}`;
        const message = document.createElement('span'); message.textContent = event.message; row.appendChild(message);
        if (event.reconciliation_url) {
          const actions = document.createElement('div'); actions.className = 'row-actions';
          const delivered = document.createElement('button'), notDelivered = document.createElement('button');
          delivered.type = notDelivered.type = 'button';
          delivered.className = 'btn btn-sm btn-secondary'; notDelivered.className = 'btn btn-sm btn-warning';
          const partial = ['PARTIAL', 'TEST_PARTIAL'].includes(event.status);
          delivered.textContent = partial ? 'Confirmar cobrança recebida' : 'Confirmar envio no Gmail';
          notDelivered.textContent = partial ? 'Liberar reenvio após conferência' : 'Confirmar que não enviou';
          delivered.onclick = () => reconcileEvent(event, 'delivered', [delivered, notDelivered]);
          notDelivered.onclick = () => reconcileEvent(event, 'not_delivered', [delivered, notDelivered]);
          actions.appendChild(delivered);
          if (!partial || !event.primary_accepted) actions.appendChild(notDelivered);
          if (partial && event.primary_accepted) {
            const explanation = document.createElement('small');
            explanation.textContent = 'A empresa principal já recebeu. Confira as cópias recusadas; se precisar de outro envio, use Reenviar cobrança na empresa responsável.';
            row.appendChild(explanation);
          }
          row.appendChild(actions);
        }
        events.appendChild(row);
      });
    }
    const attention = [];
    if (view.tested) attention.push(`${view.tested} cobrança(s) enviada(s) ao e-mail de teste.`);
    if (view.skipped) attention.push(`${view.skipped} cobrança(s) ignorada(s) após nova verificação de elegibilidade.`);
    if (view.partial || view.uncertain) attention.push(`${view.partial + view.uncertain} envio(s) precisa(m) de conferência no histórico antes de qualquer reenvio.`);
    setText('sendProgressAttention', attention.join(' ')); toggle('sendProgressAttention', !!attention.length);
    toggle('sendRefresh', view.done); toggle('sendDockDismiss', view.done); toggle('sendConnectionRetry', false);
    ['sendRetry', 'sendDockRetry'].forEach(id => {
      toggle(id, view.retryable);
      if (byId(id)) byId(id).onclick = () => retrySendErrors(data.retry_url);
    });
    lockSendButtons(!view.done);
    // A conclusão fica disponível ao navegar até o usuário fechar o cartão.
    if (currentJob) persistJob(currentJob);
  }

  async function pollJob() {
    if (!currentJob || polling) return;
    clearTimeout(jobTimer);
    const job = currentJob, version = pollVersion;
    polling = true;
    try {
      const template = document.body.dataset.jobStatusUrl;
      const data = await requestJSON(template.replace('JOBID', encodeURIComponent(job)));
      if (job !== currentJob || version !== pollVersion) return;
      failures = 0; renderJobProgress(data);
      if (lastJobDone) {
        if (byId('sendProgressModal').classList.contains('hidden')) showSendDock();
      } else jobTimer = setTimeout(pollJob, document.hidden ? 5000 : 1500);
    } catch (error) {
      if (job !== currentJob || version !== pollVersion) return;
      if ([401, 403, 404].includes(error.status)) {
        lastJobDone = true; persistJob(null); lockSendButtons(false);
        setText('sendProgressState', 'Acompanhamento indisponível');
        setText('sendProgressCurrent', error.message); setText('sendDockCurrent', error.message);
        setText('sendDockLabel', 'ACOMPANHAMENTO INDISPONÍVEL');
        toggle('sendDockDismiss', true); toggle('sendRefresh', true);
      } else {
        failures++;
        setText('sendProgressState', 'Conexão interrompida · tentando reconectar');
        setText('sendDockLabel', 'RECONECTANDO');
        const message = `${error.message || 'Sem conexão com o servidor.'} O último progresso foi mantido; o envio pode continuar no servidor.`;
        setText('sendProgressCurrent', message); setText('sendDockCurrent', message);
        toggle('sendConnectionRetry', true);
        jobTimer = setTimeout(pollJob, Math.min(15000, 1500 * Math.pow(2, failures)));
      }
    } finally { if (version === pollVersion) polling = false; }
  }

  function confirmPermanent(form, label) {
    const value = global.prompt(`Excluir definitivamente ${label}?\n\nO sistema criará backup quando aplicável.\nDigite EXCLUIR para confirmar:`);
    if (String(value || '').trim().toUpperCase() !== 'EXCLUIR') return false;
    let input = form.querySelector('input[name="confirm_text"]');
    if (!input) { input = document.createElement('input'); input.type = 'hidden'; input.name = 'confirm_text'; form.appendChild(input); }
    input.value = 'EXCLUIR'; return true;
  }

  function installForms() {
    document.querySelectorAll('.form-group').forEach((group, index) => {
      const label = group.querySelector('label'), field = group.querySelector('input:not([type="hidden"]),select,textarea');
      if (label && field && !label.htmlFor) { field.id ||= `field-${index}`; label.htmlFor = field.id; }
    });
    document.addEventListener('submit', event => {
      const form = event.target;
      if (!(form instanceof HTMLFormElement) || event.defaultPrevented || (form.method || 'get').toLowerCase() !== 'post' || form.dataset.async === '1') return;
      if (form.dataset.submitting === '1') { event.preventDefault(); return; }
      if (!form.checkValidity() && !event.submitter?.formNoValidate) return;
      if (form.dataset.requireSelection && !form.querySelector('input[name="company_ids"]:checked')) {
        event.preventDefault(); setText('selectionStatus', 'Selecione ao menos uma empresa para registrar a baixa.'); return;
      }
      if (!form.querySelector('input[name="_csrf_token"]')) {
        const input = document.createElement('input'); input.type = 'hidden'; input.name = '_csrf_token'; input.value = csrf; form.appendChild(input);
      }
      form.dataset.submitting = '1';
      setTimeout(() => {
        if (event.defaultPrevented) { delete form.dataset.submitting; return; }
        form.querySelectorAll('button:not([type="button"]):not([type="reset"]),input[type="submit"]').forEach(button => {
          if (!button.disabled) { button.dataset.operationDisabled = '1'; button.disabled = true; }
        });
        toggle('operationOverlay', true);
      }, 0);
    });
    global.addEventListener('pageshow', () => {
      toggle('operationOverlay', false);
      document.querySelectorAll('form[data-submitting="1"]').forEach(form => {
        delete form.dataset.submitting;
        form.querySelectorAll('[data-operation-disabled="1"]').forEach(button => { button.disabled = false; delete button.dataset.operationDisabled; });
      });
    });
  }

  function installTableFilters() {
    document.querySelectorAll('[data-table-search]').forEach(input => {
      const table = byId(input.dataset.tableSearch), status = byId(input.dataset.searchStatus);
      if (!table) return;
      input.addEventListener('input', () => {
        const query = normalizedText(input.value), rows = Array.from(table.querySelectorAll('tbody tr[data-search-row]'));
        let shown = 0;
        rows.forEach(row => { row.hidden = !normalizedText(row.textContent).includes(query); if (!row.hidden) shown++; });
        const units = input.dataset.searchUnits || 'empresa(s)';
        if (status) status.textContent = query ? `${shown} de ${rows.length} ${units} exibido(s). ${input.dataset.searchNote ?? 'O filtro não altera a fila de envio.'}` : `${rows.length} ${units}.`;
        table.dispatchEvent(new Event('rowsfiltered'));
      });
    });
  }

  Object.assign(global, { startBillingSend, retrySendErrors, hideSendProgress, showSendProgress, dismissSendDock, confirmPermanent, pollJob });
  document.querySelectorAll('[data-track-job]').forEach(button => button.addEventListener('click', () => trackJob(button.dataset.trackJob)));
  document.addEventListener('keydown', event => {
    const modal = byId('sendProgressModal');
    if (!modal || modal.classList.contains('hidden')) return;
    if (event.key === 'Escape') { hideSendProgress(); return; }
    if (event.key !== 'Tab') return;
    const focusable = Array.from(modal.querySelectorAll('button,a[href]')).filter(element => !element.disabled && !element.closest('.hidden'));
    const first = focusable[0], last = focusable[focusable.length - 1];
    if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus(); }
    else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus(); }
  });
  global.addEventListener('online', () => { if (currentJob && !lastJobDone) pollJob(); });
  document.addEventListener('visibilitychange', () => { if (!document.hidden && currentJob && !lastJobDone) pollJob(); });
  installForms(); installTableFilters();
  currentJob = document.body.dataset.activeSendJob || null;
  if (!currentJob) { try { currentJob = global.localStorage.getItem(JOB_KEY); } catch (_) { /* Armazenamento indisponível. */ } }
  if (currentJob) { persistJob(currentJob); showSendDock(); lockSendButtons(true); pollJob(); }
})(typeof window !== 'undefined' ? window : globalThis);
