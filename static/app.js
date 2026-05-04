const relFilesPicker = document.getElementById('relFilesPicker');
const relFolderPicker = document.getElementById('relFolderPicker');
const hiddenRelFiles = document.getElementById('hiddenRelFiles');
const fileListBody = document.getElementById('fileListBody');
const fileCount = document.getElementById('fileCount');
const clearAllBtn = document.getElementById('clearAllBtn');
const uploadForm = document.getElementById('uploadForm');
const selectionSummary = document.getElementById('selectionSummary');
const loadingOverlay = document.getElementById('loadingOverlay');
const baseFileInput = document.getElementById('baseFile');
const baseSheetSelect = document.getElementById('baseSheetSelect');

let selectedFiles = [];

function csrfHeader() {
  const meta = document.querySelector('meta[name="csrf-token"]');
  const token = meta ? meta.getAttribute('content') : '';
  return token ? { 'X-CSRF-Token': token } : {};
}


function hideLoading() {
  if (!loadingOverlay) return;
  loadingOverlay.classList.remove('is-visible');
  loadingOverlay.classList.add('is-hidden');
}

function showLoading() {
  if (!loadingOverlay) return;
  loadingOverlay.classList.remove('is-hidden');
  loadingOverlay.classList.add('is-visible');
}

function formatBytes(bytes) {
  if (!bytes) return '0 KB';
  const units = ['B', 'KB', 'MB', 'GB'];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / Math.pow(1024, index)).toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}

function setSheetOptions(options) {
  if (!baseSheetSelect) return;
  baseSheetSelect.innerHTML = '';
  options.forEach(item => {
    const option = document.createElement('option');
    option.value = item.value;
    option.textContent = item.label;
    baseSheetSelect.appendChild(option);
  });
}

async function loadBaseSheets() {
  const file = baseFileInput?.files?.[0];
  if (!file) {
    setSheetOptions([{ value: '', label: 'Carregue a planilha base primeiro' }]);
    return;
  }

  const formData = new FormData();
  formData.append('base_file', file);
  setSheetOptions([{ value: '', label: 'Lendo abas da planilha...' }]);

  try {
    const response = await fetch('/esocial/abas-base', { method: 'POST', body: formData, headers: csrfHeader() });
    const data = await response.json();

    if (!response.ok || !data.ok) {
      setSheetOptions([{ value: '', label: data.error || 'Não foi possível ler as abas' }]);
      return;
    }

    const sheets = data.sheets || [];
    if (sheets.length <= 1) {
      setSheetOptions([{ value: sheets[0] || '', label: sheets[0] || 'Planilha principal' }]);
      return;
    }

    setSheetOptions([
      { value: '', label: 'Selecione a aba do mês' },
      ...sheets.map(sheet => ({ value: sheet, label: sheet }))
    ]);
  } catch {
    setSheetOptions([{ value: '', label: 'Erro ao carregar abas' }]);
  }
}

function isValidRelFile(file) {
  return ['.xls', '.xlsx', '.html', '.htm'].some(ext => file.name.toLowerCase().endsWith(ext));
}

function buildKey(file) {
  return `${file.name}__${file.size}__${file.lastModified}`;
}

function syncHiddenInput() {
  const dt = new DataTransfer();
  selectedFiles.forEach(item => dt.items.add(item.file));
  hiddenRelFiles.files = dt.files;
}

function updateSummary() {
  const total = selectedFiles.length;
  const size = selectedFiles.reduce((acc, item) => acc + item.file.size, 0);
  fileCount.textContent = `${total} arquivo(s)`;
  selectionSummary.textContent = total === 0
    ? 'Nenhum arquivo selecionado.'
    : `${total} arquivo(s) prontos para processamento. Tamanho total: ${formatBytes(size)}.`;
}

function renderFileList() {
  fileListBody.innerHTML = '';

  if (selectedFiles.length === 0) {
    fileListBody.innerHTML = '<tr class="empty-row"><td colspan="4">Nenhum RELFUNCGERAL selecionado.</td></tr>';
    updateSummary();
    syncHiddenInput();
    return;
  }

  selectedFiles.forEach((item, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${item.file.name}</td><td>${formatBytes(item.file.size)}</td><td>${item.source}</td><td><button type="button" class="btn btn--danger remove-btn" data-index="${index}">Retirar</button></td>`;
    fileListBody.appendChild(tr);
  });

  document.querySelectorAll('.remove-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      selectedFiles.splice(Number(btn.dataset.index), 1);
      renderFileList();
    });
  });

  updateSummary();
  syncHiddenInput();
}

function addFiles(fileList, sourceLabel) {
  const existingKeys = new Set(selectedFiles.map(item => buildKey(item.file)));
  let ignored = 0;

  Array.from(fileList).forEach(file => {
    if (!isValidRelFile(file)) {
      ignored += 1;
      return;
    }

    const key = buildKey(file);
    if (existingKeys.has(key)) return;

    const source = file.webkitRelativePath || sourceLabel;
    selectedFiles.push({ file, source });
    existingKeys.add(key);
  });

  renderFileList();

  if (ignored > 0) {
    alert(`${ignored} arquivo(s) foram ignorados porque não são .xls, .xlsx, .html ou .htm.`);
  }
}

if (baseFileInput) {
  baseFileInput.addEventListener('change', loadBaseSheets);
}

if (relFilesPicker) {
  relFilesPicker.addEventListener('change', e => {
    addFiles(e.target.files, 'Seleção individual');
    e.target.value = '';
  });
}

if (relFolderPicker) {
  relFolderPicker.addEventListener('change', e => {
    addFiles(e.target.files, 'Pasta');
    e.target.value = '';
  });
}

if (clearAllBtn) {
  clearAllBtn.addEventListener('click', () => {
    selectedFiles = [];
    renderFileList();
    hideLoading();
  });
}

if (uploadForm) {
  uploadForm.addEventListener('submit', e => {
    if (selectedFiles.length === 0) {
      e.preventDefault();
      hideLoading();
      alert('Selecione pelo menos um arquivo RELFUNCGERAL.');
      return;
    }

    const optionsCount = baseSheetSelect.options.length;
    const selectedSheet = baseSheetSelect.value;
    if (optionsCount > 1 && !selectedSheet) {
      e.preventDefault();
      alert('Selecione a aba da planilha base.');
      return;
    }

    syncHiddenInput();
    showLoading();
  });
}

window.addEventListener('load', hideLoading);
window.addEventListener('pageshow', hideLoading);

if (fileListBody) {
  renderFileList();
}

hideLoading();

(function () {
  const meta = document.querySelector('meta[name="csrf-token"]');
  if (!meta) return;
  const token = meta.getAttribute('content');
  if (!token) return;
  document.querySelectorAll('form[method="post"], form[method="POST"]').forEach((form) => {
    if (!form.querySelector('input[name="_csrf_token"]')) {
      const input = document.createElement('input');
      input.type = 'hidden';
      input.name = '_csrf_token';
      input.value = token;
      form.appendChild(input);
    }
  });
})();


(function () {
  function setSubmitProcessing(submit) {
    if (!submit || submit.dataset.keepEnabled === '1') return;
    if (!submit.dataset.originalText) {
      submit.dataset.originalText = submit.textContent || submit.value || 'Processar';
    }
    submit.dataset.processing = '1';
    submit.disabled = true;
    if (submit.tagName === 'BUTTON') submit.textContent = 'Processando...';
    else submit.value = 'Processando...';
  }

  function resetProcessingButtons() {
    document.querySelectorAll('[data-processing="1"]').forEach((submit) => {
      submit.disabled = false;
      const original = submit.dataset.originalText || 'Processar';
      if (submit.tagName === 'BUTTON') submit.textContent = original;
      else submit.value = original;
      delete submit.dataset.processing;
    });
  }

  document.querySelectorAll('form').forEach((form) => {
    form.addEventListener('submit', () => {
      const submit = form.querySelector('button[type="submit"], input[type="submit"]');
      setTimeout(() => setSubmitProcessing(submit), 0);

      // Quando o POST retorna um arquivo para download, a página não recarrega.
      // Sem esse reset, o botão fica preso em "Processando..." e impede novo envio.
      const resetAfter = Number(form.dataset.resetSubmitAfter || 2500);
      if (resetAfter > 0) {
        window.setTimeout(resetProcessingButtons, resetAfter);
      }
    });
  });

  window.addEventListener('pageshow', resetProcessingButtons);
  window.addEventListener('focus', resetProcessingButtons);
})();

// Refinamento UX: mostra o nome/tamanho dos arquivos selecionados abaixo de inputs comuns.
(function () {
  function humanSize(bytes) {
    if (!bytes) return '0 KB';
    const units = ['B', 'KB', 'MB', 'GB'];
    const idx = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    return `${(bytes / Math.pow(1024, idx)).toFixed(idx === 0 ? 0 : 1)} ${units[idx]}`;
  }

  document.querySelectorAll('input[type="file"]').forEach((input) => {
    if (input.hidden || ['relFilesPicker', 'relFolderPicker', 'hiddenRelFiles'].includes(input.id)) return;

    let preview = input.parentElement.querySelector('.file-preview');
    if (!preview) {
      preview = document.createElement('small');
      preview.className = 'file-preview';
      preview.textContent = 'Nenhum arquivo selecionado.';
      input.insertAdjacentElement('afterend', preview);
    }

    input.addEventListener('change', () => {
      const files = Array.from(input.files || []);
      if (!files.length) {
        preview.textContent = 'Nenhum arquivo selecionado.';
        return;
      }
      const total = files.reduce((sum, file) => sum + file.size, 0);
      const firstNames = files.slice(0, 2).map(file => file.name).join(', ');
      const extra = files.length > 2 ? ` +${files.length - 2} arquivo(s)` : '';
      preview.textContent = `${firstNames}${extra} • ${humanSize(total)}`;
    });
  });
})();


// Acabamento premium: normaliza navegação em ícone + label e adiciona cabeçalho sem campo de busca.
(function () {
  const navLinks = document.querySelectorAll('.nav-link');
  navLinks.forEach((link) => {
    if (link.querySelector('.nav-icon')) return;
    const text = (link.textContent || '').trim();
    const match = text.match(/^([^\wÀ-ÿ]+)\s*(.*)$/u);
    const icon = match ? match[1].trim() : '•';
    const label = match ? match[2].trim() : text;
    link.innerHTML = `<span class="nav-icon" aria-hidden="true">${icon}</span><span class="nav-label">${label}</span>`;
    if (!link.getAttribute('aria-label')) link.setAttribute('aria-label', label || text);
    if (!link.getAttribute('title')) link.setAttribute('title', label || text);
  });

  const main = document.querySelector('.main-area');
  if (!main || main.querySelector('.dashboard-header')) return;
  const pageTitle = document.querySelector('.page-title h1')?.textContent?.trim() || 'Painel';
  const today = new Intl.DateTimeFormat('pt-BR', { weekday: 'short', day: '2-digit', month: 'long', year: 'numeric' }).format(new Date());
  const header = document.createElement('div');
  header.className = 'dashboard-header';
  header.innerHTML = `
    <div class="dashboard-welcome">
      <h2>Bem-vindo ao sistema EDGE</h2>
      <p>${today} • ${pageTitle}</p>
    </div>
    <div class="dashboard-tools">
      <div class="dashboard-avatar" aria-label="Usuário">E</div>
    </div>`;
  main.insertBefore(header, main.firstElementChild);
})();
