(() => {
  const sheetFile = document.getElementById('sheetFile');
  const sheetDrop = document.getElementById('sheetDrop');
  const sheetLabel = document.getElementById('sheetLabel');
  const sheetMeta = document.getElementById('sheetMeta');
  const sheetSelect = document.getElementById('sheetSelect');
  const employeeCount = document.getElementById('employeeCount');
  const examCount = document.getElementById('examCount');
  const pdfInput = document.getElementById('pdfFiles');
  const pdfDrop = document.getElementById('pdfDrop');
  const pdfCount = document.getElementById('pdfCount');
  const pdfList = document.getElementById('pdfList');
  const processBtn = document.getElementById('processBtn');
  const folderName = document.getElementById('folderName');
  const progressCard = document.getElementById('progressCard');
  const resultCard = document.getElementById('resultCard');
  const activeBar = document.getElementById('activeExtractionBar');
  const activeFolder = document.getElementById('activeFolderName');
  const clearBtn = document.getElementById('clearBtn');
  const archiveMonth = document.getElementById('archiveMonth');
  const archiveYear = document.getElementById('archiveYear');
  const archiveSaveBtn = document.getElementById('archiveSaveBtn');
  const archiveSaveInfo = document.getElementById('archiveSaveInfo');

  let listToken = '';
  let pdfFiles = [];
  let pollTimer = null;
  let currentJobId = '';
  const ACTIVE_JOB_KEY = 'edgeActiveJobId';

  function rememberJob(jobId) {
    currentJobId = jobId || '';
    try {
      if (currentJobId) localStorage.setItem(ACTIVE_JOB_KEY, currentJobId);
      else localStorage.removeItem(ACTIVE_JOB_KEY);
    } catch (_) {}
  }

  function setDrag(el) {
    ['dragenter', 'dragover'].forEach(ev => el.addEventListener(ev, e => {
      e.preventDefault(); el.classList.add('dragging');
    }));
    ['dragleave', 'drop'].forEach(ev => el.addEventListener(ev, e => {
      e.preventDefault(); el.classList.remove('dragging');
    }));
  }

  setDrag(sheetDrop); setDrag(pdfDrop);
  sheetDrop.addEventListener('drop', e => { const f = e.dataTransfer.files?.[0]; if (f) uploadSheet(f); });
  pdfDrop.addEventListener('drop', e => addPdfs([...e.dataTransfer.files]));
  sheetFile.addEventListener('change', () => sheetFile.files[0] && uploadSheet(sheetFile.files[0]));
  pdfInput.addEventListener('change', () => addPdfs([...pdfInput.files]));

  async function uploadSheet(file) {
    const fd = new FormData(); fd.append('file', file);
    sheetLabel.textContent = 'Lendo planilha...';
    try {
      const data = await edgeJson('/api/planilha', { method: 'POST', body: fd });
      listToken = data.token;
      sheetLabel.textContent = data.filename;
      sheetSelect.innerHTML = data.sheets.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
      sheetSelect.value = data.selected_sheet;
      employeeCount.textContent = data.employees;
      examCount.textContent = data.exams;
      sheetMeta.classList.remove('hidden');
      edgeToast('Planilha carregada.');
    } catch (e) {
      listToken = '';
      sheetLabel.textContent = 'Selecionar planilha';
      sheetMeta.classList.add('hidden');
      edgeToast(e.message, 'error');
    }
  }

  sheetSelect.addEventListener('change', async () => {
    if (!listToken) return;
    try {
      const data = await edgeJson(`/api/planilha/${listToken}/resumo?sheet=${encodeURIComponent(sheetSelect.value)}`);
      employeeCount.textContent = data.employees;
      examCount.textContent = data.exams;
    } catch (e) { edgeToast(e.message, 'error'); }
  });

  function addPdfs(files) {
    const existing = new Set(pdfFiles.map(f => `${f.name}|${f.size}|${f.lastModified}`));
    files.filter(f => f.name.toLowerCase().endsWith('.pdf')).forEach(f => {
      const key = `${f.name}|${f.size}|${f.lastModified}`;
      if (!existing.has(key)) { pdfFiles.push(f); existing.add(key); }
    });
    renderPdfs();
  }

  function renderPdfs() {
    pdfCount.textContent = pdfFiles.length
      ? `${pdfFiles.length} arquivo${pdfFiles.length > 1 ? 's' : ''} selecionado${pdfFiles.length > 1 ? 's' : ''}`
      : 'Nenhum arquivo selecionado';
    pdfList.innerHTML = pdfFiles.slice(0, 4).map((f, i) =>
      `<div><span>${escapeHtml(f.name)}</span><button data-i="${i}" title="Remover">×</button></div>`
    ).join('') + (pdfFiles.length > 4 ? `<small>+ ${pdfFiles.length - 4} arquivo(s)</small>` : '');
    pdfList.querySelectorAll('button').forEach(b => b.onclick = () => {
      pdfFiles.splice(+b.dataset.i, 1); renderPdfs();
    });
  }

  function initArchivePeriod() {
    if (!archiveMonth || !archiveYear) return;
    const monthNames = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
    archiveMonth.innerHTML = monthNames.map((name, i) => `<option value="${i+1}">${String(i+1).padStart(2,'0')} - ${name}</option>`).join('');
    const now = new Date();
    const y = now.getFullYear();
    archiveYear.innerHTML = Array.from({length: 13}, (_, i) => y - 10 + i).map(v => `<option value="${v}">${v}</option>`).join('');
    archiveMonth.value = String(now.getMonth() + 1);
    archiveYear.value = String(y);
  }

  function updateArchiveInfo(data) {
    if (!archiveSaveInfo || !archiveSaveBtn) return;
    const eligible = Number(data.archive_eligible || 0);
    const saved = Number(data.archive_saved || 0);
    if (!eligible) {
      archiveSaveInfo.textContent = 'Ainda não há arquivos aprovados para arquivar.';
      archiveSaveBtn.disabled = true;
      return;
    }
    archiveSaveBtn.disabled = false;
    archiveSaveInfo.textContent = saved
      ? `${saved} documento(s) deste lote já estão no Arquivo de Exames. Há ${eligible} arquivo(s) aprovados na extração atual.`
      : `${eligible} arquivo(s) aprovados estão prontos para serem salvos no Arquivo de Exames.`;
  }

  archiveSaveBtn?.addEventListener('click', async () => {
    if (!currentJobId) return edgeToast('Nenhuma extração ativa para arquivar.', 'error');
    const month = Number(archiveMonth.value);
    const year = Number(archiveYear.value);
    if (!month || !year) return edgeToast('Escolha o mês e o ano.', 'error');
    const label = `${String(month).padStart(2,'0')}/${year}`;
    if (!window.confirm(`Salvar os arquivos aprovados deste lote no Arquivo de Exames da competência ${label}?`)) return;
    archiveSaveBtn.disabled = true;
    archiveSaveBtn.textContent = 'Salvando no arquivo...';
    try {
      const data = await edgeJson(`/api/jobs/${currentJobId}/archive`, {
        method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify({month, year})
      });
      const parts = [];
      if (data.added_count) parts.push(`${data.added_count} novo(s) documento(s) salvo(s)`);
      if (data.skipped_count) parts.push(`${data.skipped_count} já estavam arquivados`);
      if (data.error_count) parts.push(`${data.error_count} com erro`);
      archiveSaveInfo.textContent = `${data.archive_saved} documento(s) deste lote arquivados na competência ${data.competency}.`;
      edgeToast(parts.join(' · ') || 'Arquivo atualizado.');
    } catch (e) {
      edgeToast(e.message, 'error');
    } finally {
      archiveSaveBtn.disabled = false;
      archiveSaveBtn.textContent = 'Salvar no Arquivo de Exames';
    }
  });

  processBtn.addEventListener('click', async () => {
    if (currentJobId) return edgeToast('Existe uma extração ativa. Clique em Limpar extração antes de iniciar um novo lote.', 'error');
    if (!listToken) return edgeToast('Carregue a lista de funcionários.', 'error');
    if (!pdfFiles.length) return edgeToast('Adicione pelo menos um PDF.', 'error');
    const name = folderName.value.trim();
    if (!name) return edgeToast('Informe o nome da pasta final.', 'error');

    processBtn.disabled = true;
    processBtn.textContent = 'Enviando arquivos...';
    resultCard.classList.add('hidden');
    activeBar?.classList.add('hidden');

    const fd = new FormData();
    fd.append('list_token', listToken);
    fd.append('sheet', sheetSelect.value || '');
    fd.append('folder_name', name);
    pdfFiles.forEach(f => fd.append('pdfs', f, f.name));

    try {
      const data = await edgeJson('/api/jobs', { method: 'POST', body: fd });
      rememberJob(data.job_id);
      showProgress(true);
      pollJob(data.job_id, false);
    } catch (e) {
      processBtn.disabled = false;
      processBtn.textContent = 'Iniciar processamento';
      edgeToast(e.message, 'error');
    }
  });

  function showProgress(scroll = false) {
    progressCard.classList.remove('hidden');
    document.getElementById('progressBar').style.width = '2%';
    document.getElementById('progressPct').textContent = '0%';
    document.getElementById('progressTitle').textContent = 'Preparando os arquivos...';
    document.getElementById('progressDetail').textContent = 'Aguarde.';
    if (scroll) progressCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  async function pollJob(jobId, restored = false) {
    clearTimeout(pollTimer);
    try {
      const data = await edgeJson(`/api/jobs/${jobId}`);
      rememberJob(jobId);
      const pct = Math.max(2, data.progress || 0);
      document.getElementById('progressBar').style.width = `${pct}%`;
      document.getElementById('progressPct').textContent = `${Math.round(data.progress || 0)}%`;
      document.getElementById('progressTitle').textContent = data.message || 'Processando...';
      document.getElementById('progressDetail').textContent = data.total ? `${data.current}/${data.total} páginas` : 'Preparando leitura';
      if (data.status === 'done') return showResults(data, { restored });
      if (data.status === 'error') {
        const msg = data.error || 'Falha no processamento.';
        try { await edgeJson(`/api/jobs/${jobId}`, { method: 'DELETE' }); } catch (_) {}
        rememberJob('');
        throw new Error(msg);
      }
      pollTimer = setTimeout(() => pollJob(jobId, restored), 750);
    } catch (e) {
      processBtn.disabled = false;
      processBtn.textContent = 'Iniciar processamento';
      edgeToast(e.message, 'error');
    }
  }

  function showResults(data, { restored = false } = {}) {
    rememberJob(data.id);
    processBtn.disabled = true;
    processBtn.textContent = 'Limpe a extração para iniciar outra';
    const s = data.summary;

    document.getElementById('progressBar').style.width = '100%';
    document.getElementById('progressPct').textContent = '100%';
    document.getElementById('progressTitle').textContent = 'Concluído';
    document.getElementById('progressDetail').textContent = `${s.total_pages} página(s) analisada(s)`;
    if (restored) progressCard.classList.add('hidden');

    document.getElementById('kpiSaved').textContent = s.saved;
    document.getElementById('kpiPending').textContent = s.pending;
    document.getElementById('kpiDuplicates').textContent = s.duplicates;
    document.getElementById('kpiMissing').textContent = s.missing;
    document.getElementById('zipBtn').href = edgeUrl(data.download_zip);
    document.getElementById('reportBtn').href = edgeUrl(data.download_report);

    if (activeFolder) activeFolder.textContent = data.folder_name || 'Extração atual';
    activeBar?.classList.remove('hidden');
    if (data.folder_name) folderName.value = data.folder_name;
    updateArchiveInfo(data);

    const review = document.getElementById('reviewBtn');
    if (s.pending + s.duplicates > 0) {
      review.classList.remove('hidden');
      review.href = edgeUrl(data.review_url);
    } else {
      review.classList.add('hidden');
    }

    if (s.pending + s.duplicates > 0) {
      document.getElementById('resultSubtitle').textContent = 'Há páginas para revisar. O que você confirmar será incluído no mesmo download.';
    } else if (s.missing) {
      document.getElementById('resultSubtitle').textContent = 'A revisão foi atualizada, mas ainda existem exames não encontrados.';
    } else {
      document.getElementById('resultSubtitle').textContent = 'Tudo que estava previsto foi concluído. O download está atualizado.';
    }

    const rows = document.getElementById('resultRows');
    rows.innerHTML = s.analyses.length ? s.analyses.map(a => {
      const statusClass = a.status.includes('SALVO') || a.status.includes('ANEXADO')
        ? 'success'
        : (a.status === 'PENDENTE' ? 'warning' : 'neutral');
      const file = a.output_file
        ? `<a href="${edgeUrl(`/download/${data.id}/file/${encodeURIComponent(a.output_file).replaceAll('%2F', '/')}`)}" class="file-link">${escapeHtml(a.output_file)}</a>`
        : '—';
      return `<tr><td><strong>${escapeHtml(a.employee_name || '—')}</strong><small>${escapeHtml(a.company || '')}</small></td><td>${escapeHtml(a.exam_type || '—')}</td><td>${escapeHtml(a.exam_subtype || '—')}</td><td><span class="badge ${statusClass}">${escapeHtml(prettyStatus(a.status))}</span></td><td>${file}</td></tr>`;
    }).join('') : '<tr><td colspan="5" class="table-empty">Nenhum resultado relevante para exibir.</td></tr>';

    resultCard.classList.remove('hidden');
    if (!restored) {
      resultCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
      edgeToast('Processamento concluído.');
    }
  }

  async function restoreActiveJob() {
    try {
      const data = await edgeJson('/api/jobs/active');
      let job = data.active && data.job ? data.job : null;

      if (!job) {
        let savedId = '';
        try { savedId = localStorage.getItem(ACTIVE_JOB_KEY) || ''; } catch (_) {}
        if (savedId) {
          try {
            job = await edgeJson(`/api/jobs/${savedId}`);
          } catch (_) {
            try { localStorage.removeItem(ACTIVE_JOB_KEY); } catch (_) {}
          }
        }
      }

      if (!job) return;
      rememberJob(job.id);
      if (job.status === 'done' && job.summary) {
        showResults(job, { restored: true });
      } else if (job.status === 'queued' || job.status === 'processing') {
        processBtn.disabled = true;
        processBtn.textContent = 'Processando...';
        showProgress(false);
        pollJob(job.id, true);
      } else if (job.status === 'error') {
        try { await edgeJson(`/api/jobs/${job.id}`, { method: 'DELETE' }); } catch (_) {}
        rememberJob('');
      }
    } catch (_) {
      // Não impede um novo processamento se não houver extração ativa.
    }
  }

  clearBtn?.addEventListener('click', async () => {
    if (!currentJobId) return;
    if (!window.confirm('Limpar esta extração? Depois disso, os PDFs e o ZIP deste processamento não ficarão mais disponíveis para download.')) return;
    clearBtn.disabled = true;
    try {
      await edgeJson(`/api/jobs/${currentJobId}`, { method: 'DELETE' });
      clearTimeout(pollTimer);
      rememberJob('');
      resultCard.classList.add('hidden');
      progressCard.classList.add('hidden');
      activeBar?.classList.add('hidden');
      processBtn.disabled = false;
      processBtn.textContent = 'Iniciar processamento';
      if (archiveSaveInfo) archiveSaveInfo.textContent = 'Nenhum arquivo arquivado deste lote ainda.';
      if (archiveSaveBtn) archiveSaveBtn.disabled = false;
      const url = new URL(window.location.href);
      url.searchParams.delete('job');
      history.replaceState({}, '', url.pathname + url.search);
      edgeToast('Extração limpa. Agora os arquivos anteriores não estão mais disponíveis.');
    } catch (e) {
      edgeToast(e.message, 'error');
    } finally {
      clearBtn.disabled = false;
    }
  });

  function prettyStatus(s) {
    return ({
      SALVO_AUTOMATICO: 'Salvo',
      SALVO_MANUAL: 'Salvo manualmente',
      ANEXADO_CONTINUACAO: 'Anexado',
      PENDENTE: 'Pendente',
      DUPLICADO: 'Duplicado',
      IGNORADO_MANUAL: 'Ignorado'
    })[s] || s;
  }

  function escapeHtml(s = '') {
    return String(s).replace(/[&<>'"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c]));
  }

  initArchivePeriod();
  restoreActiveJob();
})();
