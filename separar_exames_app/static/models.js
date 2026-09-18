(() => {
  const file = document.getElementById('modelFile'), btn = document.getElementById('addModel');
  const modelExam = document.getElementById('modelExam');
  const addExamType = document.getElementById('addExamType');
  const examTypeName = document.getElementById('examTypeName');
  const examTypeAliases = document.getElementById('examTypeAliases');
  const examTypeList = document.getElementById('examTypeList');
  const esc = s => String(s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

  file.onchange = () => {
    const n = file.files.length;
    document.getElementById('modelFileLabel').textContent = n ? (n === 1 ? file.files[0].name : `${n} PDFs selecionados`) : 'Selecionar PDFs modelo';
  };

  function renderExamTypes(types = []) {
    const current = modelExam.value;
    modelExam.innerHTML = types.map(t => `<option value="${esc(t.name)}">${esc(t.name)}</option>`).join('');
    if ([...modelExam.options].some(o => o.value === current)) modelExam.value = current;
    examTypeList.innerHTML = types.map(t => `
      <div class="mini-list-item">
        <span><strong>${esc(t.name)}</strong><small>${t.default ? 'Tipo padrão' : 'Tipo cadastrado'} · ${Number(t.model_count || 0)} modelo(s)</small></span>
        ${t.default ? '' : `<button class="icon-btn danger-text exam-type-delete" data-name="${esc(t.name)}" type="button">Excluir</button>`}
      </div>`).join('');
    document.querySelectorAll('.exam-type-delete').forEach(b => b.onclick = async () => {
      const name = b.dataset.name || '';
      if (!confirm(`Excluir o tipo de exame ${name}?`)) return;
      try { await edgeJson(`/api/tipos-exames/${encodeURIComponent(name)}`, {method:'DELETE'}); await loadTypes(); await load(); }
      catch(e){ edgeToast(e.message,'error'); }
    });
  }

  async function loadTypes() {
    const data = await edgeJson('/api/tipos-exames');
    renderExamTypes(data.types || []);
  }

  async function load(){
    try{
      const data = await edgeJson('/api/modelos');
      if (data.exam_types) renderExamTypes((data.exam_types || []).map(name => ({name, default: ['ASO','AUDIOMETRIA','ESPIROMETRIA','ACUIDADE VISUAL','LAUDO PCD'].includes(name), model_count: 0})));
      document.getElementById('modelCount').textContent = `${data.models.length} modelo(s)`;
      document.getElementById('modelRows').innerHTML = data.models.length ? data.models.map(m => `<tr><td>${esc(m.exam_type)}</td><td><strong>${esc(m.label)}</strong></td><td>${esc(m.source_filename)}</td><td>${m.page_number}</td><td><button class="icon-btn danger-text" data-id="${m.id}">Excluir</button></td></tr>`).join('') : '<tr><td colspan="5" class="table-empty">Nenhum modelo cadastrado.</td></tr>';
      document.querySelectorAll('[data-id]').forEach(b => b.onclick = async () => { if (!confirm('Excluir este modelo?')) return; await edgeJson(`/api/modelos/${b.dataset.id}`, {method:'DELETE'}); await loadTypes(); load(); });
      await loadTypes();
    }catch(e){edgeToast(e.message,'error');}
  }

  addExamType.onclick = async () => {
    const name = (examTypeName.value || '').trim();
    if (!name) return edgeToast('Informe o nome do tipo de exame.', 'error');
    addExamType.disabled = true;
    try {
      await edgeJson('/api/tipos-exames', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({name, aliases: examTypeAliases.value || ''})
      });
      examTypeName.value = ''; examTypeAliases.value = '';
      await loadTypes();
      modelExam.value = name.toUpperCase();
      edgeToast('Tipo de exame cadastrado. Agora envie os PDFs modelo desse tipo.');
    } catch(e){ edgeToast(e.message,'error'); }
    addExamType.disabled = false;
  };

  btn.onclick = async () => {
    if (!file.files.length) return edgeToast('Selecione um ou mais PDFs modelo.', 'error');
    btn.disabled=true; btn.textContent=`Cadastrando ${file.files.length} arquivo(s)...`;
    const fd = new FormData();
    [...file.files].forEach(f => fd.append('files', f));
    fd.append('exam_type', modelExam.value);
    fd.append('label',document.getElementById('modelLabel').value);
    try {
      const d = await edgeJson('/api/modelos', {method:'POST',body:fd});
      edgeToast(`${d.added} modelo(s) cadastrado(s) a partir de ${d.files || file.files.length} PDF(s).`);
      file.value=''; document.getElementById('modelFileLabel').textContent='Selecionar PDFs modelo'; await loadTypes(); load();
    } catch(e){edgeToast(e.message,'error');}
    btn.disabled=false; btn.textContent='Cadastrar modelo';
  };
  load();
})();
