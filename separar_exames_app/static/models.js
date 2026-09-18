(() => {
  const file = document.getElementById('modelFile'), btn = document.getElementById('addModel');
  const esc = s => String(s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  file.onchange = () => {
    const n = file.files.length;
    document.getElementById('modelFileLabel').textContent = n ? (n === 1 ? file.files[0].name : `${n} PDFs selecionados`) : 'Selecionar PDFs modelo';
  };
  async function load(){
    try{
      const data = await edgeJson('/api/modelos');
      document.getElementById('modelCount').textContent = `${data.models.length} modelo(s)`;
      document.getElementById('modelRows').innerHTML = data.models.length ? data.models.map(m => `<tr><td>${esc(m.exam_type)}</td><td><strong>${esc(m.label)}</strong></td><td>${esc(m.source_filename)}</td><td>${m.page_number}</td><td><button class="icon-btn danger-text" data-id="${m.id}">Excluir</button></td></tr>`).join('') : '<tr><td colspan="5" class="table-empty">Nenhum modelo cadastrado.</td></tr>';
      document.querySelectorAll('[data-id]').forEach(b => b.onclick = async () => { if (!confirm('Excluir este modelo?')) return; await edgeJson(`/api/modelos/${b.dataset.id}`, {method:'DELETE'}); load(); });
    }catch(e){edgeToast(e.message,'error');}
  }
  btn.onclick = async () => {
    if (!file.files.length) return edgeToast('Selecione um ou mais PDFs modelo.', 'error');
    btn.disabled=true; btn.textContent=`Cadastrando ${file.files.length} arquivo(s)...`;
    const fd = new FormData();
    [...file.files].forEach(f => fd.append('files', f));
    fd.append('exam_type',document.getElementById('modelExam').value);
    fd.append('label',document.getElementById('modelLabel').value);
    try {
      const d = await edgeJson('/api/modelos', {method:'POST',body:fd});
      edgeToast(`${d.added} modelo(s) cadastrado(s) a partir de ${d.files || file.files.length} PDF(s).`);
      file.value=''; document.getElementById('modelFileLabel').textContent='Selecionar PDFs modelo'; load();
    } catch(e){edgeToast(e.message,'error');}
    btn.disabled=false; btn.textContent='Cadastrar modelo';
  };
  load();
})();
