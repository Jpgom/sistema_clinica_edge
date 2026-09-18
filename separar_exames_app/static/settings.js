(() => {
  const auto = document.getElementById('autoThreshold'), emp = document.getElementById('employeeThreshold');
  auto.oninput = () => document.getElementById('autoValue').textContent = auto.value + '%';
  emp.oninput = () => document.getElementById('employeeValue').textContent = emp.value + '%';
  document.getElementById('saveSettings').onclick = async () => {
    const btn = document.getElementById('saveSettings'); btn.disabled = true;
    try {
      await edgeJson('/api/config', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({
        use_ocr: document.getElementById('useOcr').checked,
        fast_mode: document.getElementById('fastMode').checked,
        stop_when_complete: false,
        auto_threshold: +auto.value,
        employee_threshold: +emp.value
      })});
      document.getElementById('saveStatus').textContent = 'Configurações salvas.'; edgeToast('Configurações salvas.');
    } catch (e) { edgeToast(e.message, 'error'); }
    btn.disabled = false;
  };
})();
