(() => {
  const $ = (id) => document.getElementById(id);
  const form = $('paymentForm');
  const historyCard = $('historyCard');
  const formCard = $('formCard');
  const storageKey = 'oikos_payments_mvp_v1';

  $('paymentDate').valueAsDate = new Date();
  $('attachment').addEventListener('change', e => {
    $('fileLabel').textContent = e.target.files[0]?.name || 'Επιλογή αρχείου';
  });

  $('newBtn').onclick = () => { formCard.classList.remove('hidden'); historyCard.classList.add('hidden'); };
  $('historyBtn').onclick = () => { renderHistory(); formCard.classList.add('hidden'); historyCard.classList.remove('hidden'); };
  $('saveDraftBtn').onclick = () => savePayment(true);
  $('exportBtn').onclick = exportCsv;

  form.addEventListener('submit', async e => { e.preventDefault(); await savePayment(false); });

  function values() {
    const file = $('attachment').files[0];
    return {
      requestId: crypto.randomUUID ? crypto.randomUUID() : `PAY-${Date.now()}`,
      createdAt: new Date().toISOString(),
      project: $('project').value.trim(),
      beneficiary: $('beneficiary').value.trim(),
      amount: Number($('amount').value || 0),
      paymentDate: $('paymentDate').value,
      description: $('description').value.trim(),
      workCode: $('workCode').value.trim(),
      ibanRf: $('ibanRf').value.trim(),
      bank: $('bank').value.trim(),
      documentNo: $('documentNo').value.trim(),
      notes: $('notes').value.trim(),
      urgent: $('urgent').checked,
      attachmentName: file?.name || '',
      attachmentType: file?.type || '',
      attachmentSize: file?.size || 0,
      status: 'PENDING'
    };
  }

  async function savePayment(draft) {
    clearMessage();
    if (!draft && !form.reportValidity()) return;
    const data = values();
    data.status = draft ? 'DRAFT' : 'PENDING';

    try {
      if (!draft) {
        data.attachmentBase64 = await fileToBase64($('attachment').files[0]);
        const endpoint = window.OIKOS_CONFIG.apiEndpoint || '/api/submit-payment';
        const res = await fetch(endpoint, {
          method: 'POST',
          headers: {'Content-Type':'application/json'},
          body: JSON.stringify(data)
        });
        const result = await res.json().catch(() => ({}));
        if (!res.ok) throw new Error(result.error || `HTTP ${res.status}`);
        data.status = 'SUBMITTED';
      }

      persistLocal(data);
      showMessage(
        draft
          ? 'Το πρόχειρο αποθηκεύτηκε.'
          : (data.urgent
              ? 'Η επείγουσα πληρωμή υποβλήθηκε επιτυχώς.'
              : 'Η πληρωμή υποβλήθηκε επιτυχώς.'),
        'ok'
      );
      if (!draft) resetForm();
    } catch (err) {
      showMessage(`Αποτυχία καταχώρησης: ${err.message}`, 'err');
    }
  }

  function fileToBase64(file) {
    if (!file) return Promise.resolve('');
    return new Promise((resolve,reject) => {
      const r = new FileReader();
      r.onload = () => resolve(String(r.result).split(',')[1] || '');
      r.onerror = reject;
      r.readAsDataURL(file);
    });
  }

  function persistLocal(item) {
    const arr = JSON.parse(localStorage.getItem(storageKey) || '[]');
    arr.unshift(item);
    localStorage.setItem(storageKey, JSON.stringify(arr.slice(0,250)));
  }

  function renderHistory() {
    const arr = JSON.parse(localStorage.getItem(storageKey) || '[]');
    $('historyBody').innerHTML = arr.length ? arr.map(x => `
      <tr>
        <td>${esc(x.paymentDate || '')}</td>
        <td>${esc(x.project || '')}</td>
        <td>${esc(x.beneficiary || '')}</td>
        <td class="amount">${fmt(x.amount)}</td>
        <td><span class="tag ${x.status==='DRAFT'?'draft':x.urgent?'urgent':'normal'}">${x.status==='DRAFT'?'ΠΡΟΧΕΙΡΟ':x.urgent?'ΕΠΕΙΓΟΝ':'ΚΑΝΟΝΙΚΗ'}</span></td>
        <td>${esc(x.status)}</td>
      </tr>`).join('') : '<tr><td colspan="6">Δεν υπάρχουν καταχωρήσεις ακόμα.</td></tr>';
  }

  function exportCsv() {
    const arr = JSON.parse(localStorage.getItem(storageKey) || '[]');
    const cols = ['createdAt','requestId','project','beneficiary','amount','paymentDate','description','workCode','ibanRf','bank','documentNo','urgent','status','attachmentName','notes'];
    const rows = [cols.join(';'), ...arr.map(x => cols.map(c => csv(x[c])).join(';'))];
    const blob = new Blob(['\ufeff'+rows.join('\n')], {type:'text/csv;charset=utf-8'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'OIKOS_payments.csv';
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function resetForm(){ form.reset(); $('paymentDate').valueAsDate = new Date(); $('fileLabel').textContent='Επιλογή αρχείου'; }
  function showMessage(t,c){ $('message').textContent=t; $('message').className=`message ${c}`; }
  function clearMessage(){ showMessage('',''); }
  function fmt(n){ return new Intl.NumberFormat('el-GR',{style:'currency',currency:'EUR'}).format(Number(n||0)); }
  function esc(s){ return String(s??'').replace(/[&<>'"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[m])); }
  function csv(v){ const s=String(v??'').replaceAll('"','""'); return `"${s}"`; }
})();
