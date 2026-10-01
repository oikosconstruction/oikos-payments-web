import ExcelJS from 'exceljs';

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';
let tokenCache = { token: null, expiresAt: 0 };

function env(name, required = true) {
  const value = process.env[name];
  if (required && !value) throw new Error(`Missing environment variable: ${name}`);
  return value || '';
}

function json(res, status, payload) {
  res.status(status).json(payload);
}

function sanitizeFileName(name = 'attachment') {
  return name.replace(/[\\/:*?"<>|\u0000-\u001F]/g, '_').slice(0, 140) || 'attachment';
}

function graphHeaders(token, extra = {}) {
  return { Authorization: `Bearer ${token}`, ...extra };
}

async function getGraphToken() {
  if (tokenCache.token && Date.now() < tokenCache.expiresAt - 60_000) return tokenCache.token;
  const tenantId = env('OIKOS_MS_TENANT_ID');
  const clientId = env('OIKOS_MS_CLIENT_ID');
  const clientSecret = env('OIKOS_MS_CLIENT_SECRET');

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  const r = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });
  const data = await r.json();
  if (!r.ok || !data.access_token) throw new Error(`Microsoft token error: ${data.error_description || r.status}`);

  tokenCache = {
    token: data.access_token,
    expiresAt: Date.now() + Number(data.expires_in || 3600) * 1000
  };
  return tokenCache.token;
}

async function graphJson(path, token, options = {}) {
  const r = await fetch(`${GRAPH_ROOT}${path}`, {
    ...options,
    headers: graphHeaders(token, { 'Content-Type': 'application/json', ...(options.headers || {}) })
  });
  const text = await r.text();
  let data = null;
  try { data = text ? JSON.parse(text) : null; } catch { data = { raw: text }; }
  if (!r.ok) {
    const err = new Error(data?.error?.message || data?.raw || `Graph ${r.status}`);
    err.status = r.status;
    err.payload = data;
    throw err;
  }
  return data;
}

async function graphBinary(path, token) {
  const r = await fetch(`${GRAPH_ROOT}${path}`, { headers: graphHeaders(token) });
  if (!r.ok) throw new Error(`Graph binary download failed: ${r.status}`);
  return Buffer.from(await r.arrayBuffer());
}

async function getItemMetadata(driveId, itemId, token) {
  return graphJson(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}?$select=id,name,eTag,lastModifiedDateTime,size,parentReference,webUrl`, token);
}

async function listChildren(driveId, parentId, token) {
  const result = await graphJson(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}/children?$select=id,name,folder,webUrl`, token);
  return result?.value || [];
}

async function ensureFolder(driveId, parentId, name, token) {
  const children = await listChildren(driveId, parentId, token);
  const existing = children.find(x => x.folder && x.name?.toLowerCase() === name.toLowerCase());
  if (existing) return existing;
  return graphJson(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}/children`, token, {
    method: 'POST',
    body: JSON.stringify({
      name,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'fail'
    })
  });
}

async function ensureNestedFolder(driveId, parentId, segments, token) {
  let current = { id: parentId };
  for (const segment of segments) current = await ensureFolder(driveId, current.id, segment, token);
  return current;
}

async function uploadSmallFile(driveId, parentId, fileName, bytes, mimeType, token) {
  const path = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}:/${encodeURIComponent(fileName)}:/content`;
  const r = await fetch(`${GRAPH_ROOT}${path}`, {
    method: 'PUT',
    headers: graphHeaders(token, { 'Content-Type': mimeType || 'application/octet-stream' }),
    body: bytes
  });
  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(data?.error?.message || `Attachment upload failed: ${r.status}`);
  return data;
}

function validatePayment(body) {
  const required = ['requestId', 'project', 'beneficiary', 'amount', 'paymentDate', 'description'];
  const missing = required.filter(k => body[k] === undefined || body[k] === null || String(body[k]).trim() === '');
  if (missing.length) return { ok: false, error: 'Λείπουν υποχρεωτικά πεδία.', missing };
  const amount = Number(body.amount);
  if (!Number.isFinite(amount) || amount <= 0) return { ok: false, error: 'Μη έγκυρο ποσό.' };
  if ((body.attachmentBase64 || '').length > 4_000_000) return { ok: false, error: 'Το συνημμένο είναι πολύ μεγάλο για την τρέχουσα έκδοση (μέγιστο περίπου 3 MB).' };
  return { ok: true };
}

function paymentRow(body, attachmentUrl = '', notificationStatus = '') {
  return [
    String(body.project || ''),
    String(body.description || ''),
    String(body.workCode || ''),
    String(body.beneficiary || ''),
    Number(body.amount || 0),
    String(body.ibanRf || ''),
    String(body.bank || ''),
    attachmentUrl || '',
    String(body.documentNo || ''),
    'PENDING',
    body.paymentDate ? new Date(`${body.paymentDate}T12:00:00`) : null,
    String(body.notes || ''),
    body.urgent ? 'ΕΠΕΙΓΟΝ' : 'ΚΑΝΟΝΙΚΗ',
    '',
    'YES',
    body.createdAt ? new Date(body.createdAt) : new Date(),
    notificationStatus || (body.urgent ? 'PENDING' : ''),
    null,
    String(body.requestId || ''),
    'WEB APP / GRAPH'
  ];
}

async function writeWorkbookRow(body, attachmentUrl, token) {
  const driveId = env('OIKOS_SP_DRIVE_ID');
  const workbookItemId = env('OIKOS_SP_WORKBOOK_ITEM_ID');

  let lastError = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const meta = await getItemMetadata(driveId, workbookItemId, token);
      const fileBuffer = await graphBinary(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(workbookItemId)}/content`, token);

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(fileBuffer);
      const ws = workbook.getWorksheet('ΤΡΕΧΟΥΣΑ ΕΒΔΟΜΑΔΑ');
      if (!ws) throw new Error('Δεν βρέθηκε το φύλλο ΤΡΕΧΟΥΣΑ ΕΒΔΟΜΑΔΑ.');
      const table = ws.getTable('tblCurrentPayments');
      if (!table) throw new Error('Δεν βρέθηκε ο πίνακας tblCurrentPayments.');

      table.addRow(paymentRow(body, attachmentUrl));
      table.commit();

      const out = Buffer.from(await workbook.xlsx.writeBuffer());
      const r = await fetch(`${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(workbookItemId)}/content`, {
        method: 'PUT',
        headers: graphHeaders(token, {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          ...(meta?.eTag ? { 'If-Match': meta.eTag } : {})
        }),
        body: out
      });
      if (r.status === 412) throw Object.assign(new Error('Workbook changed concurrently'), { status: 412 });
      const result = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error(result?.error?.message || `Workbook upload failed: ${r.status}`);
      return result;
    } catch (err) {
      lastError = err;
      if (err.status === 412 && attempt < 3) continue;
      throw err;
    }
  }
  throw lastError || new Error('Workbook update failed.');
}

async function updateNotificationStatus(requestId, status, detail, token) {
  const driveId = env('OIKOS_SP_DRIVE_ID');
  const workbookItemId = env('OIKOS_SP_WORKBOOK_ITEM_ID');

  for (let attempt = 1; attempt <= 3; attempt++) {
    const meta = await getItemMetadata(driveId, workbookItemId, token);
    const fileBuffer = await graphBinary(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(workbookItemId)}/content`, token);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(fileBuffer);
    const ws = workbook.getWorksheet('ΤΡΕΧΟΥΣΑ ΕΒΔΟΜΑΔΑ');
    if (!ws) throw new Error('Δεν βρέθηκε το φύλλο πληρωμών.');

    let targetRow = null;
    for (let r = 10; r <= Math.max(ws.rowCount, 500); r++) {
      if (String(ws.getCell(r, 19).value || '') === String(requestId)) { targetRow = r; break; }
    }
    if (!targetRow) throw new Error('Δεν βρέθηκε η εγγραφή για ενημέρωση ειδοποίησης.');

    ws.getCell(targetRow, 17).value = status;
    ws.getCell(targetRow, 18).value = status === 'SENT' ? new Date() : null;
    ws.getCell(targetRow, 20).value = detail || '';

    const out = Buffer.from(await workbook.xlsx.writeBuffer());
    const r = await fetch(`${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(workbookItemId)}/content`, {
      method: 'PUT',
      headers: graphHeaders(token, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ...(meta?.eTag ? { 'If-Match': meta.eTag } : {})
      }),
      body: out
    });
    if (r.status === 412 && attempt < 3) continue;
    if (!r.ok) throw new Error(`Notification status update failed: ${r.status}`);
    return;
  }
}

async function sendUrgentEmail(body, attachmentUrl, token) {
  const sender = env('OIKOS_MAIL_SENDER', false);
  const recipient1 = env('OIKOS_ACCOUNTING_EMAIL_1', false);
  const recipient2 = env('OIKOS_ACCOUNTING_EMAIL_2', false);
  const recipients = [recipient1, recipient2].filter(Boolean);
  if (!sender || recipients.length < 2) return { sent: false, reason: 'Email configuration incomplete' };

  const euro = new Intl.NumberFormat('el-GR', { style: 'currency', currency: 'EUR' }).format(Number(body.amount || 0));
  const html = `
    <p><strong>ΕΠΕΙΓΟΥΣΑ ΠΛΗΡΩΜΗ</strong></p>
    <p>
      <strong>Έργο:</strong> ${escapeHtml(body.project)}<br>
      <strong>Δικαιούχος:</strong> ${escapeHtml(body.beneficiary)}<br>
      <strong>Ποσό:</strong> ${escapeHtml(euro)}<br>
      <strong>Αιτιολογία:</strong> ${escapeHtml(body.description)}<br>
      <strong>IBAN / RF:</strong> ${escapeHtml(body.ibanRf || '-')}<br>
      <strong>Παραστατικό:</strong> ${escapeHtml(body.documentNo || '-')}<br>
      ${attachmentUrl ? `<strong>Αρχείο:</strong> <a href="${escapeHtml(attachmentUrl)}">Άνοιγμα στο SharePoint</a><br>` : ''}
      <strong>Request ID:</strong> ${escapeHtml(body.requestId)}
    </p>`;

  await graphJson(`/users/${encodeURIComponent(sender)}/sendMail`, token, {
    method: 'POST',
    body: JSON.stringify({
      message: {
        subject: `ΕΠΕΙΓΟΥΣΑ ΠΛΗΡΩΜΗ — ${body.project} — ${body.beneficiary} — ${euro}`,
        body: { contentType: 'HTML', content: html },
        toRecipients: recipients.map(address => ({ emailAddress: { address } }))
      },
      saveToSentItems: true
    })
  });
  return { sent: true };
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, m => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[m]));
}

async function storeAuditJson(body, attachmentUrl, token) {
  const driveId = env('OIKOS_SP_DRIVE_ID');
  const parentId = env('OIKOS_SP_PARENT_FOLDER_ID');
  const now = new Date(body.createdAt || Date.now());
  const folder = await ensureNestedFolder(driveId, parentId, ['WEB_REQUESTS', String(now.getUTCFullYear())], token);
  const audit = Buffer.from(JSON.stringify({ ...body, attachmentBase64: undefined, attachmentUrl }, null, 2), 'utf8');
  return uploadSmallFile(driveId, folder.id, `${sanitizeFileName(body.requestId)}.json`, audit, 'application/json', token);
}

async function uploadAttachment(body, token) {
  if (!body.attachmentBase64 || !body.attachmentName) return null;
  const driveId = env('OIKOS_SP_DRIVE_ID');
  const parentId = env('OIKOS_SP_PARENT_FOLDER_ID');
  const now = new Date(body.createdAt || Date.now());
  const folder = await ensureNestedFolder(driveId, parentId, ['ATTACHMENTS', String(now.getUTCFullYear())], token);
  const name = `${sanitizeFileName(body.requestId)}_${sanitizeFileName(body.attachmentName)}`;
  const bytes = Buffer.from(body.attachmentBase64, 'base64');
  return uploadSmallFile(driveId, folder.id, name, bytes, body.attachmentType || 'application/octet-stream', token);
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return json(res, 405, { ok: false, error: 'Method not allowed' });
  }

  const body = req.body || {};
  const check = validatePayment(body);
  if (!check.ok) return json(res, 400, check);

  try {
    const token = await getGraphToken();
    const attachment = await uploadAttachment(body, token);
    const attachmentUrl = attachment?.webUrl || '';

    await storeAuditJson(body, attachmentUrl, token);
    await writeWorkbookRow(body, attachmentUrl, token);

    let email = { sent: false, reason: 'Not urgent' };
    if (body.urgent) {
      try {
        email = await sendUrgentEmail(body, attachmentUrl, token);
        if (email.sent) {
          await updateNotificationStatus(body.requestId, 'SENT', 'EMAIL SENT VIA MICROSOFT GRAPH', token);
        }
      } catch (emailError) {
        console.error('Urgent email failed', emailError);
        email = { sent: false, reason: emailError.message };
        try { await updateNotificationStatus(body.requestId, 'ERROR', `EMAIL ERROR: ${emailError.message}`, token); } catch {}
      }
    }

    return json(res, 200, {
      ok: true,
      requestId: body.requestId,
      attachmentUrl,
      urgentEmailSent: Boolean(email.sent),
      urgentEmailDetail: email.reason || null
    });
  } catch (err) {
    console.error('submit-payment failed', err);
    const configurationError = /Missing environment variable/.test(err.message || '');
    return json(res, configurationError ? 503 : 502, {
      ok: false,
      error: configurationError
        ? 'Η εταιρική σύνδεση Microsoft 365 δεν έχει ενεργοποιηθεί ακόμη.'
        : 'Η καταχώρηση στο Microsoft 365 απέτυχε.',
      detail: process.env.NODE_ENV === 'development' ? err.message : undefined
    });
  }
}
