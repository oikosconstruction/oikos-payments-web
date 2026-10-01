export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ ok: false, error: "Method not allowed" });
  }

  const flowUrl = process.env.OIKOS_POWER_AUTOMATE_URL;
  if (!flowUrl) {
    return res.status(503).json({
      ok: false,
      error: "Η σύνδεση με το Microsoft 365 δεν έχει ενεργοποιηθεί ακόμη."
    });
  }

  const body = req.body || {};
  const required = ["requestId", "project", "beneficiary", "amount", "paymentDate", "description"];
  const missing = required.filter(k => body[k] === undefined || body[k] === null || String(body[k]).trim() === "");
  if (missing.length) {
    return res.status(400).json({ ok: false, error: "Λείπουν υποχρεωτικά πεδία.", missing });
  }

  const amount = Number(body.amount);
  if (!Number.isFinite(amount) || amount <= 0) {
    return res.status(400).json({ ok: false, error: "Μη έγκυρο ποσό." });
  }

  const attachmentBase64 = typeof body.attachmentBase64 === "string" ? body.attachmentBase64 : "";
  if (attachmentBase64.length > 8_000_000) {
    return res.status(413).json({ ok: false, error: "Το συνημμένο είναι πολύ μεγάλο." });
  }

  const payload = {
    requestId: String(body.requestId),
    createdAt: String(body.createdAt || new Date().toISOString()),
    project: String(body.project),
    beneficiary: String(body.beneficiary),
    amount,
    paymentDate: String(body.paymentDate),
    description: String(body.description),
    ibanRf: String(body.ibanRf || ""),
    bank: String(body.bank || ""),
    documentNo: String(body.documentNo || ""),
    notes: String(body.notes || ""),
    urgent: Boolean(body.urgent),
    attachmentName: String(body.attachmentName || ""),
    attachmentType: String(body.attachmentType || ""),
    attachmentBase64
  };

  try {
    const upstream = await fetch(flowUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const text = await upstream.text();
    let data = {};
    try { data = text ? JSON.parse(text) : {}; } catch { data = { message: text }; }

    if (!upstream.ok) {
      console.error("Power Automate error", upstream.status, data);
      return res.status(502).json({ ok: false, error: "Αποτυχία Microsoft 365 workflow." });
    }

    return res.status(200).json({ ok: true, requestId: payload.requestId, workflow: data });
  } catch (err) {
    console.error("Submit payment error", err);
    return res.status(502).json({ ok: false, error: "Δεν ήταν δυνατή η επικοινωνία με το Microsoft 365." });
  }
}
