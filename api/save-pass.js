// Vercel serverless function: receives a gate pass PDF from the web page and
// saves it to OneDrive using the "gatepass" Entra app's secret (no user sign-in).
// The secret lives only in Vercel's environment variables, never in the web page.
//
// Required environment variables (Vercel > Project > Settings > Environment Variables):
//   TENANT_ID       f3f819ba-724b-4c0b-a9c0-2aa8d12bcbcc
//   CLIENT_ID       be3f583b-57fe-4777-8b75-c05a7d68794b
//   CLIENT_SECRET   the secret VALUE from Entra > gatepass > Certificates & secrets
//   ONEDRIVE_USER   ai@bharatsteels.in        (whose OneDrive receives the files)
//   ALLOWED_ORIGIN  https://bsc23609.github.io
//   BASE_FOLDER     Gate Pass                 (optional, defaults to "Gate Pass")

const MAX_BYTES = 3 * 1024 * 1024; // a gate pass PDF is ~100-300 KB
let cached = { token: null, expires: 0 };

async function getToken() {
  if (cached.token && Date.now() < cached.expires - 60000) return cached.token;
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const r = await fetch(
    `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
    { method: "POST", body }
  );
  const j = await r.json();
  if (!r.ok) throw new Error("token: " + (j.error_description || j.error || r.status));
  cached = { token: j.access_token, expires: Date.now() + j.expires_in * 1000 };
  return cached.token;
}

function setCors(req, res) {
  const allowed = (process.env.ALLOWED_ORIGIN || "").split(",").map((s) => s.trim());
  const origin = req.headers.origin || "";
  if (allowed.includes(origin)) res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  return allowed.includes(origin);
}

module.exports = async (req, res) => {
  const okOrigin = setCors(req, res);
  if (req.method === "OPTIONS") return res.status(204).end();
  if (req.method === "GET") return res.status(200).json({ ok: true, service: "gatepass-saver" });
  if (req.method !== "POST") return res.status(405).json({ error: "POST only" });
  if (!okOrigin) return res.status(403).json({ error: "origin not allowed" });

  try {
    const { fileName, year, month, pdfBase64 } = req.body || {};
    // Strict checks: only a PDF, only a safe file name, only inside the gate pass folder.
    if (!/^[A-Za-z0-9-]{1,40}_[A-Za-z0-9-]{0,40}\.pdf$/.test(fileName || ""))
      return res.status(400).json({ error: "bad file name" });
    if (!/^20\d\d$/.test(String(year)) || !/^(0[1-9]|1[0-2])-[A-Za-z]{3}$/.test(month || ""))
      return res.status(400).json({ error: "bad folder" });
    const pdf = Buffer.from(pdfBase64 || "", "base64");
    if (pdf.length < 100 || pdf.length > MAX_BYTES || pdf.subarray(0, 5).toString() !== "%PDF-")
      return res.status(400).json({ error: "not a PDF" });

    const base = process.env.BASE_FOLDER || "Gate Pass";
    const path = [base, String(year), month, fileName].map(encodeURIComponent).join("/");
    const user = encodeURIComponent(process.env.ONEDRIVE_USER);
    const token = await getToken();
    const up = await fetch(
      `https://graph.microsoft.com/v1.0/users/${user}/drive/root:/${path}:/content`,
      { method: "PUT", headers: { Authorization: "Bearer " + token, "Content-Type": "application/pdf" }, body: pdf }
    );
    const out = await up.json().catch(() => ({}));
    if (!up.ok) return res.status(502).json({ error: "onedrive " + up.status, detail: out.error && out.error.message });
    return res.status(200).json({ ok: true, path: `${base}/${year}/${month}/${fileName}`, webUrl: out.webUrl });
  } catch (e) {
    return res.status(500).json({ error: String(e.message || e) });
  }
};
