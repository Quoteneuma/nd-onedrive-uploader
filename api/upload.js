// api/upload.js — with verbose logs
import formidable from "formidable";
import fs from "fs";

export const config = { api: { bodyParser: false } };

function mustEnv(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing ENV: ${name}`);
  return v;
}

async function getToken() {
  const tenant = mustEnv("TENANT_ID");
  const client = mustEnv("CLIENT_ID");
  const secret = mustEnv("CLIENT_SECRET");

  const form = new URLSearchParams();
  form.append("grant_type", "client_credentials");
  form.append("client_id", client);
  form.append("client_secret", secret);
  form.append("scope", "https://graph.microsoft.com/.default");

  const url = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
  const r = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: form,
  });

  const raw = await r.text();
  let js = null;
  try { js = JSON.parse(raw); } catch { /* keep raw */ }

  if (!r.ok || !js?.access_token) {
    console.error("[TOKEN_FAIL]", r.status, r.statusText, raw?.slice(0, 400));
    throw new Error(`Token failed (${r.status})`);
  }

  return js.access_token;
}

export default async function handler(req, res) {
  const start = Date.now();
  if (req.method !== "POST") {
    return res.status(400).json({ ok: false, error: "Use POST" });
  }

  try {
    // 解析表單
    const form = formidable({ multiples: false });
    const { fields, files } = await new Promise((resolve, reject) => {
      form.parse(req, (err, flds, fls) => (err ? reject(err) : resolve({ fields: flds, files: fls })));
    });

    const driveUser = mustEnv("ONEDRIVE_USER_UPN"); // 例如 marketing@nanyaplastics-usa.com
    const subpath = String(fields.subpath || "").replace(/^\/+/, "");
    const filename = String(fields.filename || files?.file?.originalFilename || "file.bin");

    if (!files?.file) {
      return res.status(400).json({ ok: false, error: "No file uploaded" });
    }

    // 先拿 Token
    const token = await getToken();

    // 讀檔
    const buffer = fs.readFileSync(files.file.filepath);

    // 目標：User Drive（app-only 不可用 /me/drive，要用 /users/{UPN}/drive）
    const uploadUrl =
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(driveUser)}` +
      `/drive/root:/${subpath ? subpath + "/" : ""}${encodeURIComponent(filename)}:/content`;

    const upr = await fetch(uploadUrl, {
      method: "PUT",
      headers: { Authorization: `Bearer ${token}` },
      body: buffer,
    });

    const upRaw = await upr.text();
    let upJs = null;
    try { upJs = JSON.parse(upRaw); } catch {}

    if (!upr.ok) {
      console.error("[UPLOAD_FAIL]", upr.status, upr.statusText, upRaw?.slice(0, 400));
      return res.status(500).json({
        ok: false,
        error: `Upload failed (${upr.status})`,
        hint: upJs?.error?.message || upRaw?.slice(0, 200),
      });
    }

    console.log("[UPLOAD_OK]", filename, `${Date.now() - start}ms`);
    return res.status(200).json({ ok: true, item: upJs || null });
  } catch (e) {
    console.error("[SERVER_ERROR]", e?.message || e);
    return res.status(500).json({ ok: false, error: e?.message || String(e) });
  }
}
