// api/upload.js  —— 以 5 個 ENV 為準 + CORS + OPTIONS
import formidable from "formidable";
import fs from "fs";

export const config = { api: { bodyParser: false } };

// ---- CORS ----
function setCORS(res) {
  // 如果要鎖網域，把 * 換成你的 Shopify 網域（含 https）
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

// ---- ENV helpers ----
function need(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing ENV: ${name}`);
  return v;
}

async function getToken() {
  const tenant = need("TENANT_ID");
  const client = need("CLIENT_ID");
  const secret = need("CLIENT_SECRET");

  const form = new URLSearchParams();
  form.append("grant_type", "client_credentials");
  form.append("client_id", client);
  form.append("client_secret", secret);
  form.append("scope", "https://graph.microsoft.com/.default");

  const url = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
  const r = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: form
  });

  const raw = await r.text();
  let js = null;
  try { js = JSON.parse(raw); } catch {}
  if (!r.ok || !js?.access_token) {
    console.error("[TOKEN_FAIL]", r.status, r.statusText, raw?.slice(0, 400));
    throw new Error(`Token failed (${r.status})`);
  }
  return js.access_token;
}

export default async function handler(req, res) {
  setCORS(res);
  if (req.method === "OPTIONS") {
    return res.status(204).end(); // 預檢直接放行
  }
  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, error: "Use POST" });
  }

  try {
    const upn = need("ONEDRIVE_USER_UPN");      // 例如 marketing@nanyaplastics-usa.com
    const root = need("ROOT_FOLDER");           // 例如 QuoteNeuma

    // 解析 multipart
    const form = formidable({ multiples: false, keepExtensions: true });
    const { fields, files } = await new Promise((resolve, reject) => {
      form.parse(req, (err, flds, fls) => (err ? reject(err) : resolve({ fields: flds, files: fls })));
    });

    const fileObj = files?.file;
    if (!fileObj) return res.status(400).json({ ok: false, error: "No file" });

    const subpath = String(fields?.subpath || "").replace(/^\/+/, "");
    const filename = String(fields?.filename || fileObj.originalFilename || "file.bin");

    // 讀檔
    const buf = fs.readFileSync(fileObj.filepath);

    const token = await getToken();

    // 目標路徑：/users/{UPN}/drive/root:/ROOT_FOLDER/subpath/filename:/content
    const prefix = root ? `${root}/${subpath}`.replace(/\/+$/,"") : subpath;
    const drivePath = prefix ? `${prefix}/${filename}` : filename;

    const url =
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}` +
      `/drive/root:/${encodeURIComponent(drivePath).replace(/%2F/g,"/")}:/content`;

    const upr = await fetch(url, {
      method: "PUT",
      headers: { Authorization: `Bearer ${token}` },
      body: buf
    });

    const upRaw = await upr.text();
    let upJs = null;
    try { upJs = JSON.parse(upRaw); } catch {}

    if (!upr.ok) {
      console.error("[UPLOAD_FAIL]", upr.status, upr.statusText, upRaw?.slice(0, 400));
      return res.status(500).json({
        ok: false,
        error: `Upload failed (${upr.status})`,
        hint: upJs?.error?.message || upRaw?.slice(0, 200)
      });
    }

    return res.status(200).json({ ok: true, item: upJs || null });
  } catch (e) {
    console.error("[SERVER_ERROR]", e?.message || e);
    return res.status(500).json({ ok: false, error: e?.message || String(e) });
  }
}
