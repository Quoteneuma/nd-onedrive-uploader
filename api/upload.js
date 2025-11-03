// api/upload.js — use Busboy (no Vercel file service), upload <=4MB via /content
import Busboy from "busboy";

export const config = { api: { bodyParser: false } };

// ----- env helpers -----
function need(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing ENV: ${name}`);
  return v;
}

// Encode單一段路徑（不把 / 編碼）
function encSeg(s) {
  return encodeURIComponent(String(s || "")).replace(/%2F/gi, "/");
}

// ----- MS Graph token -----
async function getToken() {
  const tenant = need("TENANT_ID");
  const client = need("CLIENT_ID");
  const secret = need("CLIENT_SECRET");

  const body = new URLSearchParams();
  body.append("grant_type", "client_credentials");
  body.append("client_id", client);
  body.append("client_secret", secret);
  body.append("scope", "https://graph.microsoft.com/.default");

  const url = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
  const r = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
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

// ----- multipart 解析（Busboy）-----
function parseMultipart(req) {
  return new Promise((resolve, reject) => {
    try {
      const bb = Busboy({ headers: req.headers, limits: { files: 1, fileSize: 25 * 1024 * 1024 } }); // 25MB 上限（>4MB 要改用上傳工作階段）
      const fields = {};
      let filename = "file.bin";
      let fileBufs = [];

      bb.on("file", (_name, stream, info) => {
        if (info?.filename) filename = info.filename;
        stream.on("data", (d) => fileBufs.push(d));
        stream.on("limit", () => reject(new Error("File too large")));
        stream.on("error", reject);
      });

      bb.on("field", (name, val) => {
        fields[name] = val;
      });

      bb.on("finish", () => {
        const buffer = Buffer.concat(fileBufs);
        resolve({ fields, buffer, filename });
      });

      req.pipe(bb);
    } catch (e) {
      reject(e);
    }
  });
}

// ----- handler -----
export default async function handler(req, res) {
  const t0 = Date.now();
  if (req.method !== "POST") {
    return res.status(400).json({ ok: false, error: "Use POST (multipart/form-data)" });
  }

  try {
    // 1) 解析 multipart
    const { fields, buffer, filename: nameFromForm } = await parseMultipart(req);
    if (!buffer || buffer.length === 0) {
      return res.status(400).json({ ok: false, error: "No file received" });
    }

    // 2) 取 env 與路徑
    const driveUser = need("ONEDRIVE_USER_UPN");
    const root = need("ROOT_FOLDER"); // e.g. QuoteNeuma
    const subpathRaw = String(fields.subpath || "").trim();  // e.g. QuoteNeuma\someone@email\2025-11-03
    const overrideName = String(fields.filename || "").trim();
    const finalName = overrideName || nameFromForm || "file.bin";

    // subpath 正規化（把 \ 換成 /，分段編碼）
    const normSub = subpathRaw.replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
    const segs = [];
    if (root) segs.push(root);
    if (normSub) segs.push(...normSub.split("/").filter(Boolean));
    segs.push(finalName);
    const pathForGraph = segs.map(encSeg).join("/");

    // 3) 取得 Graph token
    const token = await getToken();

    // 4) 直接 PUT /content（<=4MB；若大於 4MB 需改為 uploadSession）
    const uploadUrl =
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(driveUser)}` +
      `/drive/root:/${pathForGraph}:/content`;

    const upr = await fetch(uploadUrl, {
      method: "PUT",
      headers: { Authorization: `Bearer ${token}` },
      body: buffer
    });

    const respText = await upr.text();
    let js = null; try { js = JSON.parse(respText); } catch {}

    if (!upr.ok) {
      console.error("[UPLOAD_FAIL]", upr.status, upr.statusText, respText?.slice(0, 400));
      return res.status(500).json({
        ok: false,
        error: `Upload failed (${upr.status})`,
        hint: js?.error?.message || respText?.slice(0, 200)
      });
    }

    console.log("[UPLOAD_OK]", finalName, `${Date.now() - t0}ms`);
    return res.status(200).json({ ok: true, item: js || null });
  } catch (e) {
    console.error("[SERVER_ERROR]", e?.message || e);
    return res.status(500).json({ ok: false, error: e?.message || String(e) });
  }
}
