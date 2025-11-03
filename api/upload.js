// api/upload.js
// 使用環境變數：TENANT_ID, CLIENT_ID, CLIENT_SECRET, ONEDRIVE_USER_UPN, ROOT_FOLDER
// 支援大檔 (Upload Session)；自動建立 ROOT_FOLDER 與 subpath 資料夾

import formidable from "formidable";
import fs from "fs";

export const config = { api: { bodyParser: false } };

function mustEnv(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing ENV: ${name}`);
  return v;
}

// 逐段編碼 path，不編碼斜線
function encodePath(p = "") {
  return String(p)
    .split("/")
    .filter(Boolean)
    .map(encodeURIComponent)
    .join("/");
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
  try { js = JSON.parse(raw); } catch {}

  if (!r.ok || !js?.access_token) {
    throw new Error(`Token failed (${r.status}) ${raw?.slice(0, 200)}`);
  }
  return js.access_token;
}

async function graph(url, token, opts = {}) {
  const r = await fetch(url, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(opts.headers || {}),
    },
  });
  const text = await r.text();
  let data = null;
  try { data = JSON.parse(text); } catch {}
  if (!r.ok) {
    const msg = data?.error?.message || text?.slice(0, 400);
    throw new Error(`Graph ${r.status}: ${msg}`);
  }
  return data ?? text;
}

// 確保 /root:/A/B/C 存在（逐層建立資料夾）
async function ensureFolders(driveBase, token, fullFolderPath) {
  const parts = String(fullFolderPath || "")
    .split("/")
    .filter(Boolean);
  if (!parts.length) return ""; // 根目錄

  let built = "";
  for (const part of parts) {
    built = built ? `${built}/${part}` : part;

    // 檢查是否存在
    const pathUrl = `${driveBase}/root:/${encodePath(built)}`;
    const exists = await fetch(pathUrl, { headers: { Authorization: `Bearer ${token}` } });
    if (exists.ok) continue;

    // 建立於 parent 的 children
    const parent = built.split("/").slice(0, -1).join("/");
    const childrenUrl = parent
      ? `${driveBase}/root:/${encodePath(parent)}:/children`
      : `${driveBase}/root/children`;

    await graph(childrenUrl, token, {
      method: "POST",
      body: JSON.stringify({
        name: part,
        folder: {},
        "@microsoft.graph.conflictBehavior": "replace",
      }),
    });
  }
  return fullFolderPath;
}

// 大檔分段上傳 (Upload Session)
async function uploadViaSession(driveBase, token, targetPath, buffer) {
  const createUrl = `${driveBase}/root:/${encodePath(targetPath)}:/createUploadSession`;
  const session = await graph(createUrl, token, {
    method: "POST",
    body: JSON.stringify({
      item: {
        "@microsoft.graph.conflictBehavior": "replace",
        name: targetPath.split("/").pop(),
      },
    }),
  });
  const uploadUrl = session?.uploadUrl;
  if (!uploadUrl) throw new Error("No uploadUrl from createUploadSession");

  const chunk = 5 * 1024 * 1024; // 5MB
  const size = buffer.length;
  let start = 0;
  while (start < size) {
    const end = Math.min(start + chunk, size) - 1;
    const slice = buffer.slice(start, end + 1);
    const r = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(slice.length),
        "Content-Range": `bytes ${start}-${end}/${size}`,
      },
      body: slice,
    });

    // 202/201/200 都可能出現；完成時會回傳檔案資訊
    if (r.status >= 400) {
      const t = await r.text();
      throw new Error(`Chunk upload failed (${r.status}): ${t?.slice(0, 200)}`);
    }

    // 完成時會回傳 item
    if (end + 1 === size) {
      const t = await r.text();
      try { return JSON.parse(t); } catch { return { ok: true, raw: t }; }
    }
    start = end + 1;
  }
  throw new Error("Unexpected end of upload loop");
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(400).json({ ok: false, error: "Use POST" });
  }

  try {
    // 解析 multipart
    const form = formidable({ multiples: false });
    const { fields, files } = await new Promise((resolve, reject) => {
      form.parse(req, (err, flds, fls) => (err ? reject(err) : resolve({ fields: flds, files: fls })));
    });

    if (!files?.file) {
      return res.status(400).json({ ok: false, error: "No file uploaded" });
    }

    const token = await getToken();
    const upn = mustEnv("ONEDRIVE_USER_UPN");
    const rootFolder = (process.env.ROOT_FOLDER || "").trim();

    const driveBase = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}/drive`;

    // 組目錄：ROOT_FOLDER / subpath
    const subpath = String(fields.subpath || "").replace(/^\/+/, "");
    const folderPath = [rootFolder, subpath].filter(Boolean).join("/");

    if (folderPath) {
      await ensureFolders(driveBase, token, folderPath);
    }

    // 檔名
    const filename =
      String(fields.filename || files.file.originalFilename || "file.bin").trim() || "file.bin";

    // 讀檔 & 上傳
    const buffer = fs.readFileSync(files.file.filepath);
    const target = [folderPath, filename].filter(Boolean).join("/");

    const item = await uploadViaSession(driveBase, token, target, buffer);

    return res.status(200).json({ ok: true, item });
  } catch (e) {
    return res.status(500).json({ ok: false, error: e?.message || String(e) });
  }
}
