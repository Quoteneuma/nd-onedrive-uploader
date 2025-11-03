// Vercel Serverless Function: /api/upload
// 解析 multipart，將 PDF / XLSX 上傳到 OneDrive/QuoteNeuma/...，並寫 metadata.json
import Busboy from "busboy";

const ROOT = process.env.ROOT_FOLDER || "QuoteNeuma";
const UPN  = process.env.ONEDRIVE_USER_UPN; // 例如 marketing@nanyaplastics-usa.com

function pad2(n){ return String(n).padStart(2,"0"); }
function safe(s){ return String(s||"").toLowerCase().replace(/[^a-z0-9._-]+/g,"-").replace(/-+/g,"-"); }

async function getAccessToken(){
  const resp = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials"
    })
  });
  const json = await resp.json();
  if (!json.access_token) throw new Error("Get token failed: " + JSON.stringify(json));
  return json.access_token;
}

// <=4MB 直接 PUT；>4MB 走 Upload Session 分段
async function uploadToOneDrive({ token, upn, path, buffer }){
  const base = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}/drive/root:/${encodeURI(path)}`;
  if (buffer.byteLength <= 4*1024*1024) {
    const r = await fetch(`${base}:/content`, {
      method: "PUT",
      headers: { Authorization: `Bearer ${token}` },
      body: buffer
    });
    if (!r.ok) throw new Error(`PUT ${r.status}: ${await r.text()}`);
    return await r.json();
  }
  const r1 = await fetch(`${base}:/createUploadSession`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } })
  });
  const sess = await r1.json();
  if (!sess.uploadUrl) throw new Error("Create session failed: " + JSON.stringify(sess));

  const url = sess.uploadUrl;
  const chunk = 5*1024*1024;
  let start = 0;
  while (start < buffer.byteLength) {
    const end = Math.min(start + chunk, buffer.byteLength);
    const slice = buffer.slice(start, end);
    const r = await fetch(url, {
      method: "PUT",
      headers: {
        "Content-Length": String(slice.byteLength),
        "Content-Range": `bytes ${start}-${end-1}/${buffer.byteLength}`
      },
      body: slice
    });
    if (!(r.ok || r.status === 202)) throw new Error(`Chunk ${start}-${end} failed: ${r.status} ${await r.text()}`);
    start = end;
  }
  return { ok: true };
}

// 解析 multipart（支援 pdf/xlsx 兩個欄位）
function parseMultipart(req){
  return new Promise((resolve, reject) => {
    const bb = Busboy({ headers: req.headers });
    const fields = {};
    const files = {}; // { pdf: {filename, mime, data:Buffer}, xlsx: {...} }
    bb.on("file", (name, file, info) => {
      const { filename, mimeType } = info;
      const chunks = [];
      file.on("data", d => chunks.push(d));
      file.on("end", () => { files[name] = { filename, mime: mimeType, data: Buffer.concat(chunks) }; });
    });
    bb.on("field", (name, val) => { fields[name] = val; });
    bb.on("error", reject);
    bb.on("finish", () => resolve({ fields, files }));
    req.pipe(bb);
  });
}

export default async function handler(req, res){
  try{
    if (req.method !== "POST") {
      res.status(405).json({ ok:false, error:"Use POST" });
      return;
    }
    if (!UPN) {
      res.status(500).json({ ok:false, error:"Missing ONEDRIVE_USER_UPN" });
      return;
    }

    const { fields, files } = await parseMultipart(req);
    const token = await getAccessToken();

    const serial = fields.serial || "GUEST-0000";
    const customerEmail = fields.customerEmail || "guest";
    const pageUrl = fields.pageUrl || "";
    const userAgent = fields.userAgent || "";
    let cart = {};
    try { cart = JSON.parse(fields.cartJson || "{}"); } catch {}

    const now = new Date();
    const yyyy = now.getFullYear();
    const mmdd = `${pad2(now.getMonth()+1)}${pad2(now.getDate())}`;
    const baseDir = `${ROOT}/${safe(customerEmail)}/${yyyy}/${mmdd}/${safe(serial)}`;

    const out = [];
    for (const name of ["pdf","xlsx"]) {
      const f = files[name];
      if (!f) continue;
      const fname = f.filename || `ND-${serial}.${name}`;
      const path = `${baseDir}/${fname}`;
      await uploadToOneDrive({ token, upn: UPN, path, buffer: f.data });
      out.push({ path, filename: fname, size: f.data.length, mime: f.mime });
    }

    const meta = {
      serial,
      customerEmail,
      when: now.toISOString(),
      pageUrl,
      userAgent,
      files: out,
      cart,
      version: "v1"
    };
    await uploadToOneDrive({
      token, upn: UPN, path: `${baseDir}/metadata.json`, buffer: Buffer.from(JSON.stringify(meta, null, 2))
    });

    res.status(200).json({ ok:true, folder: baseDir, files: out, meta: `${baseDir}/metadata.json` });
  }catch(e){
    res.status(500).json({ ok:false, error: String(e?.message || e) });
  }
}
