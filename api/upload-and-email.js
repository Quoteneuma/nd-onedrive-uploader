import formidable from "formidable";
import fs from "fs";
import { Resend } from "resend";

export const config = { api: { bodyParser: false } };

/* ---------------- CORS ---------------- */
function setCORS(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

/* --------------- ENV 取值 --------------- */
function need(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing ENV: ${name}`);
  return v;
}

/* --------------- 取 Graph Token --------------- */
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

/* --------------- 小工具 --------------- */
function firstVal(v) {
  if (Array.isArray(v)) return v[0];
  return v ?? "";
}

function firstFile(v) {
  if (Array.isArray(v)) return v[0];
  return v || null;
}

async function uploadBufferToOneDrive({ token, upn, root, subpath, filename, buffer }) {
  const cleanSubpath = String(subpath || "").replace(/^\/+/, "");
  const prefix = root ? `${root}/${cleanSubpath}`.replace(/\/+$/, "") : cleanSubpath;
  const drivePath = prefix ? `${prefix}/${filename}` : filename;

  const url =
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(upn)}` +
    `/drive/root:/${encodeURIComponent(drivePath).replace(/%2F/g, "/")}:/content`;

  const upr = await fetch(url, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}` },
    body: buffer
  });

  const upRaw = await upr.text();
  let upJs = null;
  try { upJs = JSON.parse(upRaw); } catch {}

  if (!upr.ok) {
    console.error("[UPLOAD_FAIL]", upr.status, upr.statusText, upRaw?.slice(0, 400));
    throw new Error(upJs?.error?.message || `Upload failed (${upr.status})`);
  }

  return upJs || null;
}

/* --------------- 主程式 --------------- */
export default async function handler(req, res) {
  setCORS(res);

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }
  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, error: "Use POST" });
  }

  try {
    const upn = need("ONEDRIVE_USER_UPN");
    const root = need("ROOT_FOLDER");
    const resendApiKey = need("RESEND_API_KEY");
    const fromEmail = need("QUOTE_FROM_EMAIL");

    const resend = new Resend(resendApiKey);

    const form = formidable({ multiples: false, keepExtensions: true });
    const { fields, files } = await new Promise((resolve, reject) => {
      form.parse(req, (err, flds, fls) => (err ? reject(err) : resolve({ fields: flds, files: fls })));
    });

    const pdfFile = firstFile(files?.pdf);
    const xlsxFile = firstFile(files?.xlsx);

    if (!pdfFile) {
      return res.status(400).json({ ok: false, error: "No pdf file" });
    }
    if (!xlsxFile) {
      return res.status(400).json({ ok: false, error: "No xlsx file" });
    }

    const pdfPath = pdfFile?.filepath || pdfFile?.path || null;
    const xlsxPath = xlsxFile?.filepath || xlsxFile?.path || null;

    if (!pdfPath || !xlsxPath) {
      return res.status(400).json({ ok: false, error: "Upload parse failed: no file path" });
    }

    const customerEmail = String(firstVal(fields?.customer_email)).trim();
    const customerName = String(firstVal(fields?.customer_name)).trim();
    const quotationNo = String(firstVal(fields?.quotation_no)).trim();
    const subpath = String(firstVal(fields?.subpath)).trim();
    const replyTo = String(firstVal(fields?.reply_to)).trim();

    if (!customerEmail) {
      return res.status(400).json({ ok: false, error: "Missing customer_email" });
    }

    const pdfName = String(firstVal(fields?.pdf_name) || pdfFile.originalFilename || "quote.pdf");
    const xlsxName = String(firstVal(fields?.xlsx_name) || xlsxFile.originalFilename || "quote.xlsx");

    const pdfBuffer = fs.readFileSync(pdfPath);
    const xlsxBuffer = fs.readFileSync(xlsxPath);

    const token = await getToken();

    const uploadedPdf = await uploadBufferToOneDrive({
      token,
      upn,
      root,
      subpath,
      filename: pdfName,
      buffer: pdfBuffer
    });

    const uploadedXlsx = await uploadBufferToOneDrive({
      token,
      upn,
      root,
      subpath,
      filename: xlsxName,
      buffer: xlsxBuffer
    });

    const subject = quotationNo
      ? `Your Quote Files - ${quotationNo}`
      : "Your Quote Files";

    const textLines = [
      `Hello ${customerName || "Customer"},`,
      ``,
      `Attached are your quote files.`,
      quotationNo ? `Quote No: ${quotationNo}` : ``,
      ``,
      `Please find the PDF and Excel files attached.`,
      ``,
      `Thank you.`
    ].filter(Boolean);

    const emailPayload = {
      from: fromEmail,
      to: customerEmail,
      subject,
      text: textLines.join("\n"),
      attachments: [
        {
          filename: pdfName,
          content: pdfBuffer
        },
        {
          filename: xlsxName,
          content: xlsxBuffer
        }
      ]
    };

    if (replyTo) {
      emailPayload.replyTo = replyTo;
    }

    const emailResult = await resend.emails.send(emailPayload);

if (emailResult?.error) {
  console.error("[RESEND_ERROR]", emailResult.error);
  return res.status(500).json({
    ok: false,
    error: "EMAIL_SEND_FAILED",
    resend_error: emailResult.error
  });
}

return res.status(200).json({
  ok: true,
  uploaded_pdf: uploadedPdf,
  uploaded_xlsx: uploadedXlsx,
  email_id: emailResult?.data?.id || "",
  email_result: emailResult
});
  } catch (e) {
    console.error("[SERVER_ERROR]", e?.message || e);
    return res.status(500).json({ ok: false, error: e?.message || String(e) });
  }
}
