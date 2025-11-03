// api/diag.js — 環境變數 / Token 診斷端點
export const config = { api: { bodyParser: false } };

function setCORS(res){
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

export default async function handler(req, res){
  setCORS(res);
  if (req.method === 'OPTIONS') return res.status(204).end();

  const has = k => !!(process.env[k] && String(process.env[k]).trim());

  const result = {
    envPresent: {
      TENANT_ID: has('TENANT_ID'),
      CLIENT_ID: has('CLIENT_ID'),
      CLIENT_SECRET: has('CLIENT_SECRET'),
      ONEDRIVE_USER_UPN: has('ONEDRIVE_USER_UPN'),
      ROOT_FOLDER: has('ROOT_FOLDER')
    },
    tokenOk: false,
    tokenError: null
  };

  try {
    if (has('TENANT_ID') && has('CLIENT_ID') && has('CLIENT_SECRET')) {
      const form = new URLSearchParams();
      form.append('grant_type', 'client_credentials');
      form.append('client_id', process.env.CLIENT_ID);
      form.append('client_secret', process.env.CLIENT_SECRET);
      form.append('scope', 'https://graph.microsoft.com/.default');

      const url = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
      const r = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: form
      });

      const raw = await r.text();
      let js = null; try { js = JSON.parse(raw); } catch {}

      if (r.ok && js?.access_token) {
        result.tokenOk = true;
      } else {
        result.tokenError = raw?.slice(0, 500);
      }
    } else {
      result.tokenError = 'Missing ENV';
    }
  } catch (e) {
    result.tokenError = e?.message || String(e);
  }

  return res.status(200).json(result);
}
