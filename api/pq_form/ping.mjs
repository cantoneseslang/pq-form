export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  try {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ ok: true, env: {
      sheetId: !!process.env.PQFORM_SHEET_ID,
      saLen: (process.env.GOOGLE_SA_JSON || '').length,
      sheetName: process.env.PQFORM_SHEET_NAME || '(default)'
    }});
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ ok: false, error: e?.message || String(e) });
  }
}


