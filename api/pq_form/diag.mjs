import { getSheetsClient } from '../../../lib/sheets.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  try {
    const out = {};
    out.hasSheetId = !!process.env.PQFORM_SHEET_ID;
    out.saJsonLen = (process.env.GOOGLE_SA_JSON || '').length;
    try {
      const { sheetName, spreadsheetId } = getSheetsClient();
      out.sheetName = sheetName;
      out.spreadsheetId = (spreadsheetId || '').slice(0,8) + '...';
      out.clientOk = true;
    } catch (e) {
      out.clientOk = false;
      out.clientErr = e?.message || String(e);
    }
    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ success: true, diag: out });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}


