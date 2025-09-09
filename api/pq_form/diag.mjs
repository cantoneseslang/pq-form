import { getSheetsClient } from '../../../lib/sheets.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req) {
  try {
    const out = {};
    out.hasSheetId = !!process.env.PQFORM_SHEET_ID;
    out.saJsonLen = (process.env.GOOGLE_SA_JSON || '').length;
    try {
      const { sheetName, spreadsheetId } = getSheetsClient();
      out.sheetName = sheetName;
      out.spreadsheetId = spreadsheetId?.slice(0,8) + '...';
      out.clientOk = true;
    } catch (e) {
      out.clientOk = false;
      out.clientErr = e?.message || String(e);
    }
    return new Response(JSON.stringify({ success: true, diag: out }), { headers: { 'Cache-Control': 'no-store', 'Content-Type':'application/json' } });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: e?.message || String(e) }), { status: 500, headers: { 'Cache-Control': 'no-store', 'Content-Type':'application/json' } });
  }
}


