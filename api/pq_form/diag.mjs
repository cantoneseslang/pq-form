import { getSheetsClient } from '../../lib/sheets.js';
import { getProductionOrderSheetsClient, listProductionOrderSheetTabs, resolveProductionOrderSheetName } from '../../lib/productionOrderSheets.js';

export const config = { runtime: 'nodejs' };

function getServiceAccountEmail() {
  try {
    const raw = process.env.GOOGLE_SA_JSON || '';
    if (!raw) return '';
    const obj = JSON.parse(raw);
    return String(obj?.client_email || '');
  } catch {
    return '';
  }
}

export default async function handler(req, res) {
  try {
    const out = {};
    out.hasSheetId = !!process.env.PQFORM_SHEET_ID;
    out.hasProductionOrderSheetId = !!process.env.PRODUCTION_ORDER_SHEET_ID;
    out.productionOrderSheetName = process.env.PRODUCTION_ORDER_SHEET_NAME || '202602146';
    out.saJsonLen = (process.env.GOOGLE_SA_JSON || '').length;
    out.serviceAccountEmail = getServiceAccountEmail();
    try {
      const { sheetName, spreadsheetId } = getSheetsClient();
      out.sheetName = sheetName;
      out.spreadsheetId = (spreadsheetId || '').slice(0,8) + '...';
      out.clientOk = true;
    } catch (e) {
      out.clientOk = false;
      out.clientErr = e?.message || String(e);
    }
    try {
      const { sheets, spreadsheetId, sheetName } = getProductionOrderSheetsClient();
      out.productionOrderSheetNameResolved = await resolveProductionOrderSheetName(sheets, spreadsheetId, sheetName);
      out.productionOrderSpreadsheetId = (spreadsheetId || '').slice(0, 8) + '...';
      out.productionOrderTabs = await listProductionOrderSheetTabs();
      out.productionOrderClientOk = true;
    } catch (e) {
      out.productionOrderClientOk = false;
      out.productionOrderClientErr = e?.message || String(e);
    }
    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ success: true, diag: out });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}


