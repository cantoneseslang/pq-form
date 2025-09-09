// Google Sheets helper for Vercel Functions
// Env: GOOGLE_SA_JSON (entire JSON), PQFORM_SHEET_ID, PQFORM_SHEET_NAME (default: "pq-form")

import { google } from 'googleapis';

function getServiceAccount() {
  const raw = process.env.GOOGLE_SA_JSON || '';
  if (!raw) throw new Error('GOOGLE_SA_JSON is not set');
  try {
    return JSON.parse(raw);
  } catch (e) {
    throw new Error('GOOGLE_SA_JSON is invalid JSON');
  }
}

export function getSheetsClient() {
  const sa = getServiceAccount();
  const scopes = ['https://www.googleapis.com/auth/spreadsheets'];
  const jwt = new google.auth.JWT(sa.client_email, undefined, sa.private_key, scopes);
  const sheets = google.sheets({ version: 'v4', auth: jwt });
  const spreadsheetId = process.env.PQFORM_SHEET_ID;
  const sheetName = process.env.PQFORM_SHEET_NAME || 'pq-form';
  if (!spreadsheetId) throw new Error('PQFORM_SHEET_ID is not set');
  return { sheets, spreadsheetId, sheetName };
}

export async function readRange(rangeA1) {
  const { sheets, spreadsheetId } = getSheetsClient();
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: rangeA1 });
  return res.data.values || [];
}

export async function writeRowA1(rangeA1, values) {
  const { sheets, spreadsheetId } = getSheetsClient();
  const res = await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: rangeA1,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [values] }
  });
  return res.data;
}

export async function writeRanges(updates) {
  const { sheets, spreadsheetId } = getSheetsClient();
  const data = updates.map(u => ({ range: u.range, values: [u.values] }));
  const res = await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data }
  });
  return res.data;
}

export function findFirstEmptyRowInBlock(blockValues) {
  // blockValues corresponds to A8:T16 (9 rows max)
  // Return absolute row number (8..16). Empty if C and G both empty OR A..K all empty
  for (let i = 0; i < blockValues.length; i++) {
    const row = blockValues[i] || [];
    const col = (idx) => (idx < row.length ? String(row[idx]).trim() : '');
    const c = col(2);   // C 產品編號
    const g = col(6);   // G 產品名稱
    const aToKEmpty = [0,1,2,3,4,5,6,7,8,9,10].every(j => (col(j) === ''));
    if ((c === '' && g === '') || aToKEmpty) {
      return 8 + i; // absolute row number
    }
  }
  return null;
}


