import { google } from 'googleapis';

function getServiceAccount() {
  const raw = process.env.GOOGLE_SA_JSON || '';
  if (!raw) throw new Error('GOOGLE_SA_JSON is not set');
  try {
    const obj = JSON.parse(raw);
    if (obj && typeof obj.private_key === 'string') {
      obj.private_key = obj.private_key.replace(/\\n/g, '\n');
    }
    return obj;
  } catch (e) {
    throw new Error('GOOGLE_SA_JSON is invalid JSON');
  }
}

export function getDailyReportSheetsClient() {
  const sa = getServiceAccount();
  const scopes = ['https://www.googleapis.com/auth/spreadsheets'];
  const jwt = new google.auth.JWT(sa.client_email, undefined, sa.private_key, scopes);
  const sheets = google.sheets({ version: 'v4', auth: jwt });
  const spreadsheetId = process.env.PQFORM_DAILY_REPORT_SHEET_ID
    || '14R8GVayR_Uu6zx-yBVUTTBQib_rpJo22c_nz53oNbWw';
  if (!spreadsheetId) throw new Error('PQFORM_DAILY_REPORT_SHEET_ID is not set');
  return { sheets, spreadsheetId };
}

export async function readRange(rangeA1) {
  const { sheets, spreadsheetId } = getDailyReportSheetsClient();
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: rangeA1 });
  return res.data.values || [];
}

export async function writeRanges(updates) {
  const { sheets, spreadsheetId } = getDailyReportSheetsClient();
  const data = updates.map((u) => ({ range: u.range, values: [u.values] }));
  const res = await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data },
  });
  return res.data;
}
