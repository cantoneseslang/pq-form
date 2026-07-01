// Google Sheets helper for Vercel Functions
// Env: GOOGLE_SA_JSON (entire JSON), PQFORM_SHEET_ID, PQFORM_SHEET_NAME (default: "pq-form")

import { google } from 'googleapis';

function getServiceAccount() {
  const raw = process.env.GOOGLE_SA_JSON || '';
  if (!raw) throw new Error('GOOGLE_SA_JSON is not set');
  try {
    const obj = JSON.parse(raw);
    // Normalize private_key newlines when env stores as \n
    if (obj && typeof obj.private_key === 'string') {
      obj.private_key = obj.private_key.replace(/\\n/g, '\n');
    }
    return obj;
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

export const DATA_START_ROW = 8;
export const DATA_END_ROW = 50;
export const STYLED_TEMPLATE_END_ROW = 16;
export const DATA_ROW_HEIGHT_PX = 51;
export const NAME_MERGE_START_COL = 6; // G
export const NAME_MERGE_END_COL = 10; // J (exclusive)
export const DATA_COLUMN_COUNT = 20; // A..T
const CHECKBOX_COL_INDEXES = [11, 12, 13, 14, 15]; // L..P

export function dataBlockRange(sheetName) {
  return `${sheetName}!A${DATA_START_ROW}:T${DATA_END_ROW}`;
}

async function getSheetId(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  const sheet = (res.data.sheets || []).find((s) => s.properties?.title === sheetName);
  if (!sheet?.properties?.sheetId && sheet?.properties?.sheetId !== 0) {
    throw new Error(`sheet not found: ${sheetName}`);
  }
  return sheet.properties.sheetId;
}

export async function applyDataRowLayout(sheetName, targetRow) {
  const { sheets, spreadsheetId } = getSheetsClient();
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const rowIndex = targetRow - 1;
  const nameRange = {
    sheetId,
    startRowIndex: rowIndex,
    endRowIndex: rowIndex + 1,
    startColumnIndex: NAME_MERGE_START_COL,
    endColumnIndex: NAME_MERGE_END_COL,
  };

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        { unmergeCells: { range: nameRange } },
        { mergeCells: { range: nameRange, mergeType: 'MERGE_ALL' } },
        {
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex: rowIndex,
              endIndex: rowIndex + 1,
            },
            properties: { pixelSize: DATA_ROW_HEIGHT_PX },
            fields: 'pixelSize',
          },
        },
      ],
    },
  });
}

export async function copyRowTemplate(sheetName, targetRow, templateRow = DATA_START_ROW) {
  const { sheets, spreadsheetId } = getSheetsClient();
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        copyPaste: {
          source: {
            sheetId,
            startRowIndex: templateRow - 1,
            endRowIndex: templateRow,
            startColumnIndex: 0,
            endColumnIndex: DATA_COLUMN_COUNT,
          },
          destination: {
            sheetId,
            startRowIndex: targetRow - 1,
            endRowIndex: targetRow,
            startColumnIndex: 0,
            endColumnIndex: DATA_COLUMN_COUNT,
          },
          pasteType: 'PASTE_NORMAL',
          pasteOrientation: 'NORMAL',
        },
      }],
    },
  });
}

export function normalizeRowValues(values) {
  return values.map((value, index) => {
    if (!CHECKBOX_COL_INDEXES.includes(index)) return value ?? '';
    const text = String(value ?? '').trim();
    if (text === '✓' || text === 'TRUE' || text === 'true') return true;
    if (text === '✘' || text === 'FALSE' || text === 'false') return false;
    return false;
  });
}

export function findFirstEmptyRowInBlock(blockValues) {
  const rowCount = DATA_END_ROW - DATA_START_ROW + 1;
  for (let i = 0; i < rowCount; i++) {
    const row = blockValues[i] || [];
    const col = (idx) => (idx < row.length ? String(row[idx]).trim() : '');
    const c = col(2);   // C 產品編號
    const g = col(6);   // G 產品名稱
    const aToKEmpty = [0,1,2,3,4,5,6,7,8,9,10].every(j => (col(j) === ''));
    if ((c === '' && g === '') || aToKEmpty) {
      return DATA_START_ROW + i;
    }
  }
  return null;
}


