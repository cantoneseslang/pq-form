import { google } from 'googleapis';
import { Readable } from 'stream';
import { PQFORM_SPREADSHEET_ID } from './materialWidthSheets.js';

export const RAW_MATERIAL_QC_SHEET_ID = process.env.RAW_MATERIAL_QC_SHEET_ID
  || process.env.PQFORM_SHEET_ID
  || PQFORM_SPREADSHEET_ID;
export const RAW_MATERIAL_QC_SHEET_NAME = process.env.RAW_MATERIAL_QC_SHEET_NAME || 'RAW_MATERIAL_QC';
export const RAW_MATERIAL_QC_DRIVE_FOLDER_ID = process.env.RAW_MATERIAL_QC_DRIVE_FOLDER_ID || '';

const QC_HEADERS = [
  '記録日時',
  '入庫日',
  'ロット番号',
  '材料規格',
  '仕入先',
  '検査員',
  '基材',
  '塗装厚 (μm)',
  '写真 URL',
];

function getServiceAccount() {
  const raw = process.env.GOOGLE_SA_JSON || '';
  if (!raw) throw new Error('GOOGLE_SA_JSON is not set');
  const obj = JSON.parse(raw);
  if (obj?.private_key) obj.private_key = obj.private_key.replace(/\\n/g, '\n');
  return obj;
}

export function getMaterialQcClients() {
  const sa = getServiceAccount();
  const scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file',
  ];
  const auth = new google.auth.JWT(sa.client_email, undefined, sa.private_key, scopes);
  return {
    sheets: google.sheets({ version: 'v4', auth }),
    drive: google.drive({ version: 'v3', auth }),
    serviceAccountEmail: sa.client_email,
  };
}

function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

async function resolveSheetTab(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  const match = (res.data.sheets || []).find((s) => s.properties?.title === sheetName);
  return {
    title: match?.properties?.title || null,
    sheetId: match?.properties?.sheetId ?? null,
  };
}

async function ensureSheetTab(sheets, spreadsheetId, sheetName) {
  const existing = await resolveSheetTab(sheets, spreadsheetId, sheetName);
  if (existing.title) return { ...existing, created: false };

  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        addSheet: { properties: { title: sheetName } },
      }],
    },
  });

  const newProps = addRes.data.replies?.[0]?.addSheet?.properties;
  return {
    title: newProps?.title || sheetName,
    sheetId: newProps?.sheetId ?? null,
    created: true,
  };
}

async function ensureHeaderRow(sheets, spreadsheetId, sheetTitle) {
  const range = `${quoteSheetName(sheetTitle)}!A1:I1`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  const row = res.data.values?.[0];
  if (row && row.length >= QC_HEADERS.length) return;

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [QC_HEADERS] },
  });
}

function stripBase64Prefix(data) {
  return String(data ?? '').replace(/^data:image\/\w+;base64,/, '');
}

function sanitizeFilePart(value) {
  return String(value ?? 'unknown').trim().replace(/[^\w.-]+/g, '_').slice(0, 40) || 'unknown';
}

export async function uploadQcPhotoToDrive({ imageBase64, mimeType = 'image/jpeg', receiptDate, lotNo }) {
  const folderId = RAW_MATERIAL_QC_DRIVE_FOLDER_ID;
  if (!folderId) throw new Error('RAW_MATERIAL_QC_DRIVE_FOLDER_ID is not set');

  const { drive } = getMaterialQcClients();
  const data = stripBase64Prefix(imageBase64);
  const ext = mimeType.includes('png') ? 'png' : 'jpg';
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const fileName = `${sanitizeFilePart(receiptDate)}_${sanitizeFilePart(lotNo)}_${timestamp}.${ext}`;

  const fileRes = await drive.files.create({
    requestBody: {
      name: fileName,
      parents: [folderId],
      mimeType,
    },
    media: {
      mimeType,
      body: Readable.from(Buffer.from(data, 'base64')),
    },
    fields: 'id, webViewLink, webContentLink',
  });

  const fileId = fileRes.data.id;
  await drive.permissions.create({
    fileId,
    requestBody: { role: 'reader', type: 'anyone' },
  });

  return {
    fileId,
    photoUrl: fileRes.data.webViewLink || `https://drive.google.com/file/d/${fileId}/view`,
  };
}

export async function appendMaterialQcRecord(record) {
  const spreadsheetId = RAW_MATERIAL_QC_SHEET_ID;
  if (!spreadsheetId) throw new Error('RAW_MATERIAL_QC_SHEET_ID is not set');

  const { sheets } = getMaterialQcClients();
  const tab = await ensureSheetTab(sheets, spreadsheetId, RAW_MATERIAL_QC_SHEET_NAME);
  await ensureHeaderRow(sheets, spreadsheetId, tab.title);

  const now = new Date();
  const recordedAt = now.toLocaleString('zh-Hant', { timeZone: 'Asia/Hong_Kong', hour12: false });

  const row = [
    recordedAt,
    record.receiptDate || '',
    record.lotNo || '',
    record.materialSpec || '',
    record.supplier || '',
    record.inspector || '',
    record.substrate || '',
    record.thicknessUm != null ? record.thicknessUm : '',
    record.photoUrl || '',
  ];

  const appendRes = await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${quoteSheetName(tab.title)}!A:I`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] },
  });

  const updatedRange = appendRes.data.updates?.updatedRange || '';
  const rowMatch = updatedRange.match(/!A(\d+)/);
  const row = rowMatch ? parseInt(rowMatch[1], 10) : null;

  return {
    row,
    range: updatedRange,
    sheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${tab.sheetId ?? 0}`,
    photoUrl: record.photoUrl || null,
    tabCreated: tab.created,
  };
}

export async function submitMaterialQcRecord(record) {
  let photoUrl = '';
  if (record.imageBase64) {
    const uploaded = await uploadQcPhotoToDrive({
      imageBase64: record.imageBase64,
      mimeType: record.mimeType || 'image/jpeg',
      receiptDate: record.receiptDate,
      lotNo: record.lotNo,
    });
    photoUrl = uploaded.photoUrl;
  }

  const sheetResult = await appendMaterialQcRecord({ ...record, photoUrl });
  return { ...sheetResult, photoUrl };
}
