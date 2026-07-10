import { google } from 'googleapis';
import { Readable } from 'stream';
import { PQFORM_SPREADSHEET_ID } from './materialWidthSheets.js';
import { uploadQcPhotoToSupabase } from './materialQcPhotos.js';

export const RAW_MATERIAL_QC_SHEET_ID = process.env.RAW_MATERIAL_QC_SHEET_ID
  || process.env.PQFORM_SHEET_ID
  || PQFORM_SPREADSHEET_ID;
export const RAW_MATERIAL_QC_SHEET_NAME = process.env.RAW_MATERIAL_QC_SHEET_NAME || 'RAW_MATERIAL_QC';
export const RAW_MATERIAL_QC_DRIVE_FOLDER_ID = String(process.env.RAW_MATERIAL_QC_DRIVE_FOLDER_ID || '').trim();

const QC_HEADERS = [
  '記錄日期時間',
  '入倉日期',
  '單號',
  '材料厚度',
  '材料闊度 (mm)',
  '供貨商',
  '檢查員',
  '基材',
  '塗裝厚度 (μm)',
  '基準 (μm)',
  '判定',
  '相片 URL',
];

function getServiceAccount() {
  const raw = process.env.GOOGLE_SA_JSON || '';
  if (!raw) throw new Error('GOOGLE_SA_JSON is not set');
  const obj = JSON.parse(raw);
  if (obj?.private_key) obj.private_key = obj.private_key.replace(/\\n/g, '\n');
  return obj;
}

export function getMaterialQcDriveClient() {
  const sa = getServiceAccount();
  const scopes = ['https://www.googleapis.com/auth/drive'];
  const delegatedUser = String(process.env.GOOGLE_DRIVE_DELEGATED_USER || '').trim();
  const auth = new google.auth.JWT(
    sa.client_email,
    undefined,
    sa.private_key,
    scopes,
    delegatedUser || undefined,
  );
  return {
    drive: google.drive({ version: 'v3', auth }),
    serviceAccountEmail: sa.client_email,
    delegatedUser: delegatedUser || null,
  };
}

export function getMaterialQcClients() {
  const sa = getServiceAccount();
  const scopes = ['https://www.googleapis.com/auth/spreadsheets'];
  const auth = new google.auth.JWT(sa.client_email, undefined, sa.private_key, scopes);
  return {
    sheets: google.sheets({ version: 'v4', auth }),
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
  const range = `${quoteSheetName(sheetTitle)}!A1:L1`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  const existing = res.data.values?.[0] || [];
  const matches = existing.length >= QC_HEADERS.length
    && QC_HEADERS.every((h, i) => existing[i] === h);
  if (matches) return;

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

function normalizeMimeType(mimeType) {
  const raw = String(mimeType || 'image/jpeg').split(';')[0].trim().toLowerCase();
  if (raw === 'image/png') return 'image/png';
  if (raw === 'image/webp') return 'image/webp';
  return 'image/jpeg';
}

export async function uploadQcPhotoToDrive({ imageBase64, mimeType = 'image/jpeg', receiptDate, lotNo }) {
  const folderId = RAW_MATERIAL_QC_DRIVE_FOLDER_ID;
  if (!folderId) throw new Error('RAW_MATERIAL_QC_DRIVE_FOLDER_ID is not set');

  const { drive, serviceAccountEmail, delegatedUser } = getMaterialQcDriveClient();
  const data = stripBase64Prefix(imageBase64);
  if (!data) throw new Error('imageBase64 is empty');

  const safeMime = normalizeMimeType(mimeType);
  const ext = safeMime === 'image/png' ? 'png' : safeMime === 'image/webp' ? 'webp' : 'jpg';
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const fileName = `${sanitizeFilePart(receiptDate)}_${sanitizeFilePart(lotNo)}_${timestamp}.${ext}`;
  const buffer = Buffer.from(data, 'base64');

  let fileRes;
  try {
    fileRes = await drive.files.create({
      requestBody: {
        name: fileName,
        parents: [folderId],
      },
      media: {
        mimeType: safeMime,
        body: Readable.from(buffer),
      },
      fields: 'id, webViewLink, webContentLink',
      supportsAllDrives: true,
    });
  } catch (e) {
    const msg = e?.message || String(e);
    const hint = delegatedUser
      ? `Delegated user: ${delegatedUser}.`
      : 'Service accounts cannot upload to personal Drive. Set GOOGLE_DRIVE_DELEGATED_USER, use a Shared Drive folder, or use Supabase Storage (default).';
    throw new Error(`Drive upload failed (${msg}). ${hint}`);
  }

  const fileId = fileRes.data.id;
  if (!fileId) throw new Error('Drive upload returned no file id');

  try {
    await drive.permissions.create({
      fileId,
      requestBody: { role: 'reader', type: 'anyone' },
      supportsAllDrives: true,
    });
  } catch {
    // 共有設定に失敗してもファイル自体は保存済みなので続行
  }

  return {
    fileId,
    photoUrl: fileRes.data.webViewLink || `https://drive.google.com/file/d/${fileId}/view`,
  };
}

async function uploadQcPhoto({ imageBase64, mimeType, receiptDate, lotNo }) {
  // Supabase Storage を優先（サービスアカウントの Drive 容量制限を回避）
  try {
    return await uploadQcPhotoToSupabase({ imageBase64, mimeType, receiptDate, lotNo });
  } catch (supabaseError) {
    if (!RAW_MATERIAL_QC_DRIVE_FOLDER_ID) throw supabaseError;
    try {
      return await uploadQcPhotoToDrive({ imageBase64, mimeType, receiptDate, lotNo });
    } catch (driveError) {
      throw new Error(
        `Photo upload failed. Supabase: ${supabaseError.message}. Drive: ${driveError.message}`,
      );
    }
  }
}

export async function appendMaterialQcRecord(record) {
  const spreadsheetId = RAW_MATERIAL_QC_SHEET_ID;
  if (!spreadsheetId) throw new Error('RAW_MATERIAL_QC_SHEET_ID is not set');

  const { sheets } = getMaterialQcClients();
  const tab = await ensureSheetTab(sheets, spreadsheetId, RAW_MATERIAL_QC_SHEET_NAME);
  await ensureHeaderRow(sheets, spreadsheetId, tab.title);

  const now = new Date();
  const recordedAt = now.toLocaleString('zh-Hant', { timeZone: 'Asia/Hong_Kong', hour12: false });

  const values = [
    recordedAt,
    record.receiptDate || '',
    record.lotNo || '',
    record.materialThickness || '',
    record.materialWidth || '',
    record.supplier || '',
    record.inspector || '',
    record.substrate || '',
    record.thicknessUm != null ? record.thicknessUm : '',
    record.standardUm != null ? record.standardUm : '',
    record.judgment || '',
    record.photoUrl || '',
  ];

  const appendRes = await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${quoteSheetName(tab.title)}!A:L`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [values] },
  });

  const updatedRange = appendRes.data.updates?.updatedRange || '';
  const rowMatch = updatedRange.match(/!A(\d+)/);
  const rowNumber = rowMatch ? parseInt(rowMatch[1], 10) : null;

  return {
    row: rowNumber,
    range: updatedRange,
    sheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${tab.sheetId ?? 0}`,
    photoUrl: record.photoUrl || null,
    tabCreated: tab.created,
  };
}

export async function submitMaterialQcRecord(record) {
  if (!record.imageBase64) {
    throw new Error('写真データがありません。再撮影してから記録してください。');
  }

  const uploaded = await uploadQcPhoto({
    imageBase64: record.imageBase64,
    mimeType: record.mimeType || 'image/jpeg',
    receiptDate: record.receiptDate,
    lotNo: record.lotNo,
  });

  const sheetResult = await appendMaterialQcRecord({ ...record, photoUrl: uploaded.photoUrl });
  return {
    ...sheetResult,
    photoUrl: uploaded.photoUrl,
  };
}
