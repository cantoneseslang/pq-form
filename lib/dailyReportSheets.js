import { google } from 'googleapis';
import XLSX from 'xlsx';
import { resolveDailyReportSpreadsheetId } from './dailyReportSheetMap.js';

const NATIVE_GOOGLE_SHEET = 'application/vnd.google-apps.spreadsheet';
const OFFICE_MIMES = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
]);

const officeWorkbookCache = new Map();
const fileMimeCache = new Map();

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

function getAuthClient() {
  const sa = getServiceAccount();
  const scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly',
  ];
  return new google.auth.JWT(sa.client_email, undefined, sa.private_key, scopes);
}

function getSheetsApi() {
  return google.sheets({ version: 'v4', auth: getAuthClient() });
}

function getDriveApi() {
  return google.drive({ version: 'v3', auth: getAuthClient() });
}

function formatCellValue(cell) {
  if (!cell) return '';
  if (cell.w != null && String(cell.w).trim() !== '') return String(cell.w).trim();
  if (cell.v == null) return '';
  if (cell.t === 'd' && cell.v instanceof Date) {
    const d = cell.v;
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  }
  return String(cell.v).trim();
}

function splitRangeA1(rangeA1) {
  const bang = String(rangeA1).indexOf('!');
  const sheetName = bang >= 0 ? rangeA1.slice(0, bang) : 'Sheet1';
  const a1 = bang >= 0 ? rangeA1.slice(bang + 1) : rangeA1;
  const ref = XLSX.utils.decode_range(a1.includes(':') ? a1 : `${a1}:${a1}`);
  return { sheetName, ref };
}

function readRangeFromWorkbook(workbook, rangeA1) {
  const { sheetName, ref } = splitRangeA1(rangeA1);
  const ws = workbook.Sheets[sheetName];
  if (!ws) {
    const names = workbook.SheetNames || [];
    throw new Error(`Sheet "${sheetName}" not found. Available: ${names.join(', ')}`);
  }

  const rows = [];
  for (let r = ref.s.r; r <= ref.e.r; r += 1) {
    const row = [];
    for (let c = ref.s.c; c <= ref.e.c; c += 1) {
      row.push(formatCellValue(ws[XLSX.utils.encode_cell({ r, c })]));
    }
    rows.push(row);
  }
  return rows;
}

async function getFileMimeType(sourceId) {
  if (fileMimeCache.has(sourceId)) return fileMimeCache.get(sourceId);
  const drive = getDriveApi();
  const { data } = await drive.files.get({
    fileId: sourceId,
    fields: 'mimeType',
    supportsAllDrives: true,
  });
  fileMimeCache.set(sourceId, data.mimeType || '');
  return data.mimeType || '';
}

async function loadOfficeWorkbook(sourceId) {
  if (officeWorkbookCache.has(sourceId)) return officeWorkbookCache.get(sourceId);

  const mimeType = await getFileMimeType(sourceId);
  if (!OFFICE_MIMES.has(mimeType)) {
    throw new Error(`Expected Excel daily report file, got: ${mimeType || 'unknown'}`);
  }

  const drive = getDriveApi();
  const res = await drive.files.get(
    { fileId: sourceId, alt: 'media', supportsAllDrives: true },
    { responseType: 'arraybuffer' },
  );
  const workbook = XLSX.read(res.data, { type: 'array', cellDates: true });
  officeWorkbookCache.set(sourceId, workbook);
  return workbook;
}

export async function getDailyReportSheetsClient(month, year) {
  const sourceSpreadsheetId = resolveDailyReportSpreadsheetId(month, year);
  const mimeType = await getFileMimeType(sourceSpreadsheetId);
  const isNative = mimeType === NATIVE_GOOGLE_SHEET;
  return {
    sheets: isNative ? getSheetsApi() : null,
    spreadsheetId: sourceSpreadsheetId,
    sourceSpreadsheetId,
    mimeType,
    isNative,
  };
}

export async function readRange(rangeA1, month, year) {
  const sourceSpreadsheetId = resolveDailyReportSpreadsheetId(month, year);
  const mimeType = await getFileMimeType(sourceSpreadsheetId);

  if (mimeType === NATIVE_GOOGLE_SHEET) {
    const sheets = getSheetsApi();
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: sourceSpreadsheetId,
      range: rangeA1,
    });
    return res.data.values || [];
  }

  if (OFFICE_MIMES.has(mimeType)) {
    const workbook = await loadOfficeWorkbook(sourceSpreadsheetId);
    return readRangeFromWorkbook(workbook, rangeA1);
  }

  throw new Error(`Unsupported daily report file type: ${mimeType || 'unknown'}`);
}

export async function writeRanges(updates, month, year) {
  const sourceSpreadsheetId = resolveDailyReportSpreadsheetId(month, year);
  const mimeType = await getFileMimeType(sourceSpreadsheetId);
  if (mimeType !== NATIVE_GOOGLE_SHEET) {
    throw new Error('Daily report write is only supported for native Google Sheets files');
  }

  const sheets = getSheetsApi();
  const data = updates.map((u) => ({ range: u.range, values: [u.values] }));
  const res = await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: sourceSpreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data },
  });
  return res.data;
}

export async function listDailyReportTabNames(month, year) {
  const sourceSpreadsheetId = resolveDailyReportSpreadsheetId(month, year);
  const mimeType = await getFileMimeType(sourceSpreadsheetId);

  if (mimeType === NATIVE_GOOGLE_SHEET) {
    const sheets = getSheetsApi();
    const res = await sheets.spreadsheets.get({
      spreadsheetId: sourceSpreadsheetId,
      fields: 'sheets.properties.title',
    });
    return (res.data.sheets || [])
      .map((sheet) => sheet.properties?.title)
      .filter(Boolean);
  }

  if (OFFICE_MIMES.has(mimeType)) {
    const workbook = await loadOfficeWorkbook(sourceSpreadsheetId);
    return workbook.SheetNames || [];
  }

  return [];
}

export async function describeDailyReportSpreadsheet(month, year) {
  const sourceSpreadsheetId = resolveDailyReportSpreadsheetId(month, year);
  const mimeType = await getFileMimeType(sourceSpreadsheetId);
  const info = {
    year: String(year ?? ''),
    month: String(month ?? ''),
    sourceSpreadsheetId,
    mimeType,
    isNative: mimeType === NATIVE_GOOGLE_SHEET,
    sheetNames: await listDailyReportTabNames(month, year),
  };
  return info;
}
