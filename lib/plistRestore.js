import { google } from 'googleapis';
import * as XLSX from 'xlsx';
import {
  PQFORM_SPREADSHEET_ID,
  PLIST_SHEET_NAME,
  getMaterialWidthSheetsClient,
} from './materialWidthSheets.js';

const DRIVE_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.readonly',
];

function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function getClients() {
  const raw = process.env.GOOGLE_SA_JSON || '';
  if (!raw) throw new Error('GOOGLE_SA_JSON is not set');
  const sa = JSON.parse(raw);
  if (sa?.private_key) sa.private_key = sa.private_key.replace(/\\n/g, '\n');
  const jwt = new google.auth.JWT(sa.client_email, undefined, sa.private_key, DRIVE_SCOPES);
  return {
    serviceAccountEmail: sa.client_email,
    auth: jwt,
    sheets: google.sheets({ version: 'v4', auth: jwt }),
    driveV2: google.drive({ version: 'v2', auth: jwt }),
    driveV3: google.drive({ version: 'v3', auth: jwt }),
  };
}

function isValidPlistHeader(row) {
  const a = String(row?.[0] ?? '').trim().toUpperCase();
  const b = String(row?.[1] ?? '').trim().toUpperCase();
  return a === 'PRODCODE' || b.includes('PQ-FORM-DESC') || b.includes('PQ-FORM');
}

function findPlistSheetName(workbook) {
  if (workbook.SheetNames.includes(PLIST_SHEET_NAME)) return PLIST_SHEET_NAME;
  const fuzzy = workbook.SheetNames.find((n) => /plist/i.test(n));
  return fuzzy || null;
}

function sheetToRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
}

async function downloadRevisionXlsx(auth, fileId, revisionId) {
  const token = await auth.getAccessToken();
  const accessToken = token?.token;
  if (!accessToken) throw new Error('Failed to obtain access token');

  const url = `https://docs.google.com/spreadsheets/export?id=${encodeURIComponent(fileId)}&revision=${encodeURIComponent(revisionId)}&exportFormat=xlsx`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Revision export failed (${res.status}): ${text.slice(0, 200)}`);
  }
  const buf = Buffer.from(await res.arrayBuffer());
  return XLSX.read(buf, { type: 'buffer' });
}

export async function listSpreadsheetRevisions({ limit = 30 } = {}) {
  const { driveV3, serviceAccountEmail } = getClients();
  const res = await driveV3.revisions.list({
    fileId: PQFORM_SPREADSHEET_ID,
    pageSize: limit,
    fields: 'revisions(id,modifiedTime,lastModifyingUser/displayName)',
  });
  const revisions = (res.data.revisions || []).map((r) => ({
    id: r.id,
    modifiedTime: r.modifiedTime,
    user: r.lastModifyingUser?.displayName || '',
  }));
  return { serviceAccountEmail, spreadsheetId: PQFORM_SPREADSHEET_ID, revisions };
}

export async function previewPlistRevision(revisionId) {
  const { auth, serviceAccountEmail } = getClients();
  const workbook = await downloadRevisionXlsx(auth, PQFORM_SPREADSHEET_ID, revisionId);
  const sheetName = findPlistSheetName(workbook);
  if (!sheetName) {
    return {
      serviceAccountEmail,
      revisionId,
      ok: false,
      error: `Sheet ${PLIST_SHEET_NAME} not found in revision`,
      sheetNames: workbook.SheetNames,
    };
  }
  const rows = sheetToRows(workbook, sheetName);
  const header = rows[0] || [];
  const dataRows = rows.slice(1).filter((row) => String(row?.[0] ?? '').trim());
  return {
    serviceAccountEmail,
    revisionId,
    ok: isValidPlistHeader(header),
    sheetName,
    header,
    rowCount: dataRows.length,
    preview: rows.slice(0, 5),
  };
}

export async function findLatestGoodPlistRevision({ beforeIso = '2026-07-05T05:56:00.000Z' } = {}) {
  const { revisions } = await listSpreadsheetRevisions({ limit: 50 });
  const cutoff = new Date(beforeIso).getTime();
  const candidates = revisions
    .filter((r) => new Date(r.modifiedTime).getTime() < cutoff)
    .sort((a, b) => new Date(b.modifiedTime) - new Date(a.modifiedTime));

  const tried = [];
  for (const rev of candidates) {
    try {
      const preview = await previewPlistRevision(rev.id);
      tried.push({
        id: rev.id,
        modifiedTime: rev.modifiedTime,
        ok: preview.ok,
        rowCount: preview.rowCount,
        header: preview.header,
      });
      if (preview.ok && preview.rowCount >= 100) {
        return { revision: rev, preview, tried };
      }
    } catch (e) {
      tried.push({ id: rev.id, modifiedTime: rev.modifiedTime, ok: false, error: e.message });
    }
  }
  return { revision: null, preview: null, tried };
}

export async function restorePlistFromRevision({ revisionId, dryRun = true } = {}) {
  const { auth, sheets, serviceAccountEmail } = getClients();
  let targetRevisionId = revisionId;

  if (!targetRevisionId) {
    const found = await findLatestGoodPlistRevision();
    if (!found.revision) {
      return {
        dryRun,
        restored: false,
        serviceAccountEmail,
        tried: found.tried,
        error: 'No valid plist revision found before incident time',
      };
    }
    targetRevisionId = found.revision.id;
  }

  const preview = await previewPlistRevision(targetRevisionId);
  if (!preview.ok) {
    return {
      dryRun,
      restored: false,
      serviceAccountEmail,
      revisionId: targetRevisionId,
      preview,
      error: 'Selected revision does not contain a valid PQ-Form-plist header',
    };
  }

  const workbook = await downloadRevisionXlsx(auth, PQFORM_SPREADSHEET_ID, targetRevisionId);
  const sheetName = findPlistSheetName(workbook);
  const rows = sheetToRows(workbook, sheetName);
  const normalized = rows.map((row) => {
    const out = row.slice(0, 9);
    while (out.length < 9) out.push('');
    return out;
  });

  if (dryRun) {
    return {
      dryRun: true,
      restored: false,
      serviceAccountEmail,
      revisionId: targetRevisionId,
      sheetName,
      rowCount: normalized.length - 1,
      header: normalized[0],
      preview: normalized.slice(0, 5),
    };
  }

  const quoted = quoteSheetName(PLIST_SHEET_NAME);
  await sheets.spreadsheets.values.clear({
    spreadsheetId: PQFORM_SPREADSHEET_ID,
    range: `${quoted}!A:I`,
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: PQFORM_SPREADSHEET_ID,
    range: `${quoted}!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: normalized },
  });

  const verify = await sheets.spreadsheets.values.get({
    spreadsheetId: PQFORM_SPREADSHEET_ID,
    range: `${quoted}!A1:I5`,
  });

  return {
    dryRun: false,
    restored: true,
    serviceAccountEmail,
    revisionId: targetRevisionId,
    sheetName,
    rowCount: normalized.length - 1,
    header: normalized[0],
    verify: verify.data.values || [],
  };
}
