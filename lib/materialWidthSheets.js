import { google } from 'googleapis';
import { thicknessForProductLookup } from './plist.js';

export const PQFORM_SPREADSHEET_ID = process.env.PQFORM_SHEET_ID || '1u_fsEVAumMySLx8fZdMP5M4jgHiGG6ncPjFEXSXHQ1M';
export const PLIST_SHEET_NAME = process.env.PQFORM_PLIST_SHEET_NAME || 'PQ-Form-plist';
export const EOFFICE_SHEET_CANDIDATES = ['e-office-material width', 'e-office-material width '];
export const MONTHLY_DES_SPREADSHEET_ID = process.env.MATERIAL_WIDTH_DES_SHEET_ID || '1R-xjzmki0pzMlXJzhzVee_y6XWqFbrUSlSL04osBhbc';
export const MONTHLY_DES_SHEET_NAME = process.env.MATERIAL_WIDTH_DES_SHEET_NAME || 'Monthly Des.';
export const MONTHLY_DES_GID = Number(process.env.MATERIAL_WIDTH_DES_SHEET_GID || 2089764826);

function getServiceAccount() {
  const raw = process.env.GOOGLE_SA_JSON || '';
  if (!raw) throw new Error('GOOGLE_SA_JSON is not set');
  const obj = JSON.parse(raw);
  if (obj?.private_key) obj.private_key = obj.private_key.replace(/\\n/g, '\n');
  return obj;
}

export function getMaterialWidthSheetsClient(scopes = ['https://www.googleapis.com/auth/spreadsheets']) {
  const sa = getServiceAccount();
  const jwt = new google.auth.JWT(sa.client_email, undefined, sa.private_key, scopes);
  return { sheets: google.sheets({ version: 'v4', auth: jwt }), serviceAccountEmail: sa.client_email };
}

function cell(row, index) {
  return index < row.length ? String(row[index] ?? '').trim() : '';
}

export function normalizeMaterialWidth(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  const num = parseFloat(text.replace(/mm$/i, ''));
  return Number.isFinite(num) ? String(num) : text;
}

export function formatThicknessKey(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  if (/[A-Za-z]/.test(text)) return text.toUpperCase();
  const num = parseFloat(text);
  return Number.isFinite(num) ? num.toFixed(1) : text;
}

export function parseProductName(name) {
  const text = String(name ?? '').trim();
  const m = text.match(/^([\d.]+)\s*x\s*([\d.]+)\s*x\s*([\d.]+)\s+(.+?)\s+([\d.]+)\s*mm/i);
  if (!m) return null;
  return {
    thickness: formatThicknessKey(m[1]),
    width: normalizeMaterialWidth(m[2]),
    height: normalizeMaterialWidth(m[3]),
    type: m[4].trim(),
    length: normalizeMaterialWidth(m[5]),
    name: text,
  };
}

export function buildSpecKey({ type, thickness, width, height }) {
  return [
    String(type ?? '').trim(),
    formatThicknessKey(thickness),
    normalizeMaterialWidth(width),
    normalizeMaterialWidth(height),
  ].join('|');
}

async function readSheetValues(sheets, spreadsheetId, range) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  return res.data.values || [];
}

async function resolveSheetTitle(sheets, spreadsheetId, { gid, preferredTitle } = {}) {
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  const tabs = (res.data.sheets || []).map((s) => s.properties).filter(Boolean);
  let match = preferredTitle
    ? tabs.find((p) => p.title === preferredTitle)
    : null;
  if (!match && gid) match = tabs.find((p) => p.sheetId === gid);
  return { title: match?.title || null, tabs };
}

async function resolveEofficeSheetName(sheets, spreadsheetId) {
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties.title' });
  const titles = (res.data.sheets || []).map((s) => s.properties?.title).filter(Boolean);
  for (const candidate of EOFFICE_SHEET_CANDIDATES) {
    if (titles.includes(candidate)) return candidate;
  }
  const fuzzy = titles.find((t) => /e-office.*material.*width/i.test(t));
  return fuzzy || null;
}

function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

export async function loadPlistWithMaterialWidth(sheets) {
  const values = await readSheetValues(sheets, PQFORM_SPREADSHEET_ID, `${quoteSheetName(PLIST_SHEET_NAME)}!A:I`);
  const header = values[0] || [];
  const rows = values.slice(1).filter((row) => cell(row, 0));
  return { header, rows };
}

export async function loadEofficeMaterialWidth(sheets) {
  const sheetName = await resolveEofficeSheetName(sheets, PQFORM_SPREADSHEET_ID);
  if (!sheetName) return { sheetName: null, rows: [] };
  const values = await readSheetValues(
    sheets,
    PQFORM_SPREADSHEET_ID,
    `${quoteSheetName(sheetName)}!A:C`,
  );
  const rows = values.slice(1).filter((row) => cell(row, 0) || cell(row, 1));
  return { sheetName, rows };
}

export async function loadMonthlyDes(sheets) {
  const { title, tabs } = await resolveSheetTitle(sheets, MONTHLY_DES_SPREADSHEET_ID, {
    gid: MONTHLY_DES_GID,
    preferredTitle: MONTHLY_DES_SHEET_NAME,
  });
  if (!title) {
    return { sheetName: null, header: [], rows: [], tabs, error: 'Monthly Des. tab not found' };
  }
  const values = await readSheetValues(
    sheets,
    MONTHLY_DES_SPREADSHEET_ID,
    `${quoteSheetName(title)}!A:Z`,
  );
  return { sheetName: title, header: values[0] || [], rows: values.slice(1), tabs };
}

function indexByCode(rows, codeIndex, widthIndex, nameIndex = 1) {
  const map = new Map();
  for (const row of rows) {
    const code = cell(row, codeIndex);
    const width = normalizeMaterialWidth(cell(row, widthIndex));
    if (!code || !width) continue;
    map.set(code, { width, name: cell(row, nameIndex) });
  }
  return map;
}

function indexByName(rows, nameIndex, widthIndex) {
  const map = new Map();
  for (const row of rows) {
    const name = cell(row, nameIndex);
    const width = normalizeMaterialWidth(cell(row, widthIndex));
    if (!name || !width) continue;
    map.set(name.trim(), width);
    const parsed = parseProductName(name);
    if (parsed?.name) map.set(parsed.name, width);
  }
  return map;
}

export function normalizeMonthlyProductName(name) {
  return String(name ?? '')
    .replace(/\s*--\s*Sample.*$/i, '')
    .replace(/\s+Sample\s*$/i, '')
    .replace(/"/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function isMonthlyDesDataRow(row) {
  const t = cell(row, 0).replace(/\s/g, '');
  return /^\d+(\.\d+)?([A-Z])?$/i.test(t);
}

export function buildEnglishProductKeys(type, width, height) {
  const w = normalizeMaterialWidth(width);
  const h = normalizeMaterialWidth(height);
  const wNum = parseFloat(w);
  const keys = [];
  const add = (s) => {
    const k = normalizeMonthlyProductName(s);
    if (k) keys.push(k);
  };

  switch (String(type ?? '').trim()) {
    case '企筒':
      add(`${w}mm Stud (${h}mm H)`);
      if (Number.isFinite(wNum)) {
        add(`${wNum + 1}mm Stud (${h}mm H)`);
        add(`${wNum - 1}mm Stud (${h}mm H)`);
      }
      break;
    case '地槽':
      add(`${w}mm Runner (${h}mm H)`);
      if (Number.isFinite(wNum)) {
        add(`${wNum + 1}mm Runner (${h}mm H)`);
        add(`${wNum - 1}mm Runner (${h}mm H)`);
      }
      break;
    case '鐵角':
      add(`${w}mm L-Metal Angle`);
      add(`${h}mm L-Metal Angle`);
      if (Number.isFinite(wNum)) {
        add(`${wNum + 1}mm L-Metal Angle`);
        add(`${wNum - 1}mm L-Metal Angle`);
      }
      break;
    case 'W角':
      add('w-bar');
      add(`${w}mm w-bar`);
      break;
    case '闊槽':
    case 'C槽':
      add(`${w}mm Runner (${h}mm H)`);
      add('cw-255');
      break;
    default:
      break;
  }
  return [...new Set(keys)];
}

function parseMonthlyDesIndex(rows) {
  const byThicknessAndName = new Map();
  for (const row of rows) {
    if (!isMonthlyDesDataRow(row)) continue;
    const thickness = formatThicknessKey(cell(row, 0));
    const materialWidth = normalizeMaterialWidth(cell(row, 1));
    const productName = normalizeMonthlyProductName(cell(row, 3));
    if (!thickness || !materialWidth || !productName) continue;
    const key = `${thickness}|${productName}`;
    if (!byThicknessAndName.has(key)) byThicknessAndName.set(key, materialWidth);
  }
  return byThicknessAndName;
}

function lookupMonthlyWidth(monthlyIndex, thickness, englishKeys) {
  const formatted = formatThicknessKey(thickness);
  const lookup = thicknessForProductLookup(thickness);
  const thicknesses = [...new Set([formatted, lookup, formatted.replace(/\.0$/, '')].filter(Boolean))];
  for (const t of thicknesses) {
    for (const key of englishKeys) {
      const hit = monthlyIndex.get(`${t}|${key}`);
      if (hit) return hit;
    }
  }
  return '';
}

export function buildMaterialWidthRecommendations({ plistRows, eofficeRows, monthlyRows }) {
  const eofficeByCode = indexByCode(eofficeRows, 0, 2, 1);
  const eofficeByName = indexByName(eofficeRows, 1, 2);
  const monthlyIndex = parseMonthlyDesIndex(monthlyRows);

  const recommendations = [];
  const stats = {
    plistTotal: plistRows.length,
    alreadyFilled: 0,
    fromEoffice: 0,
    fromEofficeName: 0,
    fromMonthlyDes: 0,
    unresolved: 0,
    conflicts: 0,
  };

  for (let i = 0; i < plistRows.length; i += 1) {
    const row = plistRows[i];
    const code = cell(row, 0);
    const name = cell(row, 1);
    const type = cell(row, 3);
    const w = cell(row, 4);
    const h = cell(row, 5);
    const t = cell(row, 7);
    const existing = normalizeMaterialWidth(cell(row, 8));

    if (existing) {
      stats.alreadyFilled += 1;
      recommendations.push({
        rowIndex: i + 2,
        code,
        name,
        existing,
        recommended: existing,
        source: 'existing',
      });
      continue;
    }

    const candidates = [];
    const eoffice = eofficeByCode.get(code);
    if (eoffice) candidates.push({ width: eoffice.width, source: 'e-office' });

    const eofficeName = eofficeByName.get(name);
    if (eofficeName) candidates.push({ width: eofficeName, source: 'e-office-name' });

    const englishKeys = buildEnglishProductKeys(type, w, h);
    const monthlyWidth = lookupMonthlyWidth(monthlyIndex, t, englishKeys);
    if (monthlyWidth) candidates.push({ width: monthlyWidth, source: 'monthly-des' });

    const uniqueWidths = [...new Set(candidates.map((c) => c.width))];
    let recommended = '';
    let source = 'unresolved';

    if (uniqueWidths.length === 1) {
      recommended = uniqueWidths[0];
      source = candidates[0].source;
    } else if (uniqueWidths.length > 1) {
      stats.conflicts += 1;
      recommended = candidates[0].width;
      source = `${candidates[0].source}+conflict`;
    }

    if (recommended) {
      if (source.includes('e-office-name')) stats.fromEofficeName += 1;
      else if (source.includes('e-office')) stats.fromEoffice += 1;
      else if (source.includes('monthly-des')) stats.fromMonthlyDes += 1;
    } else {
      stats.unresolved += 1;
    }

    recommendations.push({
      rowIndex: i + 2,
      code,
      name,
      existing,
      recommended,
      source,
      candidates,
    });
  }

  return { recommendations, stats, eofficeByCode };
}

export async function analyzeMaterialWidthSheets() {
  const { sheets, serviceAccountEmail } = getMaterialWidthSheetsClient([
    'https://www.googleapis.com/auth/spreadsheets.readonly',
  ]);

  const access = {
    pqform: true,
    eoffice: false,
    monthlyDes: false,
    errors: [],
  };

  let plist = { header: [], rows: [] };
  let eoffice = { sheetName: null, rows: [] };
  let monthly = { sheetName: null, header: [], rows: [], tabs: [] };

  try {
    plist = await loadPlistWithMaterialWidth(sheets);
  } catch (e) {
    access.pqform = false;
    access.errors.push(`plist: ${e.message}`);
  }

  try {
    eoffice = await loadEofficeMaterialWidth(sheets);
    access.eoffice = !!eoffice.sheetName;
    if (!eoffice.sheetName) access.errors.push('e-office-material width tab not found');
  } catch (e) {
    access.errors.push(`e-office: ${e.message}`);
  }

  try {
    monthly = await loadMonthlyDes(sheets);
    access.monthlyDes = !!monthly.sheetName;
    if (monthly.error) access.errors.push(monthly.error);
  } catch (e) {
    access.errors.push(`monthly-des: ${e.message}`);
  }

  const { recommendations, stats, eofficeByCode } = buildMaterialWidthRecommendations({
    plistRows: plist.rows,
    eofficeRows: eoffice.rows,
    monthlyRows: monthly.rows,
  });

  const plistCodes = new Set(plist.rows.map((r) => cell(r, 0)).filter(Boolean));
  const eofficeOnly = [...eofficeByCode.keys()].filter((code) => !plistCodes.has(code));
  const plistMissingEoffice = [...plistCodes].filter((code) => !eofficeByCode.has(code));

  return {
    serviceAccountEmail,
    access,
    sheets: {
      plist: { name: PLIST_SHEET_NAME, header: plist.header, rowCount: plist.rows.length },
      eoffice: { name: eoffice.sheetName, rowCount: eoffice.rows.length },
      monthlyDes: {
        name: monthly.sheetName,
        header: monthly.header,
        rowCount: monthly.rows.length,
        tabs: (monthly.tabs || []).map((p) => ({ title: p.title, sheetId: p.sheetId })),
      },
    },
    stats,
    crossRef: {
      eofficeOnlyCount: eofficeOnly.length,
      eofficeOnlySample: eofficeOnly.slice(0, 20),
      plistMissingEofficeCount: plistMissingEoffice.length,
      plistMissingEofficeSample: plistMissingEoffice.slice(0, 20),
    },
    unresolvedSample: recommendations.filter((r) => !r.recommended && !r.existing).slice(0, 30),
    conflictSample: recommendations.filter((r) => r.source.includes('conflict')).slice(0, 20),
    recommendations,
  };
}

export async function syncMaterialWidthToPlist({ dryRun = true } = {}) {
  const { sheets } = getMaterialWidthSheetsClient([
    'https://www.googleapis.com/auth/spreadsheets',
  ]);

  const plist = await loadPlistWithMaterialWidth(sheets);
  const eoffice = await loadEofficeMaterialWidth(sheets);
  const monthly = await loadMonthlyDes(sheets);

  if (!monthly.sheetName) {
    throw new Error(
      `Monthly Des. sheet not accessible. Share ${MONTHLY_DES_SPREADSHEET_ID} with service account.`,
    );
  }

  const headerHasMaterialWidth = (plist.header[8] || '').includes('用料') || (plist.header[8] || '').toLowerCase().includes('material');
  const updates = [];
  const headerUpdates = [];

  if (!headerHasMaterialWidth) {
    headerUpdates.push({
      range: `${quoteSheetName(PLIST_SHEET_NAME)}!I1`,
      values: [['用料闊度']],
    });
  }

  const { recommendations, stats } = buildMaterialWidthRecommendations({
    plistRows: plist.rows,
    eofficeRows: eoffice.rows,
    monthlyRows: monthly.rows,
  });

  for (const rec of recommendations) {
    if (!rec.recommended || rec.existing) continue;
    updates.push({
      range: `${quoteSheetName(PLIST_SHEET_NAME)}!I${rec.rowIndex}`,
      values: [[rec.recommended]],
    });
  }

  if (dryRun) {
    return { dryRun: true, stats, updateCount: updates.length, headerUpdates, sampleUpdates: updates.slice(0, 20) };
  }

  const data = [...headerUpdates, ...updates].map((u) => ({ range: u.range, values: u.values }));
  if (data.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: PQFORM_SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data },
    });
  }

  return { dryRun: false, stats, updateCount: updates.length, headerUpdated: headerUpdates.length > 0 };
}
