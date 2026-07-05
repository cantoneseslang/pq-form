import {
  PQFORM_SPREADSHEET_ID,
  PLIST_SHEET_NAME,
  getMaterialWidthSheetsClient,
} from './materialWidthSheets.js';
import { analyzeRawMaterialInventory } from './rawMaterialInventorySheets.js';

export const MATERIAL_STOCK_OUTPUT_SPREADSHEET_ID = process.env.MATERIAL_STOCK_OUTPUT_SHEET_ID
  || process.env.PQFORM_SHEET_ID
  || PQFORM_SPREADSHEET_ID;
export const MATERIAL_STOCK_SHEET_NAME = process.env.MATERIAL_STOCK_SHEET_NAME || 'material-stock';
export const MATERIAL_STOCK_SUMMARY_SHEET_NAME = process.env.MATERIAL_STOCK_SUMMARY_SHEET_NAME || 'material-stock-summary';

const PROTECTED_SHEET_NAMES = new Set([
  PLIST_SHEET_NAME,
  'PQ-Form-plist',
  'pq-form',
  process.env.PQFORM_SHEET_NAME,
  process.env.PQFORM_PLIST_SHEET_NAME,
].filter(Boolean));

export function assertSafeMaterialStockSheetName(sheetName) {
  const title = String(sheetName ?? '').trim();
  if (!title) throw new Error('Material stock sheet name is empty');
  if (PROTECTED_SHEET_NAMES.has(title)) {
    throw new Error(`Refusing to write material stock to protected sheet "${title}"`);
  }
  if (/plist/i.test(title)) {
    throw new Error(`Refusing to write material stock to plist-like sheet "${title}"`);
  }
  return title;
}

const LOT_HEADERS = [
  'tabTitle',
  'thicknessKey',
  'receiptDate',
  'lotNo',
  'materialWidthMm',
  'weight',
  'inboundRolls',
  'outboundRolls',
  'remainingRolls',
  'inboundKg',
  'outboundKg',
  'availableKg',
  'unitPricePerKg',
  'totalPrice',
  'densityGcm3',
];

const SUMMARY_HEADERS = [
  'thicknessKey',
  'tabTitle',
  'materialWidthMm',
  'totalKg',
  'totalRolls',
  'lotCount',
];

function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

/** 15-12-2019 / 28-9-2017 / 28/9/2017 → 2019/12/15 */
export function formatReceiptDateForSheet(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';

  const dmyDash = text.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (dmyDash) {
    const month = String(dmyDash[2]).padStart(2, '0');
    const day = String(dmyDash[1]).padStart(2, '0');
    return `${dmyDash[3]}/${month}/${day}`;
  }

  const dmySlash = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmySlash) {
    const month = String(dmySlash[2]).padStart(2, '0');
    const day = String(dmySlash[1]).padStart(2, '0');
    return `${dmySlash[3]}/${month}/${day}`;
  }

  const ymd = text.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (ymd) {
    const month = String(ymd[2]).padStart(2, '0');
    const day = String(ymd[3]).padStart(2, '0');
    return `${ymd[1]}/${month}/${day}`;
  }

  return text;
}

function receiptDateSortKey(value) {
  const formatted = formatReceiptDateForSheet(value);
  const ymd = formatted.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
  if (ymd) {
    return new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3])).getTime();
  }
  return 0;
}

function lotToRow(lot) {
  return [
    lot.tabTitle,
    lot.thicknessKey,
    formatReceiptDateForSheet(lot.receiptDate),
    lot.lotNo,
    lot.materialWidth,
    lot.weight ?? '',
    lot.inboundRolls ?? '',
    lot.outboundRolls ?? '',
    lot.rolls ?? '',
    lot.totalKg ?? '',
    lot.outboundKg ?? '',
    lot.availableKg ?? '',
    lot.unitPricePerKg ?? '',
    lot.totalPrice ?? '',
    lot.densityGcm3 ?? '',
  ];
}

function summaryToRow(summary) {
  return [
    summary.thicknessKey,
    summary.tabTitle,
    summary.materialWidth,
    summary.totalKg,
    summary.totalRolls,
    summary.lotCount,
  ];
}

export function buildMaterialStockLotsSheetValues({ lots, stats, fetchedAt, sourceSpreadsheetId }) {
  const sortedLots = [...lots].sort((a, b) => {
    const tab = String(a.tabTitle).localeCompare(String(b.tabTitle));
    if (tab) return tab;
    const mw = String(a.materialWidth).localeCompare(String(b.materialWidth), undefined, { numeric: true });
    if (mw) return mw;
    return receiptDateSortKey(a.receiptDate) - receiptDateSortKey(b.receiptDate);
  });

  return [
    ['pq-form material-stock (lots)'],
    ['最終更新', fetchedAt],
    ['来源シート', sourceSpreadsheetId],
    ['ロット数', stats.totalLots, '可用kg合計', stats.totalKg],
    [],
    LOT_HEADERS,
    ...sortedLots.map(lotToRow),
  ];
}

export function buildMaterialStockSummarySheetValues({ summaries, stats, fetchedAt, sourceSpreadsheetId }) {
  const sortedSummaries = [...summaries].sort((a, b) => {
    const tab = String(a.tabTitle).localeCompare(String(b.tabTitle));
    if (tab) return tab;
    return String(a.materialWidth).localeCompare(String(b.materialWidth), undefined, { numeric: true });
  });

  return [
    ['pq-form material-stock (summary)'],
    ['最終更新', fetchedAt],
    ['来源シート', sourceSpreadsheetId],
    ['幅×厚組合', stats.uniqueThicknessWidth, '可用kg合計', stats.totalKg],
    [],
    SUMMARY_HEADERS,
    ...sortedSummaries.map(summaryToRow),
  ];
}

async function resolveSheetTab(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  const tabs = (res.data.sheets || []).map((s) => s.properties).filter(Boolean);
  const match = tabs.find((p) => p.title === sheetName);
  return { title: match?.title || null, sheetId: match?.sheetId ?? null, tabs };
}

async function ensureSheetTab(sheets, spreadsheetId, sheetName) {
  const existing = await resolveSheetTab(sheets, spreadsheetId, sheetName);
  if (existing.title) return { ...existing, created: false };

  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        addSheet: {
          properties: { title: sheetName },
        },
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

async function writeSheetValues(sheets, spreadsheetId, sheetTitle, values, { valueInputOption = 'RAW' } = {}) {
  const safeTitle = assertSafeMaterialStockSheetName(sheetTitle);
  const quoted = quoteSheetName(safeTitle);
  await sheets.spreadsheets.values.clear({
    spreadsheetId,
    range: `${quoted}!A:ZZ`,
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${quoted}!A1`,
    valueInputOption,
    requestBody: { values },
  });
}

function sheetUrl(spreadsheetId, sheetId) {
  return sheetId
    ? `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetId}`
    : `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
}

export async function syncMaterialStockLotsToSheet({ dryRun = false } = {}) {
  const report = await analyzeRawMaterialInventory();
  const lots = report.tabReports.flatMap((t) => t.lots);
  const fetchedAt = new Date().toISOString();
  const lotsValues = buildMaterialStockLotsSheetValues({
    lots,
    stats: report.stats,
    fetchedAt,
    sourceSpreadsheetId: report.spreadsheetId,
  });
  const summaryValues = buildMaterialStockSummarySheetValues({
    summaries: report.summaries,
    stats: report.stats,
    fetchedAt,
    sourceSpreadsheetId: report.spreadsheetId,
  });

  if (dryRun) {
    return {
      dryRun: true,
      fetchedAt,
      lotCount: lots.length,
      summaryCount: report.summaries.length,
      lotsRowCount: lotsValues.length,
      summaryRowCount: summaryValues.length,
      stats: report.stats,
      lotsPreview: lotsValues.slice(0, 10),
      summaryPreview: summaryValues.slice(0, 10),
    };
  }

  const { sheets, serviceAccountEmail } = getMaterialWidthSheetsClient([
    'https://www.googleapis.com/auth/spreadsheets',
  ]);
  const spreadsheetId = MATERIAL_STOCK_OUTPUT_SPREADSHEET_ID;

  const lotsTab = await ensureSheetTab(sheets, spreadsheetId, assertSafeMaterialStockSheetName(MATERIAL_STOCK_SHEET_NAME));
  const summaryTab = await ensureSheetTab(
    sheets,
    spreadsheetId,
    assertSafeMaterialStockSheetName(MATERIAL_STOCK_SUMMARY_SHEET_NAME),
  );

  await writeSheetValues(sheets, spreadsheetId, lotsTab.title, lotsValues);
  await writeSheetValues(sheets, spreadsheetId, summaryTab.title, summaryValues);

  return {
    dryRun: false,
    fetchedAt,
    spreadsheetId,
    serviceAccountEmail,
    lotCount: lots.length,
    summaryCount: report.summaries.length,
    lotsRowCount: lotsValues.length,
    summaryRowCount: summaryValues.length,
    stats: report.stats,
    lotsSheet: {
      title: lotsTab.title,
      sheetId: lotsTab.sheetId,
      sheetUrl: sheetUrl(spreadsheetId, lotsTab.sheetId),
      created: lotsTab.created,
    },
    summarySheet: {
      title: summaryTab.title,
      sheetId: summaryTab.sheetId,
      sheetUrl: sheetUrl(spreadsheetId, summaryTab.sheetId),
      created: summaryTab.created,
    },
  };
}
