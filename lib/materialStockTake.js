import {
  MATERIAL_STOCK_OUTPUT_SPREADSHEET_ID,
  MATERIAL_STOCK_SHEET_NAME,
  formatReceiptDateForSheet,
} from './materialStockSync.js';
import { getMaterialWidthSheetsClient } from './materialWidthSheets.js';
import { getCachedMaterialStock } from './materialStockCache.js';

const LOT_FIELD_INDEX = {
  tabTitle: 0,
  thicknessKey: 1,
  receiptDate: 2,
  lotNo: 3,
  materialWidth: 4,
  weight: 5,
  remainingRolls: 8,
  availableKg: 11,
};

function parseSheetNumber(value) {
  const n = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : 0;
}

function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatNumber(value) {
  if (value === null || value === undefined || value === '') return '';
  const n = Number(value);
  if (!Number.isFinite(n)) return String(value);
  return n.toLocaleString('en-US', { maximumFractionDigits: 4 });
}

function formatInteger(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.round(n);
}

function parseWidth(value) {
  const n = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : 0;
}

function receiptDateSortKey(value) {
  const formatted = formatReceiptDateForSheet(value);
  const ymd = formatted.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
  if (ymd) {
    return new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3])).getTime();
  }
  return 0;
}

function hongKongTodayStartMs(now = new Date()) {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Hong_Kong',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(now);
  const pick = (type) => parts.find((p) => p.type === type)?.value ?? '';
  return new Date(
    Number(pick('year')),
    Number(pick('month')) - 1,
    Number(pick('day')),
  ).getTime();
}

export function receiptAgeDays(receiptDate, now = new Date()) {
  const receiptMs = receiptDateSortKey(receiptDate);
  if (!receiptMs) return null;
  const diffMs = hongKongTodayStartMs(now) - receiptMs;
  return Math.max(0, Math.floor(diffMs / 86400000));
}

function receiptAgeClass(receiptDate) {
  const days = receiptAgeDays(receiptDate);
  if (days === null) return 'mst-age';
  if (days <= 365) return 'mst-age mst-age--fresh';
  if (days <= 730) return 'mst-age mst-age--medium';
  return 'mst-age mst-age--old';
}

export function aggregateAgeKgStats(lots) {
  const stats = { within365Kg: 0, within730Kg: 0, over730Kg: 0 };

  for (const lot of lots) {
    const kg = parseSheetNumber(lot.availableKg);
    if (kg <= 0) continue;
    const days = receiptAgeDays(lot.receiptDate);
    if (days === null) continue;
    if (days <= 365) stats.within365Kg += kg;
    else if (days <= 730) stats.within730Kg += kg;
    else stats.over730Kg += kg;
  }

  return {
    within365Kg: Math.round(stats.within365Kg),
    within730Kg: Math.round(stats.within730Kg),
    over730Kg: Math.round(stats.over730Kg),
  };
}

function lotRowKey(lot) {
  return [lot.tabTitle, lot.materialWidth, lot.lotNo, lot.receiptDate].join('|');
}

function rowToLot(row) {
  if (!row?.length) return null;
  const tabTitle = String(row[LOT_FIELD_INDEX.tabTitle] ?? '').trim();
  const materialWidth = String(row[LOT_FIELD_INDEX.materialWidth] ?? '').trim();
  const lotNo = String(row[LOT_FIELD_INDEX.lotNo] ?? '').trim();
  if (!tabTitle || !materialWidth || !lotNo) return null;

  return {
    tabTitle,
    thicknessKey: row[LOT_FIELD_INDEX.thicknessKey] ?? '',
    receiptDate: row[LOT_FIELD_INDEX.receiptDate] ?? '',
    lotNo,
    materialWidth,
    weight: row[LOT_FIELD_INDEX.weight] ?? '',
    rolls: row[LOT_FIELD_INDEX.remainingRolls] ?? '',
    availableKg: row[LOT_FIELD_INDEX.availableKg] ?? '',
  };
}

function sumLotAvailableKg(lots) {
  return lots.reduce((sum, lot) => sum + parseSheetNumber(lot.availableKg), 0);
}

function parseSheetPayload(rows) {
  const fetchedAt = rows?.[1]?.[1] || '';
  const sourceSpreadsheetId = rows?.[2]?.[1] || '';
  const totalLots = parseSheetNumber(rows?.[3]?.[1]);
  const headerRow = rows?.[5] || [];
  const dataRows = rows.slice(6);
  const lots = dataRows.map(rowToLot).filter(Boolean);
  const totalKgFromMeta = parseSheetNumber(rows?.[3]?.[3]);
  const totalKgFromLots = sumLotAvailableKg(lots);

  return {
    fetchedAt,
    sourceSpreadsheetId,
    stats: {
      totalLots: totalLots || lots.length,
      totalKg: totalKgFromMeta || totalKgFromLots,
    },
    lots,
    headerRow,
  };
}

export async function readMaterialStockLotsFromSheet() {
  const { sheets } = getMaterialWidthSheetsClient([
    'https://www.googleapis.com/auth/spreadsheets.readonly',
  ]);
  const spreadsheetId = MATERIAL_STOCK_OUTPUT_SPREADSHEET_ID;
  const quoted = quoteSheetName(MATERIAL_STOCK_SHEET_NAME);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${quoted}!A1:O2000`,
  });
  const rows = res.data.values || [];
  const parsed = parseSheetPayload(rows);
  if (!parsed.lots.length) {
    throw new Error('material-stock sheet has no lot rows');
  }
  return parsed;
}

function flattenLotsFromStock(stock) {
  const byTab = stock?.byTab || {};
  return Object.values(byTab).flatMap((bucket) => bucket.lots || []);
}

export function sortLots(items) {
  return [...items].sort((a, b) => {
    const tabCmp = String(a.tabTitle).localeCompare(String(b.tabTitle));
    if (tabCmp !== 0) return tabCmp;
    const widthCmp = parseWidth(a.materialWidth) - parseWidth(b.materialWidth);
    if (widthCmp !== 0) return widthCmp;
    const dateCmp = receiptDateSortKey(a.receiptDate) - receiptDateSortKey(b.receiptDate);
    if (dateCmp !== 0) return dateCmp;
    return String(a.lotNo).localeCompare(String(b.lotNo));
  });
}

export function buildTakeTableHtml(lots) {
  const parts = [];
  let currentTab = null;
  let currentWidth = null;

  for (const lot of lots) {
    const tab = String(lot.tabTitle ?? '');
    const width = String(lot.materialWidth ?? '');

    if (tab !== currentTab) {
      parts.push(`<tr class="mst-group mst-group--tab"><td colspan="7">${escapeHtml(tab)}</td></tr>`);
      currentTab = tab;
      currentWidth = null;
    }
    if (width !== currentWidth) {
      parts.push(`<tr class="mst-group mst-group--width"><td colspan="7">${escapeHtml(width)}</td></tr>`);
      currentWidth = width;
    }

    const key = escapeHtml(lotRowKey(lot));
    const receiptDate = escapeHtml(formatReceiptDateForSheet(lot.receiptDate));
    const ageClass = receiptAgeClass(lot.receiptDate);
    parts.push(
      `<tr data-row-key="${key}">`
      + `<td>${escapeHtml(tab)}</td>`
      + `<td>${escapeHtml(width)}</td>`
      + `<td class="${ageClass}">${receiptDate}</td>`
      + `<td>${escapeHtml(lot.lotNo)}</td>`
      + `<td class="num">${escapeHtml(formatNumber(lot.weight))}</td>`
      + `<td class="num">${escapeHtml(formatNumber(lot.rolls))}</td>`
      + `<td><input type="text" class="mst-adjust-input" inputmode="decimal" autocomplete="off" data-row-key="${key}" value=""></td>`
      + '</tr>',
    );
  }

  return parts.join('');
}

function buildTakePayloadBase({ source, fetchedAt, sourceSpreadsheetId, stats, lots }) {
  const sortedLots = sortLots(lots);
  return {
    source,
    fetchedAt,
    sourceSpreadsheetId,
    stats,
    lots: sortedLots,
    ageStats: aggregateAgeKgStats(sortedLots),
    rowsHtml: buildTakeTableHtml(sortedLots),
  };
}

export async function loadMaterialStockTakePayload() {
  try {
    const sheetData = await readMaterialStockLotsFromSheet();
    return buildTakePayloadBase({
      source: 'sheet',
      fetchedAt: sheetData.fetchedAt,
      sourceSpreadsheetId: sheetData.sourceSpreadsheetId,
      stats: sheetData.stats,
      lots: sheetData.lots,
    });
  } catch {
    const stock = await getCachedMaterialStock();
    return buildTakePayloadBase({
      source: 'live',
      fetchedAt: stock.fetchedAt,
      sourceSpreadsheetId: stock.sourceSpreadsheetId,
      stats: stock.stats,
      lots: flattenLotsFromStock(stock),
    });
  }
}
