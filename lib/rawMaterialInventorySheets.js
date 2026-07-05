import { getMaterialWidthSheetsClient, normalizeMaterialWidth } from './materialWidthSheets.js';
import { densityForMaterial } from './materialYield.js';

export const RAW_MATERIAL_SPREADSHEET_ID = process.env.RAW_MATERIAL_INVENTORY_SHEET_ID
  || process.env.MATERIAL_WIDTH_DES_SHEET_ID
  || '1R-xjzmki0pzMlXJzhzVee_y6XWqFbrUSlSL04osBhbc';

export const EXCLUDED_TAB_PATTERNS = [
  /^monthly des\./i,
  /^kpi$/i,
  /^purchase$/i,
  /^update bal\./i,
  /^calculation of/i,
  /^at yiu lee/i,
  /^mat'?ls used/i,
  /^re-slit/i,
  /^bal\./i,
  /^feet per/i,
  /^e-man/i,
  /^tk\./i,
];

export const COL = {
  receiptDate: 0,   // A 入庫日
  lotNo: 1,         // B ロット
  thickness: 2,     // C 厚度
  materialWidth: 3, // D 材料闊度
  weight: 4,        // E 単重 kg
  rolls: 5,         // F 卷數（入庫）
  totalKg: 6,       // G 總數 kg（入庫）
  totalPrice: 7,    // H 總價
};

/** AP/AQ = 出庫卷數・出庫kg。在庫 = F−AP、G−AQ */
export const OUTBOUND_COL = {
  rolls: 41, // AP
  kg: 42,    // AQ
};

export const INVENTORY_LAST_COL = 'AQ';

const SUMMARY_MARKERS = /^(合計|总计|總計|total|subtotal|balance|結存|结存)/i;

function cell(row, index) {
  return index < row.length ? String(row[index] ?? '').trim() : '';
}

function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function parseNumber(value) {
  const text = String(value ?? '').replace(/,/g, '').trim();
  if (!text) return null;
  const num = parseFloat(text);
  return Number.isFinite(num) ? num : null;
}

export function isThicknessInventoryTab(title) {
  const t = String(title ?? '').trim();
  if (!t || EXCLUDED_TAB_PATTERNS.some((re) => re.test(t))) return false;
  return /^\d/.test(t) && /mm/i.test(t);
}

export function parseThicknessKeyFromTab(tabName) {
  const t = String(tabName ?? '').trim();
  const dashA = t.match(/^([\d.]+)-A\s*mm$/i);
  if (dashA) {
    const num = parseFloat(dashA[1]);
    return Number.isFinite(num) ? `${num}A`.replace(/\.0A$/, 'A') : '0.8A';
  }

  const astm = t.match(/^([\d.]+)\s*mm\s*ASTM\s*\(([^)]+)\)/i);
  if (astm) {
    const num = parseFloat(astm[1]);
    const grade = astm[2].trim().toUpperCase();
    return Number.isFinite(num) ? `${num}|ASTM_${grade}` : t;
  }

  const paren = t.match(/^([\d.]+)\s*mm\s*\(([^)]+)\)/i);
  if (paren) {
    const num = parseFloat(paren[1]);
    const suffix = paren[2].trim();
    if (/^B$/i.test(suffix)) return Number.isFinite(num) ? `${num}B` : '0.4B';
    if (/^W$/i.test(suffix)) return Number.isFinite(num) ? `${num}W` : '0.4W';
    if (/aluminium/i.test(suffix)) return Number.isFinite(num) ? `${num}AL` : '0.4AL';
    if (/Z-120/i.test(suffix)) return Number.isFinite(num) ? `${num}C` : '0.8C';
    if (/Z-275|G90/i.test(suffix)) return Number.isFinite(num) ? `${num}|Z275_G90` : t;
    return Number.isFinite(num) ? `${num}|${suffix.toUpperCase()}` : t;
  }

  const plain = t.match(/^([\d.]+)\s*mm$/i);
  if (plain) {
    const num = parseFloat(plain[1]);
    return Number.isFinite(num) ? (Number.isInteger(num) ? String(num) : num.toFixed(1).replace(/\.0$/, '')) : t;
  }

  return t;
}

/** Map pq-form thickness select value → inventory tab title candidates */
export function thicknessToTabCandidates(thickness) {
  const t = String(thickness ?? '').trim().toUpperCase();
  const map = {
    '0.8A': ['0.8-A mm'],
    '0.4D': ['0.4mm (B)'],
    '0.4B': ['0.4mm (B)'],
    '0.4W': ['0.4mm (W)'],
    '0.4AL': ['0.4mm (Aluminium)'],
    '0.8C': ['0.8mm C (Z-120)'],
    '0.3': ['0.3mm'],
    '0.4': ['0.4mm'],
    '0.45': ['0.45mm'],
    '0.5': ['0.5mm'],
    '0.6': ['0.6mm'],
    '0.8': ['0.8mm'],
    '1.0': ['1.0mm'],
    '1.2': ['1.2mm'],
    '1.5': ['1.5mm', '1.5mm ASTM (G90)'],
    '3.0': ['3.0mm'],
  };
  if (map[t]) return map[t];
  const num = parseFloat(t);
  if (Number.isFinite(num)) return [`${num}mm`];
  return [];
}

function inventoryLookupKey(thicknessKey, materialWidth) {
  return `${thicknessKey}|${normalizeMaterialWidth(materialWidth)}`;
}

function tabLookupKey(tabTitle, materialWidth) {
  return `${tabTitle}|${normalizeMaterialWidth(materialWidth)}`;
}

function outboundQtyFromRow(row, colIndex) {
  const n = parseNumber(cell(row, colIndex));
  return n !== null && n > 0 ? n : 0;
}

function remainingFromRow(row) {
  const inboundKg = parseNumber(cell(row, COL.totalKg));
  if (inboundKg === null) return null;

  const inboundRolls = parseNumber(cell(row, COL.rolls));
  const outboundKg = outboundQtyFromRow(row, OUTBOUND_COL.kg);
  const outboundRolls = outboundQtyFromRow(row, OUTBOUND_COL.rolls);

  const availableKg = Math.max(0, inboundKg - outboundKg);
  const availableRolls = inboundRolls !== null
    ? Math.max(0, inboundRolls - outboundRolls)
    : null;

  return {
    inboundKg,
    inboundRolls,
    outboundKg,
    outboundRolls,
    availableKg,
    availableRolls,
  };
}

function detectHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 20); i += 1) {
    const row = rows[i] || [];
    const joined = row.join(' ').toLowerCase();
    if (/raw\s*materials\s*consumed/i.test(joined)) continue;
    if (/^june,|^july,|^august,/i.test(cell(row, 4))) continue;
    if (
      joined.includes('入庫') || joined.includes('ロット') || joined.includes('lot')
      || joined.includes('厚度') || (joined.includes('材料') && joined.includes('闊'))
      || (joined.includes('卷') && joined.includes('數'))
      || (joined.includes('總') && joined.includes('價'))
    ) {
      return i;
    }
  }
  return 0;
}

function isDataRow(row, headerRowIndex, rowIndex) {
  if (rowIndex <= headerRowIndex) return false;
  const lot = cell(row, COL.lotNo);
  const width = cell(row, COL.materialWidth);
  const totalKg = cell(row, COL.totalKg);
  const joined = row.slice(0, 8).join(' ');
  if (!lot && !width && !totalKg) return false;
  if (SUMMARY_MARKERS.test(lot) || SUMMARY_MARKERS.test(joined)) return false;
  if (!width && !totalKg) return false;
  return true;
}

export function parseInventoryRow({
  row,
  rowIndex,
  tabTitle,
  thicknessKey,
  headerRowIndex,
}) {
  if (!isDataRow(row, headerRowIndex, rowIndex)) return null;

  const materialWidth = normalizeMaterialWidth(cell(row, COL.materialWidth));
  const remaining = remainingFromRow(row);
  if (!materialWidth || !remaining || remaining.availableKg <= 0) return null;

  return {
    tabTitle,
    thicknessKey,
    receiptDate: cell(row, COL.receiptDate),
    lotNo: cell(row, COL.lotNo),
    thickness: cell(row, COL.thickness),
    materialWidth,
    weight: parseNumber(cell(row, COL.weight)),
    rolls: remaining.availableRolls,
    inboundRolls: remaining.inboundRolls,
    outboundRolls: remaining.outboundRolls,
    totalKg: remaining.inboundKg,
    outboundKg: remaining.outboundKg,
    availableKg: remaining.availableKg,
    totalPrice: parseNumber(cell(row, COL.totalPrice)),
    densityGcm3: densityForMaterial({ tabName: tabTitle, thicknessKey }),
  };
}

export function aggregateInventoryLots(lots) {
  const byKey = new Map();

  for (const lot of lots) {
    const key = inventoryLookupKey(lot.thicknessKey, lot.materialWidth);
    const tabKey = tabLookupKey(lot.tabTitle, lot.materialWidth);

    if (!byKey.has(key)) {
      byKey.set(key, {
        thicknessKey: lot.thicknessKey,
        tabTitle: lot.tabTitle,
        materialWidth: lot.materialWidth,
        totalKg: 0,
        totalRolls: 0,
        lotCount: 0,
        lots: [],
        tabKeys: new Set(),
      });
    }

    const bucket = byKey.get(key);
    bucket.totalKg += lot.availableKg || 0;
    bucket.totalRolls += lot.rolls || 0;
    bucket.lotCount += 1;
    bucket.lots.push(lot);
    bucket.tabKeys.add(tabKey);
  }

  const summaries = [...byKey.values()].map((b) => ({
    thicknessKey: b.thicknessKey,
    tabTitle: b.tabTitle,
    materialWidth: b.materialWidth,
    totalKg: Math.round(b.totalKg * 1000) / 1000,
    totalRolls: b.totalRolls,
    lotCount: b.lotCount,
    lots: b.lots,
    tabKeys: [...b.tabKeys],
  }));

  const byTab = new Map();
  for (const lot of lots) {
    const tk = tabLookupKey(lot.tabTitle, lot.materialWidth);
    if (!byTab.has(tk)) {
      byTab.set(tk, {
        tabTitle: lot.tabTitle,
        thicknessKey: lot.thicknessKey,
        materialWidth: lot.materialWidth,
        totalKg: 0,
        totalRolls: 0,
        lotCount: 0,
        lots: [],
      });
    }
    const bucket = byTab.get(tk);
    bucket.totalKg += lot.availableKg || 0;
    bucket.totalRolls += lot.rolls || 0;
    bucket.lotCount += 1;
    bucket.lots.push(lot);
  }

  return {
    summaries,
    byTab: Object.fromEntries(
      [...byTab.entries()].map(([k, v]) => [k, { ...v, totalKg: Math.round(v.totalKg * 1000) / 1000 }]),
    ),
  };
}

async function readSheetValues(sheets, spreadsheetId, range) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  return res.data.values || [];
}

export async function loadRawMaterialTab(sheets, spreadsheetId, tabTitle) {
  const range = `${quoteSheetName(tabTitle)}!A:${INVENTORY_LAST_COL}`;
  const rows = await readSheetValues(sheets, spreadsheetId, range);
  const headerRowIndex = detectHeaderRow(rows);
  const headerRows = rows.slice(0, headerRowIndex + 1);
  const thicknessKey = parseThicknessKeyFromTab(tabTitle);

  const lots = [];
  const skipped = { empty: 0, summary: 0, invalid: 0 };

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    if (i <= headerRowIndex) continue;
    if (!row || row.every((c) => !String(c ?? '').trim())) {
      skipped.empty += 1;
      continue;
    }
    const parsed = parseInventoryRow({
      row,
      rowIndex: i,
      tabTitle,
      thicknessKey,
      headerRowIndex,
    });
    if (!parsed) {
      if (SUMMARY_MARKERS.test(cell(row, COL.lotNo))) skipped.summary += 1;
      else skipped.invalid += 1;
      continue;
    }
    lots.push(parsed);
  }

  return {
    tabTitle,
    sheetId: null,
    thicknessKey,
    headerRowIndex,
    headerSample: headerRows,
    inventorySource: 'remaining: F−AP rolls, G−AQ kg',
    rowCount: rows.length,
    lotCount: lots.length,
    skipped,
    lots,
  };
}

export async function listInventoryTabs(sheets, spreadsheetId = RAW_MATERIAL_SPREADSHEET_ID) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: 'properties.title,sheets.properties' });
  const tabs = (meta.data.sheets || []).map((s) => s.properties).filter(Boolean);
  const inventoryTabs = tabs.filter((p) => isThicknessInventoryTab(p.title));
  return {
    spreadsheetTitle: meta.data.properties?.title || '',
    allTabs: tabs.map((p) => ({ title: p.title, sheetId: p.sheetId })),
    inventoryTabs: inventoryTabs.map((p) => ({
      title: p.title,
      sheetId: p.sheetId,
      thicknessKey: parseThicknessKeyFromTab(p.title),
    })),
  };
}

export async function analyzeRawMaterialInventory({
  spreadsheetId = RAW_MATERIAL_SPREADSHEET_ID,
  tabFilter = null,
} = {}) {
  const { sheets, serviceAccountEmail } = getMaterialWidthSheetsClient([
    'https://www.googleapis.com/auth/spreadsheets.readonly',
  ]);

  const { spreadsheetTitle, allTabs, inventoryTabs } = await listInventoryTabs(sheets, spreadsheetId);
  const targetTabs = tabFilter
    ? inventoryTabs.filter((t) => tabFilter.includes(t.title))
    : inventoryTabs;

  const tabReports = [];
  const allLots = [];
  const errors = [];

  for (const tab of targetTabs) {
    try {
      const report = await loadRawMaterialTab(sheets, spreadsheetId, tab.title);
      report.sheetId = tab.sheetId;
      tabReports.push(report);
      allLots.push(...report.lots);
    } catch (e) {
      errors.push({ tab: tab.title, error: e.message });
    }
  }

  const { summaries, byTab } = aggregateInventoryLots(allLots);
  const anomalies = allLots.filter((l) => !l.materialWidth || !l.availableKg);

  return {
    serviceAccountEmail,
    spreadsheetId,
    spreadsheetTitle,
    access: { ok: true, tabCount: targetTabs.length, errors },
    inventoryTabs,
    allTabs,
    tabReports,
    summaries,
    byTab,
    stats: {
      totalLots: allLots.length,
      totalKg: Math.round(summaries.reduce((s, x) => s + x.totalKg, 0) * 1000) / 1000,
      uniqueThicknessWidth: summaries.length,
      anomalyCount: anomalies.length,
    },
    anomalies: anomalies.slice(0, 30),
  };
}

export async function fetchMaterialStockMap({
  spreadsheetId = RAW_MATERIAL_SPREADSHEET_ID,
} = {}) {
  const report = await analyzeRawMaterialInventory({ spreadsheetId });
  return {
    fetchedAt: new Date().toISOString(),
    byTab: report.byTab,
    summaries: report.summaries,
    stats: report.stats,
  };
}

export function lookupMaterialStock({
  byTab,
  thickness,
  materialWidth,
  tabCandidates = null,
}) {
  const mw = normalizeMaterialWidth(materialWidth);
  if (!mw) return null;

  const candidates = tabCandidates || thicknessToTabCandidates(thickness);
  let totalKg = 0;
  let totalRolls = 0;
  const lots = [];
  const matchedTabs = [];

  for (const tabTitle of candidates) {
    const key = tabLookupKey(tabTitle, mw);
    const hit = byTab[key];
    if (!hit) continue;
    matchedTabs.push(tabTitle);
    totalKg += hit.totalKg || 0;
    totalRolls += hit.totalRolls || 0;
    lots.push(...(hit.lots || []));
  }

  if (!matchedTabs.length) return null;

  return {
    thickness: String(thickness ?? '').trim(),
    materialWidth: mw,
    tabTitles: matchedTabs,
    totalKg: Math.round(totalKg * 1000) / 1000,
    totalRolls,
    lotCount: lots.length,
    lots,
  };
}

export function buildSummaryCsvRows(summaries) {
  const lines = ['thicknessKey,tabTitle,materialWidthMm,totalKg,totalRolls,lotCount'];
  for (const s of summaries) {
    lines.push([
      s.thicknessKey,
      `"${String(s.tabTitle).replace(/"/g, '""')}"`,
      s.materialWidth,
      s.totalKg,
      s.totalRolls,
      s.lotCount,
    ].join(','));
  }
  return lines;
}

export function buildLotsCsvRows(lots) {
  const lines = ['tabTitle,thicknessKey,receiptDate,lotNo,materialWidthMm,weight,inboundRolls,outboundRolls,remainingRolls,inboundKg,outboundKg,availableKg,totalPrice'];
  for (const l of lots) {
    lines.push([
      `"${String(l.tabTitle).replace(/"/g, '""')}"`,
      l.thicknessKey,
      `"${String(l.receiptDate).replace(/"/g, '""')}"`,
      `"${String(l.lotNo).replace(/"/g, '""')}"`,
      l.materialWidth,
      l.weight ?? '',
      l.inboundRolls ?? '',
      l.outboundRolls ?? '',
      l.rolls ?? '',
      l.totalKg ?? '',
      l.outboundKg ?? '',
      l.availableKg ?? '',
      l.totalPrice ?? '',
    ].join(','));
  }
  return lines;
}
