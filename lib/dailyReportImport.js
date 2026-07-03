import { readRange } from './dailyReportSheets.js';
import {
  TYPE_ENGLISH_NAMES,
  resolveMoldingMachineBlock,
  resolveAutoMachineBlock,
  isDailyReportRowEmpty,
  MOLDING_MACHINE_NAMES,
} from './dailyReport.js';
import { searchPlist } from './plist.js';
import { formatRecordDateDisplay } from './productionRecords.js';

const ENGLISH_TO_TYPE = Object.fromEntries(
  Object.entries(TYPE_ENGLISH_NAMES).map(([key, en]) => [en.toLowerCase(), key]),
);

function cellValue(row, index) {
  return index < row.length ? String(row[index] ?? '').trim() : '';
}

export function parseDailyReportBColumn(bText) {
  const b = String(bText ?? '').trim();
  const m = b.match(/^([\d.]+\D?)\s*x\s*([\d.]+)\s*x\s*([\d.]+)\s+(.+)$/i);
  if (!m) return null;
  return {
    thickness: m[1],
    width: m[2],
    height: m[3],
    englishType: m[4].trim(),
  };
}

export function englishTypeToProductTypeKey(englishType) {
  const lower = String(englishType ?? '').trim().toLowerCase();
  if (!lower) return '';
  if (ENGLISH_TO_TYPE[lower]) return ENGLISH_TO_TYPE[lower];
  for (const [en, key] of Object.entries(ENGLISH_TO_TYPE)) {
    if (lower.includes(en) || en.includes(lower)) return key;
  }
  return '';
}

export function buildProductTypesFromKey(typeKey) {
  const types = {};
  for (const key of Object.keys(TYPE_ENGLISH_NAMES)) types[key] = false;
  if (typeKey) types[typeKey] = true;
  return types;
}

export function buildMachinesFromName(machine) {
  const machines = {};
  for (const name of MOLDING_MACHINE_NAMES) machines[name] = false;
  machines['16 噸 (自動) 啤 機'] = false;
  if (machine) machines[machine] = true;
  return machines;
}

export function parseDailySheetB1Date(rows) {
  const b1 = cellValue(rows[0] || [], 1);
  const m = b1.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (!m) return null;
  const d = String(parseInt(m[1], 10)).padStart(2, '0');
  const mo = String(parseInt(m[2], 10)).padStart(2, '0');
  return { y: m[3], m: mo, d, iso: `${m[3]}-${mo}-${d}` };
}

export async function resolveProductFromDailyRow(parsed, length) {
  const typeKey = englishTypeToProductTypeKey(parsed.englishType);
  if (!typeKey) {
    return { productNo: '', name: '', typeKey: '', ambiguous: false, candidates: [] };
  }
  const matches = await searchPlist({
    type: typeKey,
    t: parsed.thickness,
    w: parsed.width,
    h: parsed.height,
    l: length,
  });
  if (!matches.length) {
    return { productNo: '', name: '', typeKey, ambiguous: false, candidates: [] };
  }
  return {
    productNo: matches[0].code,
    name: matches[0].name,
    typeKey,
    ambiguous: matches.length > 1,
    candidates: matches,
  };
}

function splitOperators(text) {
  return String(text ?? '').split('/').map((s) => s.trim()).filter(Boolean);
}

export function buildProductionRecordFromDailyRow({
  tabName,
  rowNum,
  row,
  machine,
  pageType,
  recordDateIso,
  productInfo,
}) {
  const parsed = parseDailyReportBColumn(cellValue(row, 1));
  const operators = splitOperators(cellValue(row, 0));
  const operator = operators.join('/') || operators[0] || '';
  const length = cellValue(row, 3);
  const qty = cellValue(row, 5);
  let speed = cellValue(row, 11);
  const transferMins = cellValue(row, 8);
  if (!speed && transferMins) speed = '轉機';

  const typeKey = productInfo.typeKey || englishTypeToProductTypeKey(parsed?.englishType);

  const mainLine = {
    operator,
    thickness: parsed?.thickness || '',
    width: parsed?.width || '',
    height: parsed?.height || '',
    length,
    productNo: productInfo.productNo || '',
    name: productInfo.name || '',
    speed,
    load: '',
    start: '',
    finish: '',
    other: productInfo.ambiguous ? '日報取込: 產品編號候選複數' : '',
  };

  const materialLine = {
    orderNo: '',
    thickness1: '',
    width1: '',
    weight: '',
    thickness2: parsed?.thickness || '',
    width2: parsed?.width || '',
    height: parsed?.height || '',
    productNo: productInfo.productNo || '',
    name: productInfo.name || '',
    length,
    qty,
    complete: true,
    oldCoil: false,
    incomplete: false,
  };

  return {
    recordDate: formatRecordDateDisplay(recordDateIso),
    pageType,
    productTypes: buildProductTypesFromKey(typeKey),
    machines: buildMachinesFromName(machine),
    main: mainLine,
    material: materialLine,
    mainLines: [mainLine],
    materialLines: [materialLine],
    dailyReport: { tabName, row: rowNum, machine: machine || '' },
    importedFromDailyReport: true,
    correctionNote: '生產日報取込',
  };
}

export async function scanDailyReportTab(tabName) {
  const rows = await readRange(`${tabName}!A1:L150`);
  const dateInfo = parseDailySheetB1Date(rows);
  const entries = [];

  for (const machine of MOLDING_MACHINE_NAMES) {
    const block = resolveMoldingMachineBlock(rows, machine);
    if (!block) continue;
    for (const rowNum of block.slotRows) {
      const row = rows[rowNum - 1] || [];
      if (isDailyReportRowEmpty(row)) continue;
      entries.push({ tabName, rowNum, row, machine, pageType: 'molding' });
    }
  }

  const autoBlock = resolveAutoMachineBlock(rows);
  if (autoBlock) {
    for (const rowNum of autoBlock.slotRows) {
      const row = rows[rowNum - 1] || [];
      if (isDailyReportRowEmpty(row)) continue;
      entries.push({
        tabName,
        rowNum,
        row,
        machine: '16 噸 (自動) 啤 機',
        pageType: 'auto',
      });
    }
  }

  return { tabName, recordDateIso: dateInfo?.iso || null, dateInfo, entries };
}

export async function buildImportRecordsFromDailyTab(tabName, existingLinks = new Set()) {
  const { recordDateIso, entries } = await scanDailyReportTab(tabName);
  const imported = [];
  const skipped = [];
  const errors = [];

  for (const entry of entries) {
    const linkKey = `${entry.tabName}:${entry.rowNum}`;
    if (existingLinks.has(linkKey)) {
      skipped.push({ row: entry.rowNum, reason: 'already imported' });
      continue;
    }

    const parsed = parseDailyReportBColumn(cellValue(entry.row, 1));
    if (!parsed) {
      skipped.push({ row: entry.rowNum, reason: 'B column not parseable', b: cellValue(entry.row, 1) });
      continue;
    }

    if (!recordDateIso) {
      errors.push({ row: entry.rowNum, error: 'B1 date missing on sheet tab' });
      continue;
    }

    try {
      const productInfo = await resolveProductFromDailyRow(parsed, cellValue(entry.row, 3));
      const record = buildProductionRecordFromDailyRow({
        ...entry,
        recordDateIso,
        productInfo,
      });
      imported.push(record);
    } catch (e) {
      errors.push({ row: entry.rowNum, error: e?.message || String(e) });
    }
  }

  return { tabName, recordDateIso, imported, skipped, errors };
}

export function dailyReportLinkKey(dailyReport) {
  if (!dailyReport?.tabName || !dailyReport?.row) return '';
  return `${dailyReport.tabName}:${dailyReport.row}`;
}

export function collectExistingDailyReportLinks(records) {
  const links = new Set();
  for (const record of records || []) {
    const key = dailyReportLinkKey(record?.dailyReport);
    if (key) links.add(key);
  }
  return links;
}
