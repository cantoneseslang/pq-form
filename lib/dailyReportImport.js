import { readRange } from './dailyReportSheets.js';
import {
  TYPE_ENGLISH_NAMES,
  resolveMoldingMachineBlock,
  resolveAutoMachineBlock,
  isDailyReportRowEmpty,
  MOLDING_MACHINE_NAMES,
} from './dailyReport.js';
import { searchPlist, thicknessForProductLookup, buildProvisionalProductName, applyRecordedThicknessToProductName, NOT_FOUND_PRODUCT_CODE } from './plist.js';
import {
  formatRecordDateDisplay,
  dbRowToClient,
  packMainData,
  packMaterialData,
  unpackMaterialData,
} from './productionRecords.js';

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
    thickness: m[1].trim(),
    width: m[2].trim(),
    height: m[3].trim(),
    englishType: m[4].trim(),
  };
}

export function normalizeEnglishTypeText(englishType) {
  return String(englishType ?? '')
    .trim()
    .toLowerCase()
    .replace(/alumi[uw]*i?u?m/g, 'aluminium')
    .replace(/\s+/g, ' ');
}

export function englishTypeToProductTypeKey(englishType) {
  const lower = normalizeEnglishTypeText(englishType);
  if (!lower) return '';
  if (ENGLISH_TO_TYPE[lower]) return ENGLISH_TO_TYPE[lower];
  for (const [en, key] of Object.entries(ENGLISH_TO_TYPE)) {
    if (lower.includes(en) || en.includes(lower)) return key;
  }
  if (/corner bead|l-?bead/.test(lower)) return '批灰角';
  if (/\bstud\b/.test(lower)) return '企筒';
  if (/\brunner\b/.test(lower)) return '地槽';
  if (/\bw angle\b/.test(lower)) return 'W角';
  if (/\bc channel\b/.test(lower)) return 'C槽';
  if (/\bchannel\b/.test(lower)) return '闊槽';
  if (/\bangle\b/.test(lower)) return '鐵角';
  return '';
}

export function inferProductTypeKeyFromName(name) {
  const text = String(name ?? '');
  for (const key of Object.keys(TYPE_ENGLISH_NAMES)) {
    if (text.includes(key)) return key;
  }
  const lower = text.toLowerCase();
  if (/corner bead|l-?bead/.test(lower)) return '批灰角';
  if (/\bstud\b/.test(lower)) return '企筒';
  if (/\brunner\b/.test(lower)) return '地槽';
  if (/\bw angle\b/.test(lower)) return 'W角';
  if (/\bc channel\b/.test(lower)) return 'C槽';
  if (/\bchannel\b/.test(lower)) return '闊槽';
  if (/\bangle\b/.test(lower)) return '鐵角';
  return '';
}

function isEnglishOnlyImportName(name) {
  const text = String(name ?? '').trim();
  if (!text) return false;
  return !/[\u4e00-\u9fff]/.test(text);
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
  const spec = {
    thickness: parsed.thickness,
    width: parsed.width,
    height: parsed.height,
    length: String(length ?? '').trim(),
  };
  const lookupT = thicknessForProductLookup(spec.thickness);

  const resolveWithType = async (typeKey) => {
    if (!typeKey) return null;
    const matches = await searchPlist({
      type: typeKey,
      t: lookupT,
      w: spec.width,
      h: spec.height,
      l: spec.length,
    });
    if (!matches.length) return null;
    return {
      productNo: matches[0].code,
      name: applyRecordedThicknessToProductName(matches[0].name, spec.thickness),
      typeKey,
      ambiguous: matches.length > 1,
      candidates: matches,
      plistMiss: false,
    };
  };

  let typeKey = englishTypeToProductTypeKey(parsed.englishType);
  let resolved = await resolveWithType(typeKey);
  if (resolved) return resolved;

  const inferred = inferProductTypeKeyFromName(parsed.englishType)
    || inferProductTypeKeyFromName(buildProvisionalProductName(typeKey, spec));
  if (inferred && inferred !== typeKey) {
    typeKey = inferred;
    resolved = await resolveWithType(typeKey);
    if (resolved) return resolved;
  }

  typeKey = typeKey || inferred;
  if (typeKey) {
    return {
      productNo: NOT_FOUND_PRODUCT_CODE,
      name: buildProvisionalProductName(typeKey, spec),
      typeKey,
      ambiguous: false,
      candidates: [],
      plistMiss: true,
    };
  }

  return {
    productNo: '',
    name: buildProvisionalProductName(parsed.englishType || '', spec),
    typeKey: '',
    ambiguous: false,
    candidates: [],
    plistMiss: true,
  };
}

const KNOWN_OPERATORS = ['達', '嫻', '林'];

function isDittoMark(text) {
  const t = String(text ?? '').trim();
  if (!t) return false;
  if (t === '同上') return true;
  return t === '"' || t === '“' || t === '”' || t === '〃' || t === '″' || /^[""〃″]+$/u.test(t);
}

function splitOperators(text) {
  return String(text ?? '').split('/').map((s) => s.trim()).filter(Boolean);
}

export function findKnownOperator(text) {
  const raw = String(text ?? '').trim();
  if (!raw || isDittoMark(raw)) return '';
  const parts = splitOperators(raw).filter((part) => KNOWN_OPERATORS.includes(part));
  if (parts.length) return parts.join('/');
  if (KNOWN_OPERATORS.includes(raw)) return raw;
  return '';
}

function resolveBlockForMachine(rows, machine) {
  if (MOLDING_MACHINE_NAMES.includes(machine)) {
    return resolveMoldingMachineBlock(rows, machine);
  }
  return resolveAutoMachineBlock(rows);
}

export function resolveDailyReportOperator(rows, rowNum, machine) {
  const block = resolveBlockForMachine(rows, machine);
  const raw = cellValue(rows[rowNum - 1] || [], 0);
  if (!isDittoMark(raw)) {
    return findKnownOperator(raw) || raw;
  }
  if (!block) return '';
  const idx = block.slotRows.indexOf(rowNum);
  if (idx <= 0) return '';
  for (let i = idx - 1; i >= 0; i -= 1) {
    const prevRaw = cellValue(rows[block.slotRows[i] - 1] || [], 0);
    const op = findKnownOperator(prevRaw);
    if (op) return op;
    if (prevRaw && !isDittoMark(prevRaw)) break;
  }
  return '';
}

export function resolveOperatorFromDbPeers(record, recordsForTab) {
  const dr = record?.dailyReport;
  if (!dr?.machine || !dr?.row) return '';
  const peers = (recordsForTab || [])
    .filter((item) => item.dailyReport?.machine === dr.machine)
    .sort((a, b) => Number(a.dailyReport?.row || 0) - Number(b.dailyReport?.row || 0));
  const idx = peers.findIndex((item) => item.id === record.id);
  if (idx <= 0) return '';
  for (let i = idx - 1; i >= 0; i -= 1) {
    const prevOp = findKnownOperator(peers[i].main?.operator);
    if (prevOp) return prevOp;
    const raw = String(peers[i].main?.operator ?? '').trim();
    if (raw && !isDittoMark(raw)) break;
  }
  return '';
}

export function buildProductionRecordFromDailyRow({
  tabName,
  rowNum,
  row,
  machine,
  pageType,
  recordDateIso,
  productInfo,
  operator: resolvedOperator,
}) {
  const parsed = parseDailyReportBColumn(cellValue(row, 1));
  const operator = String(resolvedOperator ?? '').trim()
    || findKnownOperator(cellValue(row, 0))
    || splitOperators(cellValue(row, 0)).join('/')
    || '';
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

  return { tabName, recordDateIso: dateInfo?.iso || null, dateInfo, entries, rows };
}

export async function buildImportRecordsFromDailyTab(tabName, existingLinks = new Set()) {
  const { recordDateIso, entries, rows } = await scanDailyReportTab(tabName);
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
      const operator = resolveDailyReportOperator(rows, entry.rowNum, entry.machine);
      const record = buildProductionRecordFromDailyRow({
        ...entry,
        recordDateIso,
        productInfo,
        operator,
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

function applyProductInfoToLines(lines, productInfo, { refreshName = false } = {}) {
  return (lines || []).map((line) => {
    const copy = { ...line };
    const currentName = String(copy.name ?? '').trim();
    const currentCode = String(copy.productNo ?? '').trim();
    const nextName = String(productInfo.name ?? '').trim();
    const nextCode = String(productInfo.productNo ?? '').trim();
    const shouldRefreshName = refreshName
      || !currentName
      || (isEnglishOnlyImportName(currentName) && /[\u4e00-\u9fff]/.test(nextName));
    if (shouldRefreshName && nextName) copy.name = productInfo.name;
    if ((!currentCode || currentCode === NOT_FOUND_PRODUCT_CODE) && nextCode) {
      copy.productNo = productInfo.productNo;
    }
    return copy;
  });
}

export async function repairImportedProductNamesForTab(tabName, supabase) {
  const { recordDateIso, entries, rows: dailySheetRows } = await scanDailyReportTab(tabName);
  const entryByRow = new Map(entries.map((entry) => [entry.rowNum, entry]));

  const { data: allDbRows, error: fetchAllError } = await supabase
    .from('pq_production_records')
    .select('*')
    .is('deleted_at', null);
  if (fetchAllError) throw fetchAllError;

  const dbRows = (allDbRows || []).filter((row) => {
    const { dailyReport } = unpackMaterialData(row.material_data || {});
    return dailyReport && String(dailyReport.tabName) === String(tabName);
  });
  const clientRecords = dbRows.map((row) => dbRowToClient(row));

  if (!dbRows.length) {
    return { tabName, recordDateIso: recordDateIso || null, repaired: [], skipped: [{ reason: 'no records for tab' }] };
  }

  const repaired = [];
  const skipped = [];

  for (const row of dbRows || []) {
    const record = dbRowToClient(row);
    const dr = record.dailyReport;
    if (!dr || String(dr.tabName) !== String(tabName)) continue;

    const entry = entryByRow.get(Number(dr.row));
    const parsed = entry ? parseDailyReportBColumn(cellValue(entry.row, 1)) : null;
    const productInfo = parsed
      ? await resolveProductFromDailyRow(parsed, cellValue(entry.row, 3))
      : { productNo: record.main?.productNo || '', name: record.main?.name || '', typeKey: '' };
    let resolvedOperator = resolveOperatorFromDbPeers(record, clientRecords);
    if (entry && dailySheetRows?.length) {
      const fromSheet = resolveDailyReportOperator(dailySheetRows, entry.rowNum, entry.machine);
      if (fromSheet) resolvedOperator = fromSheet;
    }
    if (entry && !parsed) {
      skipped.push({ id: record.id, row: dr.row, reason: 'B column not parseable' });
      continue;
    }

    const mainName = String(record.main?.name ?? '').trim();
    const mainCode = String(record.main?.productNo ?? '').trim();
    const mainOperator = String(record.main?.operator ?? '').trim();
    const needsType = !Object.values(record.productTypes || {}).some(Boolean) && !!productInfo.typeKey;
    const needsName = !mainName || isEnglishOnlyImportName(mainName);
    const needsCode = !mainCode || mainCode === NOT_FOUND_PRODUCT_CODE;
    const needsOperator = isDittoMark(mainOperator) || mainOperator === '"' || (
      resolvedOperator && resolvedOperator !== mainOperator
    );
    if (!needsType && !needsName && !needsCode && !needsOperator) {
      skipped.push({ id: record.id, row: dr.row, reason: 'already complete' });
      continue;
    }

    const mainLines = applyProductInfoToLines(
      record.mainLines?.length ? record.mainLines : [record.main],
      productInfo,
      { refreshName: needsName },
    ).map((line) => ({
      ...line,
      operator: isDittoMark(line.operator)
        ? (resolvedOperator || '')
        : (resolvedOperator || line.operator || ''),
    }));
    const materialLines = applyProductInfoToLines(
      record.materialLines?.length ? record.materialLines : [record.material],
      productInfo,
      { refreshName: needsName },
    );
    const productTypes = productInfo.typeKey
      ? buildProductTypesFromKey(productInfo.typeKey)
      : (record.productTypes || {});
    const { material, sheetRows: storedSheetRows, dailyReport, importedFromDailyReport } = unpackMaterialData(row.material_data || {});
    const materialForPack = {
      ...material,
      productNo: materialLines[0]?.productNo ?? material.productNo,
      name: materialLines[0]?.name ?? material.name,
      ...(importedFromDailyReport ? { importedFromDailyReport: true } : {}),
    };

    const { error: updateError } = await supabase
      .from('pq_production_records')
      .update({
        product_types: productTypes,
        main_data: packMainData(mainLines[0] || {}, mainLines),
        material_data: packMaterialData(
          materialForPack,
          materialLines,
          storedSheetRows,
          dailyReport || record.dailyReport,
        ),
        updated_at: new Date().toISOString(),
      })
      .eq('id', record.id);
    if (updateError) throw updateError;

    repaired.push({
      id: record.id,
      row: dr.row,
      name: mainLines[0]?.name || '',
      productNo: mainLines[0]?.productNo || '',
      operator: mainLines[0]?.operator || '',
    });
  }

  return { tabName, recordDateIso, repaired, skipped };
}
