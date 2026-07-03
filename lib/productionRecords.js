import { normalizeRowValues, writeRowA1, getSheetsClient } from './sheets.js';

export function normalizeRecordDate(dateStr) {
  const text = String(dateStr ?? '').trim();
  if (!text) return '';
  const m = text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!m) return text;
  const y = m[1];
  const mo = String(parseInt(m[2], 10)).padStart(2, '0');
  const d = String(parseInt(m[3], 10)).padStart(2, '0');
  return `${y}-${mo}-${d}`;
}

export function formatRecordDateDisplay(isoDate) {
  if (!isoDate) return '';
  const m = String(isoDate).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return isoDate;
  return `${m[1]}/${m[2]}/${m[3]}`;
}

export function mainDataToSheetValues(main) {
  if (!main) return null;
  return [
    main.load || '',
    main.start || '',
    String(main.productNo || '').toUpperCase(),
    main.thickness || '',
    main.width || '',
    main.height || '',
    main.name || '',
    '', '', '',
    main.length || '',
    main.length_tolerance || '',
    main.section_size || '',
    main.left_right_bend || '',
    main.up_down_bend || '',
    main.twist || '',
    main.operator || '',
    main.finish || '',
    main.speed || '',
    main.other || '',
  ];
}

export function packMainData(main, mainLines) {
  const summary = { ...(main || {}) };
  delete summary.lines;
  const lines = (Array.isArray(mainLines) && mainLines.length ? mainLines : [summary]).map((line) => {
    const copy = { ...line };
    delete copy.lines;
    return copy;
  });
  return { ...summary, lines };
}

export function unpackMainData(mainData) {
  if (!mainData || typeof mainData !== 'object') {
    return { main: {}, mainLines: [] };
  }
  const { lines, ...summary } = mainData;
  const mainLines = Array.isArray(lines) && lines.length
    ? lines.map((line) => ({ ...line }))
    : (summary.productNo || summary.load ? [{ ...summary }] : []);
  return { main: summary, mainLines };
}

export function packMaterialData(material, materialLines, sheetRows, dailyReport = null) {
  const summary = { ...(material || {}) };
  delete summary.lines;
  delete summary.sheetRows;
  delete summary.dailyReport;
  const lines = (Array.isArray(materialLines) && materialLines.length ? materialLines : [summary]).map((line) => {
    const copy = { ...line };
    delete copy.lines;
    delete copy.sheetRows;
    delete copy.dailyReport;
    return copy;
  });
  const resolvedDailyReport = dailyReport || summary.dailyReport || null;
  return {
    ...summary,
    lines,
    sheetRows: Array.isArray(sheetRows) ? sheetRows.filter(Boolean) : [],
    ...(resolvedDailyReport ? { dailyReport: resolvedDailyReport } : {}),
  };
}

export function unpackMaterialData(materialData) {
  if (!materialData || typeof materialData !== 'object') {
    return { material: {}, materialLines: [], sheetRows: [] };
  }
  const { lines, sheetRows, dailyReport, ...summary } = materialData;
  const materialLines = Array.isArray(lines) && lines.length
    ? lines.map((line) => ({ ...line }))
    : (summary.orderNo || summary.qty ? [{ ...summary }] : []);
  return {
    material: summary,
    materialLines,
    sheetRows: Array.isArray(sheetRows) ? sheetRows : [],
    dailyReport: dailyReport || null,
  };
}

export async function writeMainDataToSheetRow(main, sheetRow, sheetName) {
  if (!sheetRow || !main) return null;
  const { sheetName: defaultName } = getSheetsClient();
  const name = sheetName || defaultName;
  const values = normalizeRowValues(mainDataToSheetValues(main));
  const writeRange = `${name}!A${sheetRow}:T${sheetRow}`;
  await writeRowA1(writeRange, values);
  return writeRange;
}

export async function writeMainLinesToSheet(mainLines, sheetRows, sheetName) {
  if (!Array.isArray(mainLines)) return [];
  const written = [];
  for (let i = 0; i < mainLines.length; i++) {
    const sheetRow = sheetRows?.[i];
    if (sheetRow && mainLines[i]) {
      written.push(await writeMainDataToSheetRow(mainLines[i], sheetRow, sheetName));
    }
  }
  return written;
}

export function dbRowToClient(row) {
  const { main, mainLines } = unpackMainData(row.main_data || {});
  const { material, materialLines, sheetRows, dailyReport } = unpackMaterialData(row.material_data || {});
  return {
    id: row.id,
    recordDate: formatRecordDateDisplay(row.record_date),
    recordDateIso: row.record_date,
    pageType: row.page_type,
    productTypes: row.product_types || {},
    machines: row.machines || {},
    main,
    material,
    mainLines,
    materialLines,
    sheetRows: sheetRows.length ? sheetRows : (row.sheet_row ? [row.sheet_row] : []),
    sheetRow: sheetRows[0] || row.sheet_row,
    sheetName: row.sheet_name || 'pq-form',
    dailyReport: dailyReport || null,
    correctionNote: row.correction_note || '',
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    correctedAt: row.corrected_at,
    deletedAt: row.deleted_at || null,
  };
}

export function getMachineFromMachines(machines = {}) {
  for (const [name, on] of Object.entries(machines)) {
    if (on) return name;
  }
  return '';
}

export function parseRecordDateForDaily(isoDate) {
  const m = String(isoDate ?? '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return { y: m[1], m: m[2], d: m[3] };
}

export function clientToDbInsert(record) {
  const mainLines = record.mainLines || (record.main ? [record.main] : []);
  const materialLines = record.materialLines || (record.material ? [record.material] : []);
  const sheetRows = record.sheetRows || (record.sheetRow ? [record.sheetRow] : []);
  return {
    record_date: normalizeRecordDate(record.recordDate),
    page_type: record.pageType,
    product_types: record.productTypes || {},
    machines: record.machines || {},
    main_data: packMainData(record.main || mainLines[0] || {}, mainLines),
    material_data: packMaterialData(
      record.material || materialLines[0] || {},
      materialLines,
      sheetRows,
      record.dailyReport || null,
    ),
    sheet_row: sheetRows[0] || record.sheetRow || null,
    sheet_name: record.sheetName || 'pq-form',
  };
}
