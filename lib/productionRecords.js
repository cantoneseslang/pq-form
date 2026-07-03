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

export async function writeMainDataToSheetRow(main, sheetRow, sheetName) {
  if (!sheetRow || !main) return null;
  const { sheetName: defaultName } = getSheetsClient();
  const name = sheetName || defaultName;
  const values = normalizeRowValues(mainDataToSheetValues(main));
  const writeRange = `${name}!A${sheetRow}:T${sheetRow}`;
  await writeRowA1(writeRange, values);
  return writeRange;
}

export function dbRowToClient(row) {
  return {
    id: row.id,
    recordDate: formatRecordDateDisplay(row.record_date),
    recordDateIso: row.record_date,
    pageType: row.page_type,
    productTypes: row.product_types || {},
    machines: row.machines || {},
    main: row.main_data || {},
    material: row.material_data || {},
    sheetRow: row.sheet_row,
    sheetName: row.sheet_name || 'pq-form',
    correctionNote: row.correction_note || '',
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    correctedAt: row.corrected_at,
    deletedAt: row.deleted_at || null,
  };
}

export function clientToDbInsert(record) {
  return {
    record_date: normalizeRecordDate(record.recordDate),
    page_type: record.pageType,
    product_types: record.productTypes || {},
    machines: record.machines || {},
    main_data: record.main || {},
    material_data: record.material || {},
    sheet_row: record.sheetRow || null,
    sheet_name: record.sheetName || 'pq-form',
  };
}
