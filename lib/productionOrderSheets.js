import { google } from 'googleapis';

const DEFAULT_SHEET_NAME = '202602146';
const DEFAULT_SHEET_GID = 863784501;
export const PRODUCT_BLOCK_START_ROW = 10;
export const PRODUCT_BLOCK_HEIGHT = 5;
export const PRODUCT_FOOTER_ROW = 40;
export const PRODUCT_FOOTER_NOTE_ROW = 41;
export const PRODUCT_FOOTER_SIGN_ROW = 42;
export const DEFAULT_ITEM_COUNT = 6;
export const FOOTER_RETURN_NOTE = '(起貨後請將此單交回寫字樓)';

export const PRODUCT_TYPES = ['企筒', '地槽', '鐵角', '批灰角', 'W角', '闊槽', 'C槽'];

export const MACHINE_OPTIONS = [
  '1號滾壓成型機',
  '2號滾壓成型機',
  '3號滾壓成型機',
  '4號滾壓成型機',
  '5號滾壓成型機',
];

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

export function getProductionOrderSheetsClient() {
  const sa = getServiceAccount();
  const scopes = ['https://www.googleapis.com/auth/spreadsheets'];
  const jwt = new google.auth.JWT(sa.client_email, undefined, sa.private_key, scopes);
  const sheets = google.sheets({ version: 'v4', auth: jwt });
  const spreadsheetId = process.env.PRODUCTION_ORDER_SHEET_ID;
  const sheetName = process.env.PRODUCTION_ORDER_SHEET_NAME || DEFAULT_SHEET_NAME;
  if (!spreadsheetId) throw new Error('PRODUCTION_ORDER_SHEET_ID is not set');
  return { sheets, spreadsheetId, sheetName };
}

export async function resolveProductionOrderSheetName(sheets, spreadsheetId, preferredName) {
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  const tabs = (res.data.sheets || []).map((s) => s.properties).filter(Boolean);
  if (preferredName) {
    const exact = tabs.find((t) => t.title === preferredName);
    if (exact) return exact.title;
  }
  const gid = parseInt(process.env.PRODUCTION_ORDER_SHEET_GID || String(DEFAULT_SHEET_GID), 10);
  const byGid = tabs.find((t) => t.sheetId === gid);
  if (byGid) return byGid.title;
  const names = tabs.map((t) => t.title).join(', ');
  throw new Error(`Production order sheet tab not found (${preferredName || gid}). Available: ${names}`);
}

export async function listProductionOrderSheetTabs() {
  const { sheets, spreadsheetId } = getProductionOrderSheetsClient();
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  return (res.data.sheets || []).map((s) => ({
    title: s.properties?.title || '',
    sheetId: s.properties?.sheetId,
  }));
}

function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function cellRef(sheetName, a1) {
  return `${quoteSheetName(sheetName)}!${a1}`;
}

function trimText(value) {
  return String(value ?? '').trim();
}

export function formatSheetDate(value) {
  const text = trimText(value);
  if (!text) return '';
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) return `${iso[1]}/${parseInt(iso[2], 10)}/${parseInt(iso[3], 10)}`;
  return text;
}

export function formatSheetDateChinese(value) {
  const text = trimText(value);
  if (!text) return '';
  const slash = text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (slash) {
    return `${slash[1]}年${parseInt(slash[2], 10)}月${parseInt(slash[3], 10)}日`;
  }
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    return `${iso[1]}年${parseInt(iso[2], 10)}月${parseInt(iso[3], 10)}日`;
  }
  return text;
}

function footerDateSource(header) {
  return trimText(header.deliveryDate) || trimText(header.orderDate);
}

export function formatLength(value) {
  const n = trimText(value);
  if (!n) return '';
  return `${n}mm長`;
}

export function formatMachine(machineLabel) {
  const text = trimText(machineLabel);
  const m = text.match(/^(\d+)號/);
  if (m) return `(#${m[1]})`;
  return text;
}

export function formatWidth(value) {
  const n = trimText(value);
  if (!n) return '';
  return `${n}mm`;
}

export function formatHeight(value) {
  const n = trimText(value);
  if (!n) return '';
  return `${n}mm`;
}

export function formatThickness(value) {
  const t = trimText(value);
  if (!t) return '';
  const num = parseFloat(t);
  const display = Number.isFinite(num) ? String(num) : t;
  return `(${display}mm厚)`;
}

export function formatQuantity(value) {
  const n = trimText(value);
  if (!n) return '';
  return `${n}支`;
}

function productHasAnyValue(product) {
  return [
    product.type, product.machine, product.thickness, product.width, product.height,
    product.length, product.quantity, product.productCode, product.productName,
    product.materialWidth, product.packagingNote, product.bandNote,
  ].some((v) => trimText(v));
}

function validateProduct(product, itemNo) {
  const errors = [];
  const prefix = `項目${itemNo}`;
  if (!productHasAnyValue(product)) return errors;

  if (!trimText(product.type)) errors.push(`${prefix}：產品種類為必填`);
  else if (!PRODUCT_TYPES.includes(trimText(product.type))) errors.push(`${prefix}：產品種類無效`);

  if (!trimText(product.machine)) errors.push(`${prefix}：生產機械名稱為必填`);
  else if (!MACHINE_OPTIONS.includes(trimText(product.machine))) errors.push(`${prefix}：生產機械名稱無效`);

  if (!trimText(product.thickness)) errors.push(`${prefix}：材料厚度為必填`);
  if (!trimText(product.width)) errors.push(`${prefix}：闊度為必填`);
  if (!trimText(product.height)) errors.push(`${prefix}：高度為必填`);
  if (!trimText(product.length)) errors.push(`${prefix}：長度為必填`);
  if (!trimText(product.quantity)) errors.push(`${prefix}：數量為必填`);

  return errors;
}

export function normalizeProductsPayload(payload) {
  if (Array.isArray(payload?.products)) return payload.products;
  if (payload?.product) return [payload.product];
  return [];
}

export function validateProductionOrderPayload(payload) {
  const errors = [];
  const header = payload?.header || {};
  const products = normalizeProductsPayload(payload);

  products.forEach((product, index) => {
    errors.push(...validateProduct(product || {}, index + 1));
  });

  if (products.filter((p) => productHasAnyValue(p)).length === 0) {
    errors.push('請至少填寫一個項目');
  }

  return { valid: errors.length === 0, errors, header, products };
}

export function blockStartRowForItem(itemNo, baseRow = PRODUCT_BLOCK_START_ROW) {
  return baseRow + (itemNo - 1) * PRODUCT_BLOCK_HEIGHT;
}

function formatCheckMark(value) {
  return value ? '✓' : '';
}

export function buildProductBlockUpdates(sheetName, product, itemNo, blockStartRow) {
  const r = blockStartRow;
  const empty = {
    type: '', machine: '', thickness: '', width: '', height: '', length: '',
    quantity: '', productCode: '', productName: '', materialWidth: '', packagingNote: '', bandNote: '',
    cuttingComplete: false, packagingComplete: false, punchingComplete: false,
  };
  const p = productHasAnyValue(product) ? product : empty;

  return [
    { range: cellRef(sheetName, `A${r}`), values: [String(itemNo)] },
    { range: cellRef(sheetName, `C${r}`), values: [trimText(p.productCode)] },
    { range: cellRef(sheetName, `F${r}`), values: [trimText(p.productName)] },
    { range: cellRef(sheetName, `H${r}`), values: [formatQuantity(p.quantity)] },
    { range: cellRef(sheetName, `B${r + 1}`), values: [trimText(p.type)] },
    { range: cellRef(sheetName, `E${r + 1}`), values: [formatLength(p.length)] },
    { range: cellRef(sheetName, `F${r + 1}`), values: [formatMachine(p.machine)] },
    { range: cellRef(sheetName, `B${r + 2}`), values: [formatWidth(p.width)] },
    { range: cellRef(sheetName, `E${r + 2}`), values: [formatHeight(p.height)] },
    { range: cellRef(sheetName, `F${r + 2}`), values: [formatThickness(p.thickness)] },
    { range: cellRef(sheetName, `C${r + 3}`), values: [trimText(p.materialWidth)] },
    { range: cellRef(sheetName, `E${r + 3}`), values: [trimText(p.packagingNote)] },
    { range: cellRef(sheetName, `F${r + 3}`), values: [trimText(p.bandNote)] },
    { range: cellRef(sheetName, `C${r + 4}`), values: [formatCheckMark(p.cuttingComplete)] },
    { range: cellRef(sheetName, `F${r + 4}`), values: [formatCheckMark(p.packagingComplete)] },
    { range: cellRef(sheetName, `G${r + 4}`), values: [formatCheckMark(p.punchingComplete)] },
  ];
}

export function buildProductionOrderUpdates(sheetName, header, products, itemCount = DEFAULT_ITEM_COUNT) {
  const updates = [
    { range: cellRef(sheetName, 'B3'), values: [trimText(header.deliveryNoteNo)] },
    { range: cellRef(sheetName, 'E3'), values: [trimText(header.customerNo)] },
    { range: cellRef(sheetName, 'G3'), values: [trimText(header.orderingCompany)] },
    { range: cellRef(sheetName, 'B4'), values: [formatSheetDate(header.deliveryDate)] },
    { range: cellRef(sheetName, 'E4'), values: [formatSheetDate(header.orderDate)] },
    { range: cellRef(sheetName, 'B6'), values: [trimText(header.estimatedProductionPeriod)] },
    { range: cellRef(sheetName, 'E6'), values: [formatSheetDate(header.completionDate)] },
    { range: cellRef(sheetName, 'B7'), values: [trimText(header.personInCharge)] },
    { range: cellRef(sheetName, 'E7'), values: [trimText(header.signature)] },
  ];

  for (let itemNo = 1; itemNo <= itemCount; itemNo += 1) {
    const blockStartRow = blockStartRowForItem(itemNo);
    updates.push(...buildProductBlockUpdates(sheetName, products[itemNo - 1] || {}, itemNo, blockStartRow));
  }

  updates.push({
    range: cellRef(sheetName, `C${PRODUCT_FOOTER_ROW}`),
    values: [trimText(header.deliveryDestination)],
  });
  updates.push({
    range: cellRef(sheetName, `B${PRODUCT_FOOTER_NOTE_ROW}`),
    values: [FOOTER_RETURN_NOTE],
  });
  updates.push({
    range: cellRef(sheetName, `E${PRODUCT_FOOTER_SIGN_ROW}`),
    values: [trimText(header.preparerSignature)],
  });
  updates.push({
    range: cellRef(sheetName, `G${PRODUCT_FOOTER_SIGN_ROW}`),
    values: [formatSheetDateChinese(footerDateSource(header))],
  });

  return updates;
}

export async function writeProductionOrder(payload) {
  const { valid, errors, header, products } = validateProductionOrderPayload(payload);
  if (!valid) {
    const err = new Error(errors.join('；'));
    err.validationErrors = errors;
    throw err;
  }

  const { sheets, spreadsheetId, sheetName: preferredName } = getProductionOrderSheetsClient();
  const sheetName = await resolveProductionOrderSheetName(sheets, spreadsheetId, preferredName);
  const itemCount = Math.max(DEFAULT_ITEM_COUNT, products.length);
  const paddedProducts = Array.from({ length: itemCount }, (_, i) => products[i] || {});
  const updates = buildProductionOrderUpdates(sheetName, header, paddedProducts, itemCount);
  const data = updates.map((u) => ({ range: u.range, values: [u.values] }));

  const res = await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: { valueInputOption: 'USER_ENTERED', data },
  });

  return {
    updatedRanges: updates.map((u) => u.range),
    updatedCells: updates.length,
    itemCount,
    spreadsheetId,
    sheetName,
    apiResponse: res.data,
  };
}
