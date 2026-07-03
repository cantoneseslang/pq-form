import { getDailyReportSheetsClient, readRange, writeRanges } from './dailyReportSheets.js';

/** 產品種類 → 日報 B 列英語種別 */
export const TYPE_ENGLISH_NAMES = {
  地槽: 'Runner',
  企筒: 'Stud',
  批灰角: 'Aluminium Corner Bead',
  鐵角: 'Angle',
  W角: 'W Angle',
  闊槽: 'Channel',
  C槽: 'C Channel',
  CT企筒打孔: 'CT Stud Punch',
  其他: 'Other',
};

const LUNCH_START = 12 * 60;
const LUNCH_END = 13 * 60;

const MOLDING_MACHINE_PATTERNS = {
  '1號滾壓成型機': /#?1號滾壓成型機/,
  '2號滾壓成型機': /#?2號滾壓成型機/,
  '3號滾壓成型機': /#?3號滾壓成型機/,
  '4號滾壓成型機': /#?4號滾壓成型機/,
  '5號滾壓成型機': /#?5號滾壓成型機/,
};

const AUTO_MACHINE_PATTERN = /16\s*噸\s*\(自動\)\s*啤\s*機/;

/** 成形機 #1/#3/#5 は先頭データ行の C 列に機械名 */
const MACHINE_NAME_ON_FIRST_SLOT = new Set([
  '1號滾壓成型機',
  '3號滾壓成型機',
  '5號滾壓成型機',
]);

export function parseTimeToMinutes(timeStr) {
  const text = String(timeStr ?? '').trim();
  if (!text) return null;
  const match = text.match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return null;
  const h = parseInt(match[1], 10);
  const m = parseInt(match[2], 10);
  if (!Number.isFinite(h) || !Number.isFinite(m) || h < 0 || h > 23 || m < 0 || m > 59) return null;
  return h * 60 + m;
}

function minuteDiff(startMin, endMin) {
  if (startMin === null || endMin === null) return null;
  let diff = endMin - startMin;
  if (diff < 0) diff += 24 * 60;
  return diff;
}

/** 轉機時間：昼休 12:00–13:00 を跨ぐ場合は差し引く */
export function transferMinutes(startStr, finishStr) {
  const start = parseTimeToMinutes(startStr);
  const finish = parseTimeToMinutes(finishStr);
  if (start === null || finish === null) return '';
  let total = minuteDiff(start, finish);
  const overlapStart = Math.max(start, LUNCH_START);
  const overlapEnd = Math.min(finish, LUNCH_END);
  if (overlapEnd > overlapStart) total -= overlapEnd - overlapStart;
  return total > 0 ? total : '';
}

export function selectedProductType(productTypes = {}) {
  for (const [key, label] of Object.entries(TYPE_ENGLISH_NAMES)) {
    if (productTypes[key]) return key;
  }
  const otherText = String(productTypes['其他入力'] ?? '').trim();
  if (productTypes['其他'] && otherText) return '其他';
  if (productTypes['其他']) return '其他';
  return '';
}

export function formatDailyProductName(mainLine, productTypes = {}) {
  const thickness = String(mainLine?.thickness ?? '').trim();
  const width = String(mainLine?.width ?? '').trim();
  const height = String(mainLine?.height ?? '').trim();
  const typeKey = selectedProductType(productTypes);
  const english = TYPE_ENGLISH_NAMES[typeKey] || String(mainLine?.name ?? '').trim() || 'Product';
  if (!thickness || !width || !height) return english;
  return `${thickness} x ${width} x ${height} ${english}`;
}

export function sumMaterialQty(materialLines = []) {
  return materialLines.reduce((sum, line) => {
    const qty = parseInt(String(line?.qty ?? '').trim(), 10);
    return Number.isFinite(qty) && qty > 0 ? sum + qty : sum;
  }, 0);
}

export function countMaterialRolls(materialLines = []) {
  return materialLines.filter((line) => {
    const qty = parseInt(String(line?.qty ?? '').trim(), 10);
    return Number.isFinite(qty) && qty > 0;
  }).length;
}

export function hasProduction(materialLines = [], mainLines = []) {
  if (countMaterialRolls(materialLines) > 0) return true;
  return mainLines.some((line) => String(line?.speed ?? '').trim() !== '轉機');
}

export function buildDailyReportRowValues(mainLines = [], materialLines = [], productTypes = {}) {
  const firstLine = mainLines[0] || {};
  const operators = [];
  let transferTotal = 0;
  let loadTotal = 0;
  let workTotal = 0;
  let lastSpeed = '';

  for (const line of mainLines) {
    const op = String(line?.operator ?? '').trim();
    if (op && !operators.includes(op)) operators.push(op);

    const speed = String(line?.speed ?? '').trim();
    const start = line?.start;
    const finish = line?.finish;
    const load = line?.load;

    if (speed === '轉機') {
      const mins = transferMinutes(start, finish);
      if (mins !== '') transferTotal += mins;
    } else {
      const loadMin = parseTimeToMinutes(load);
      const startMin = parseTimeToMinutes(start);
      const finishMin = parseTimeToMinutes(finish);
      const loadGap = minuteDiff(loadMin, startMin);
      const workGap = minuteDiff(startMin, finishMin);
      if (loadGap !== null) loadTotal += loadGap;
      if (workGap !== null) workTotal += workGap;
      if (speed && speed !== '轉機') lastSpeed = speed;
    }
  }

  const produced = hasProduction(materialLines, mainLines);
  const qty = sumMaterialQty(materialLines);
  const rolls = countMaterialRolls(materialLines);
  const machineLabel = '';

  const values = {
    A: operators.join('/'),
    B: formatDailyProductName(firstLine, productTypes),
    C: machineLabel,
    D: String(firstLine?.length ?? '').trim(),
    F: produced && qty > 0 ? qty : '',
    H: produced && rolls > 0 ? rolls : '',
    I: transferTotal > 0 ? transferTotal : '',
    J: produced && loadTotal > 0 ? loadTotal : '',
    K: produced && workTotal > 0 ? workTotal : '',
    L: produced && lastSpeed ? lastSpeed : '',
  };

  return values;
}

function cellValue(row, index) {
  return index < row.length ? String(row[index] ?? '').trim() : '';
}

function isTotalRow(row) {
  const e = cellValue(row, 4);
  const a = cellValue(row, 0);
  return e.startsWith('Total:') || a.startsWith('Total:');
}

function findBlockByPattern(rows, pattern) {
  let headerRow = -1;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const matched = [0, 1, 2].some((col) => pattern.test(cellValue(row, col)));
    if (matched) {
      headerRow = i;
      break;
    }
  }
  if (headerRow < 0) return null;

  const header = rows[headerRow] || [];
  const headerHasData = cellValue(header, 0) || cellValue(header, 1) || cellValue(header, 3);
  const startIndex = headerHasData ? headerRow : headerRow + 1;

  const slotRows = [];
  for (let i = startIndex; i < rows.length; i++) {
    if (isTotalRow(rows[i])) break;
    slotRows.push(i + 1);
    if (slotRows.length >= 5) break;
  }
  return { headerRow: headerRow + 1, slotRows };
}

export function resolveMoldingMachineBlock(rows, machine) {
  const pattern = MOLDING_MACHINE_PATTERNS[machine];
  if (!pattern) return null;
  return findBlockByPattern(rows, pattern);
}

export function resolveAutoMachineBlock(rows) {
  return findBlockByPattern(rows, AUTO_MACHINE_PATTERN);
}

function isSlotEmpty(row) {
  const a = cellValue(row, 0);
  const b = cellValue(row, 1);
  const d = cellValue(row, 3);
  const f = cellValue(row, 5);
  return !a && !b && !d && !f;
}

function productMatchesRow(row, productLabel) {
  const b = cellValue(row, 1);
  if (!b || !productLabel) return false;
  return b.replace(/\s+/g, ' ').toLowerCase() === productLabel.replace(/\s+/g, ' ').toLowerCase();
}

export function pickTargetRow(rows, block, productLabel, slotIndex = 0) {
  if (!block?.slotRows?.length) return null;

  for (const rowNum of block.slotRows) {
    const row = rows[rowNum - 1] || [];
    if (productMatchesRow(row, productLabel)) return rowNum;
  }

  if (slotIndex >= 0 && slotIndex < block.slotRows.length) {
    return block.slotRows[slotIndex];
  }

  for (const rowNum of block.slotRows) {
    const row = rows[rowNum - 1] || [];
    if (isSlotEmpty(row)) return rowNum;
  }

  return block.slotRows[block.slotRows.length - 1];
}

export function formatDailySheetDate(date) {
  const y = String(date?.y ?? '').trim();
  const m = String(parseInt(date?.m || '0', 10)).padStart(2, '0');
  const d = String(parseInt(date?.d || '0', 10)).padStart(2, '0');
  if (!y || !m || !d || m === '00' || d === '00') return '';
  return `${d}-${m}-${y}`;
}

export function dailySheetTabName(date) {
  const day = parseInt(String(date?.d ?? '').trim(), 10);
  if (!Number.isFinite(day) || day < 1 || day > 31) return '';
  return String(day);
}

function machineLabelFor(pageType, machine) {
  if (pageType === 'auto') return '16 噸 (自動) 啤 機';
  return machine || '';
}

export async function writeDailyReportEntry(payload) {
  const {
    date,
    pageType = 'molding',
    machine,
    productTypes = {},
    mainLines = [],
    materialLines = [],
    slotIndex = 0,
  } = payload;

  const tabName = dailySheetTabName(date);
  if (!tabName) throw new Error('invalid date for daily report tab');

  const values = buildDailyReportRowValues(mainLines, materialLines, productTypes);
  const sheetValues = await readRange(`${tabName}!A1:L100`);
  const block = pageType === 'auto'
    ? resolveAutoMachineBlock(sheetValues)
    : resolveMoldingMachineBlock(sheetValues, machine);

  if (!block) {
    throw new Error(pageType === 'auto'
      ? 'daily report auto machine block not found'
      : `daily report machine block not found: ${machine}`);
  }

  const targetRow = pickTargetRow(sheetValues, block, values.B, slotIndex);
  if (!targetRow) throw new Error('daily report target row not found');

  const isFirstSlot = targetRow === block.slotRows[0];
  const machineName = machineLabelFor(pageType, machine);
  if (isFirstSlot && MACHINE_NAME_ON_FIRST_SLOT.has(machine) && machineName) {
    values.C = machineName;
  }

  const updates = [];
  const dateLabel = formatDailySheetDate(date);
  if (dateLabel) {
    updates.push({ range: `${tabName}!B1`, values: [dateLabel] });
  }

  const columnMap = ['A', 'B', 'C', 'D', 'F', 'H', 'I', 'J', 'K', 'L'];
  for (const col of columnMap) {
    updates.push({ range: `${tabName}!${col}${targetRow}`, values: [values[col] ?? ''] });
  }

  await writeRanges(updates);
  return { tabName, row: targetRow, values, range: `${tabName}!A${targetRow}:L${targetRow}` };
}
