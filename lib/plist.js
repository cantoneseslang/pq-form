import { readRange } from './sheets.js';

const PLIST_SHEET_NAME = process.env.PQFORM_PLIST_SHEET_NAME || 'PQ-Form-plist';
const CACHE_TTL_MS = 5 * 60 * 1000;

let cache = { rows: null, loadedAt: 0 };

export function normalizeNum(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  const num = parseFloat(text);
  return Number.isFinite(num) ? String(num) : text;
}

const INCLUDES_MATCH_TYPES = new Set(['闊槽', 'C槽', '批灰角']);

export function matchCpdesc1(cpdesc1, selectedType, otherText = '') {
  const cp = String(cpdesc1 ?? '').trim();
  if (!selectedType) return false;
  if (selectedType === '其他') {
    const other = String(otherText ?? '').trim();
    return !!other && (cp === other || cp.includes(other));
  }
  if (INCLUDES_MATCH_TYPES.has(selectedType)) {
    // plist 側は「L形批灰角」「批灰角(鋁)」「C 槽」のように派生名・スペースあり
    const cpCompact = cp.replace(/\s+/g, '');
    const typeCompact = selectedType.replace(/\s+/g, '');
    return cp.includes(selectedType) || cpCompact.includes(typeCompact);
  }
  return cp === selectedType;
}

export function extractTypeLabel(name, cpdesc1) {
  const m = String(name ?? '').match(/^[\d.]+x[\d.]+x[\d.]+\s+(.+?)\s+[\d.]+mm/i);
  if (m) return m[1].trim();
  return String(cpdesc1 ?? '').trim();
}

export function formatPlistDisplayName(row) {
  const [, name, , cpdesc1, wdesc1, hdesc1, ldesc1, tdesc1] = row;
  const typeLabel = extractTypeLabel(name, cpdesc1);
  const t = normalizeNum(tdesc1);
  const w = normalizeNum(wdesc1);
  const h = normalizeNum(hdesc1);
  const l = normalizeNum(ldesc1);
  if (!t && !w && !h && !l) return String(name ?? '').trim();
  return `${t}x${w}x${h} ${typeLabel} ${l}mm`;
}

export async function loadPlistRows() {
  if (cache.rows && Date.now() - cache.loadedAt < CACHE_TTL_MS) {
    return cache.rows;
  }
  const values = await readRange(`${PLIST_SHEET_NAME}!A:H`);
  cache = {
    rows: (values || []).slice(1).filter((row) => row?.[0]),
    loadedAt: Date.now(),
  };
  return cache.rows;
}

export async function searchPlist({ type, t, w, h, l, other = '' }) {
  const rows = await loadPlistRows();
  const nt = normalizeNum(t);
  const nw = normalizeNum(w);
  const nh = normalizeNum(h);
  const nl = normalizeNum(l);

  const seen = new Set();
  const matches = [];

  for (const row of rows) {
    const [code, name, , cpdesc1, wdesc1, hdesc1, ldesc1, tdesc1] = row;
    if (!matchCpdesc1(cpdesc1, type, other)) continue;
    if (normalizeNum(tdesc1) !== nt) continue;
    if (normalizeNum(wdesc1) !== nw) continue;
    if (normalizeNum(hdesc1) !== nh) continue;
    if (normalizeNum(ldesc1) !== nl) continue;

    const key = `${code}::${name}`;
    if (seen.has(key)) continue;
    seen.add(key);
    matches.push({
      code: String(code ?? '').trim(),
      name: formatPlistDisplayName(row),
    });
  }

  return matches;
}

export async function getPlistLengthHints({ type, t, w, h, other = '' }) {
  const rows = await loadPlistRows();
  const nt = normalizeNum(t);
  const nw = normalizeNum(w);
  const nh = normalizeNum(h);
  const lengths = new Set();

  for (const row of rows) {
    const [, , , cpdesc1, wdesc1, hdesc1, ldesc1, tdesc1] = row;
    if (!matchCpdesc1(cpdesc1, type, other)) continue;
    if (normalizeNum(tdesc1) !== nt) continue;
    if (normalizeNum(wdesc1) !== nw) continue;
    if (normalizeNum(hdesc1) !== nh) continue;
    const l = normalizeNum(ldesc1);
    if (l) lengths.add(l);
  }

  return [...lengths].sort((a, b) => parseFloat(a) - parseFloat(b));
}

export async function getThicknessList() {
  const rows = await loadPlistRows();
  const set = new Set();
  for (const row of rows) {
    const t = normalizeNum(row[7]);
    if (t) set.add(t);
  }
  return [...set].sort((a, b) => parseFloat(a) - parseFloat(b));
}
