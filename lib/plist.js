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

export const PRODUCT_TYPES = ['企筒', '地槽', '鐵角', '批灰角', 'W角', '闊槽', 'C槽'];

export function thicknessMatchesPlist(requestedT, plistT) {
  const a = formatThicknessNum(requestedT);
  const b = formatThicknessNum(plistT);
  if (!a || !b) return false;
  if (a === b) return true;
  if (thicknessForProductLookup(a) === b) return true;
  if (a === thicknessForProductLookup(b)) return true;
  return false;
}

export function lengthMatchesPlist(requestedL, plistL) {
  const a = normalizeNum(requestedL);
  const b = normalizeNum(plistL);
  if (!a || !b) return false;
  return parseFloat(a) === parseFloat(b);
}

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

export function parseNameDims(name) {
  const m = String(name ?? '').trim().match(/^([\d.]+)x([\d.]+)x([\d.]+)/i);
  if (!m) return null;
  return {
    t: normalizeNum(m[1]),
    w: normalizeNum(m[2]),
    h: normalizeNum(m[3]),
  };
}

/** WDESC1/HDESC1 または PQ-FORM-DESC 先頭の TxWxH（例: 0.4x13x32）で闊度・高度を照合 */
export function matchRowWidthHeight(row, nw, nh) {
  const [, name, , , wdesc1, hdesc1] = row;
  if (normalizeNum(wdesc1) === nw && normalizeNum(hdesc1) === nh) return true;
  const parsed = parseNameDims(name);
  return !!(parsed && parsed.w === nw && parsed.h === nh);
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
  const parsed = parseNameDims(name);
  // 批灰角系は plist 列(E/F)と表示名の寸法がずれることがある → 名称側を優先
  if (
    parsed &&
    (parsed.w !== w || parsed.h !== h) &&
    matchCpdesc1(cpdesc1, '批灰角')
  ) {
    const displayT = parsed.t || t;
    return `${displayT}x${parsed.w}x${parsed.h} ${typeLabel} ${l}mm`;
  }
  if (!t && !w && !h && !l) return String(name ?? '').trim();
  return `${t}x${w}x${h} ${typeLabel} ${l}mm`;
}

export async function loadPlistRows() {
  if (cache.rows && Date.now() - cache.loadedAt < CACHE_TTL_MS) {
    return cache.rows;
  }
  const values = await readRange(`${PLIST_SHEET_NAME}!A:I`);
  cache = {
    rows: (values || []).slice(1).filter((row) => row?.[0]),
    loadedAt: Date.now(),
  };
  return cache.rows;
}

export async function searchPlist({ type, t, w, h, l, other = '' }) {
  const rows = await loadPlistRows();
  const nt = formatThicknessNum(t);
  const nw = normalizeNum(w);
  const nh = normalizeNum(h);
  const nl = normalizeNum(l);

  const seen = new Set();
  const matches = [];

  for (const row of rows) {
    const [code, name, , cpdesc1, wdesc1, hdesc1, ldesc1, tdesc1] = row;
    if (!matchCpdesc1(cpdesc1, type, other)) continue;
    if (!thicknessMatchesPlist(nt, tdesc1)) continue;
    if (!matchRowWidthHeight(row, nw, nh)) continue;
    if (!lengthMatchesPlist(nl, ldesc1)) continue;

    const key = `${code}::${name}`;
    if (seen.has(key)) continue;
    seen.add(key);
    matches.push({
      code: String(code ?? '').trim(),
      name: formatPlistDisplayName(row),
      materialWidth: normalizeNum(row[8]),
    });
  }

  return matches;
}

/** 選んだ種類で見つからなければ他の種類も試す（寸法は同じ） */
export async function searchPlistWithTypeFallback({ type, t, w, h, l, other = '' }) {
  const primary = await searchPlist({ type, t, w, h, l, other });
  if (primary.length) {
    return { matches: primary, resolvedType: type, typeAdjusted: false };
  }

  const altHits = [];
  for (const altType of PRODUCT_TYPES) {
    if (altType === type) continue;
    const matches = await searchPlist({ type: altType, t, w, h, l, other });
    if (matches.length) altHits.push({ type: altType, matches });
  }

  if (altHits.length === 1) {
    return {
      matches: altHits[0].matches,
      resolvedType: altHits[0].type,
      typeAdjusted: true,
    };
  }

  if (altHits.length > 1) {
    const byCode = new Map();
    for (const hit of altHits) {
      for (const m of hit.matches) {
        if (!byCode.has(m.code)) {
          byCode.set(m.code, { ...m, types: new Set([hit.type]) });
        } else {
          byCode.get(m.code).types.add(hit.type);
        }
      }
    }
    if (byCode.size === 1) {
      const only = [...byCode.values()][0];
      const resolvedType = [...only.types][0];
      return {
        matches: [{ code: only.code, name: only.name, materialWidth: only.materialWidth }],
        resolvedType,
        typeAdjusted: resolvedType !== type,
      };
    }
    const merged = [...byCode.values()].map((m) => ({
      code: m.code,
      name: m.name,
      materialWidth: m.materialWidth,
    }));
    return {
      matches: merged,
      resolvedType: type,
      typeAdjusted: false,
      multiType: altHits.map((h) => h.type),
    };
  }

  return { matches: [], resolvedType: type, typeAdjusted: false };
}

export async function getPlistCoverageStats() {
  const rows = await loadPlistRows();
  const thicknesses = new Set();
  let missingMaterialWidth = 0;
  for (const row of rows) {
    const t = formatThicknessNum(row[7]);
    if (t) thicknesses.add(t);
    if (!normalizeNum(row[8])) missingMaterialWidth += 1;
  }
  return {
    rowCount: rows.length,
    thicknesses: sortThicknessList([...thicknesses]),
    missingMaterialWidth,
  };
}

export async function getPlistLengthHints({ type, t, w, h, other = '' }) {
  const rows = await loadPlistRows();
  const nt = formatThicknessNum(t);
  const nw = normalizeNum(w);
  const nh = normalizeNum(h);
  const lengths = new Set();

  for (const row of rows) {
    const [, , , cpdesc1, , , ldesc1, tdesc1] = row;
    if (!matchCpdesc1(cpdesc1, type, other)) continue;
    if (!thicknessMatchesPlist(nt, tdesc1)) continue;
    if (!matchRowWidthHeight(row, nw, nh)) continue;
    const l = normalizeNum(ldesc1);
    if (l) lengths.add(l);
  }

  return [...lengths].sort((a, b) => parseFloat(a) - parseFloat(b));
}

export function formatThicknessNum(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  if (/[A-Za-z]/.test(text)) return text;
  const num = parseFloat(text);
  return Number.isFinite(num) ? num.toFixed(1) : text;
}

/** plist 照合のみ 0.8A→0.8、0.4D→0.4、1.0A→1.0 */
export function thicknessForProductLookup(value) {
  const formatted = formatThicknessNum(value);
  if (!formatted) return '';
  const upper = formatted.toUpperCase();
  if (upper === '0.8A') return '0.8';
  if (upper === '0.4D') return '0.4';
  if (upper === '1.0A') return '1.0';
  return formatted;
}

export const NOT_FOUND_PRODUCT_CODE = '暫時未搵到產品編碼';

export function buildProvisionalProductName(typeLabel, spec = {}) {
  const label = String(typeLabel ?? '').trim() || '其他';
  const t = formatThicknessNum(spec.thickness) || String(spec.thickness ?? '').trim();
  const w = String(spec.width ?? '').trim();
  const h = String(spec.height ?? '').trim();
  const l = String(spec.length ?? '').trim();
  return `${t}x${w}x${h} ${label} ${l}mm`;
}

export function applyRecordedThicknessToProductName(name, recordedThickness) {
  const displayT = formatThicknessNum(recordedThickness);
  const lookupT = thicknessForProductLookup(recordedThickness);
  if (!name || !displayT || displayT === lookupT) return name;
  const escaped = lookupT.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return String(name).replace(new RegExp(`^${escaped}x`, 'i'), `${displayT}x`);
}

export const MANUAL_THICKNESS_OPTIONS = ['0.4D', '0.8A', '1.0A'];

export function sortThicknessList(list) {
  return [...list].sort((a, b) => {
    const na = parseFloat(a);
    const nb = parseFloat(b);
    if (na !== nb) return na - nb;
    return String(a).localeCompare(String(b));
  });
}

export async function getThicknessList() {
  const rows = await loadPlistRows();
  const set = new Set();
  for (const row of rows) {
    const t = formatThicknessNum(row[7]);
    if (t) set.add(t);
  }
  for (const t of MANUAL_THICKNESS_OPTIONS) {
    const formatted = formatThicknessNum(t);
    if (formatted) set.add(formatted);
  }
  return sortThicknessList(set);
}
