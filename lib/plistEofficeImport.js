import { englishTypeToProductTypeKey } from './dailyReportImport.js';
import { formatThicknessKey, normalizeMaterialWidth } from './materialWidthSheets.js';

function parseCsvRow(line) {
  const quoted = line.match(/^([^,]+),"([^"]*)",([^,]+)$/);
  if (quoted) {
    return { code: quoted[1].trim(), name: quoted[2].trim(), width: quoted[3].trim() };
  }
  const parts = line.split(',');
  return {
    code: parts[0]?.trim() || '',
    name: parts.slice(1, -1).join(',').replace(/^"|"$/g, '').trim(),
    width: parts[parts.length - 1]?.trim() || '',
  };
}

function feetToMm(feet) {
  const n = parseFloat(feet);
  if (!Number.isFinite(n)) return '';
  if (Math.abs(n - 8) < 0.01) return '2440';
  if (Math.abs(n - 9) < 0.01) return '2745';
  if (Math.abs(n - 10) < 0.01) return '3000';
  return String(Math.round(n * 304.8));
}

export function parseLengthFromText(name, code) {
  const text = String(name ?? '');
  let m = text.match(/L[-=]\s*(\d+(?:\.\d+)?)\s*m(?:m)?/i);
  if (m) {
    const v = parseFloat(m[1]);
    return v < 20 ? String(Math.round(v * 1000)) : m[1].replace(/\.0$/, '');
  }
  m = text.match(/L[-=]\s*I?(\d+(?:\.\d+)?)\s*['′]/i);
  if (m) return feetToMm(m[1]);
  m = text.match(/L-(\d+(?:[,\.]\d+)?)\s*m\b/i);
  if (m) return m[1].replace(',', '');
  m = String(code ?? '').match(/MM(\d+)/i);
  if (m) return m[1];
  m = text.match(/(\d{3,5})\s*mm/i);
  return m ? m[1] : '';
}

export function parseDimsFromText(name) {
  const text = String(name ?? '');
  let m = text.match(/\((\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)\s*mm\)/i);
  if (m) return { w: m[1], h: m[2], t: formatThicknessKey(m[3]) };
  m = text.match(/(\d+(?:\.\d+)?)\s*mm\s*x\s*(\d+(?:\.\d+)?)\s*mm\s*x\s*(\d+(?:\.\d+)?)\s*mm(?:\s*x\s*([\d,\.]+)\s*mm)?/i);
  if (m) {
    return {
      w: m[1],
      h: m[2],
      t: formatThicknessKey(m[3]),
      l: m[4] ? m[4].replace(',', '') : '',
    };
  }
  m = text.match(/(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)\s*mm/i);
  if (m) return { w: m[1], h: m[2], t: formatThicknessKey(m[3]) };
  m = text.match(/\((\d+(?:\.\d+)?)\s*mm\s*thk\.?\)/i);
  if (m) return { t: formatThicknessKey(m[1]) };
  return null;
}

export function parseDimsFromCode(code) {
  const c = String(code ?? '').trim().toUpperCase();
  let m = c.match(/^(US|UT)(\d{3})(\d{2})(\d{2,3})MI(\d{4})/i);
  if (m) {
    const tRaw = m[4];
    const tVal = tRaw.length === 3 ? parseInt(tRaw, 10) / 100 : parseInt(tRaw, 10) / 10;
    const lenCode = m[5];
    const l = lenCode.startsWith('10') ? '3000' : lenCode.startsWith('08') ? '2440' : '';
    return {
      type: m[1] === 'US' ? '企筒' : '地槽',
      w: String(parseInt(m[2], 10)),
      h: m[3],
      t: formatThicknessKey(tVal),
      l,
    };
  }
  m = c.match(/^(US|UT)(\d{3})(\d{2})(\d{2,3})MI1000/i);
  if (m) {
    const tRaw = m[4];
    const tVal = tRaw.length === 3 ? parseInt(tRaw, 10) / 100 : parseInt(tRaw, 10) / 10;
    return {
      type: m[1] === 'US' ? '企筒' : '地槽',
      w: String(parseInt(m[2], 10)),
      h: m[3],
      t: formatThicknessKey(tVal),
      l: '3000',
    };
  }
  m = c.match(/^(US|UT)(\d{3})(\d{2})(\d{2})MM(\d+)/);
  if (m) {
    return {
      type: m[1] === 'US' ? '企筒' : '地槽',
      w: String(parseInt(m[2], 10)),
      h: m[3],
      t: formatThicknessKey(parseInt(m[4], 10) / 10),
      l: m[5],
    };
  }
  m = c.match(/^V(\d{3})(\d{2})(\d{2})MM(\d+)/);
  if (m) {
    return {
      w: String(parseInt(m[2], 10)),
      h: m[3],
      t: formatThicknessKey(parseInt(m[4], 10) / 10),
      l: m[5],
    };
  }
  m = c.match(/^MA(\d{2})(\d{2})(\d{2})M(\d+)/);
  if (m) {
    return {
      type: '鐵角',
      w: m[1],
      h: m[2],
      t: formatThicknessKey(parseInt(m[3], 10) / 10),
      l: m[4],
    };
  }
  m = c.match(/^GHW(\d{2})M(\d+)/);
  if (m) {
    return {
      type: '闊槽',
      w: '50',
      h: '25',
      t: formatThicknessKey(parseInt(m[1], 10) / 10),
      l: m[2],
    };
  }
  m = c.match(/^TNIW(\d{2})(\d{2})I1000/i);
  if (m) {
    return {
      type: 'W角',
      w: m[1],
      h: m[2],
      t: '0.4',
      l: '3000',
    };
  }
  m = c.match(/^GSW(\d{2})1?1000B/i);
  if (m) {
    return {
      type: 'W角',
      w: '0',
      h: '0',
      t: formatThicknessKey(parseInt(m[1], 10) / 10),
      l: '3000',
    };
  }
  m = c.match(/^GSW(\d{2})M(\d+)/);
  if (m) {
    return {
      type: 'W角',
      w: '0',
      h: '0',
      t: formatThicknessKey(parseInt(m[1], 10) / 10),
      l: m[2],
    };
  }
  m = c.match(/^GSW(\d{2})I(\d{2})(\d{2})/);
  if (m) {
    return {
      type: 'W角',
      w: '0',
      h: '0',
      t: formatThicknessKey(parseInt(m[1], 10) / 10),
      l: feetToMm(parseInt(m[2], 10)),
    };
  }
  m = c.match(/^GSC(\d{2,3})?I(\d{2})(\d{2})/);
  if (m) {
    return {
      type: 'C槽',
      w: '38',
      h: '38',
      t: formatThicknessKey((parseInt(m[1] || '08', 10) || 8) / 10),
      l: feetToMm(parseInt(m[2], 10)),
    };
  }
  m = c.match(/^GSC\d*M(\d+)/);
  if (m) {
    return { type: 'C槽', w: '38', h: '38', t: '0.8', l: m[1] };
  }
  m = c.match(/^GHC(\d{2})M(\d+)/);
  if (m) {
    const cc = parseInt(m[1], 10);
    return {
      type: '闊槽',
      w: String(cc),
      h: String(cc),
      t: formatThicknessKey(cc / 10),
      l: m[2],
    };
  }
  m = c.match(/^VUU(\d{3})(\d{2})(\d{2})(\d{2})MM(\d+)/);
  if (m) {
    return {
      type: '地槽',
      w: String(parseInt(m[2], 10)),
      h: m[4],
      t: formatThicknessKey(parseInt(m[5], 10) / 10),
      l: m[6],
    };
  }
  return null;
}

export function inferProductType(code, name, partial = {}) {
  if (partial.type) return partial.type;
  const fromEnglish = englishTypeToProductTypeKey(name);
  if (fromEnglish) return fromEnglish;
  const c = String(code ?? '').trim().toUpperCase();
  const n = String(name ?? '').toLowerCase();
  if (/^US/.test(c) || (/\bstud\b/.test(n) && !/\brunner\b/.test(n))) {
    return /c-t|ct stud/i.test(name) ? 'CT企筒打孔' : '企筒';
  }
  if (/^UT/.test(c) || /\brunner\b|j-runner|i-runner/.test(n)) return '地槽';
  if (/R$/.test(c) && /^V/.test(c)) return '地槽';
  if (/S$/.test(c) && /^V/.test(c)) return '企筒';
  if (/\bw-bar|main bar/.test(n)) return 'W角';
  if (/main channel|c38/.test(n)) return 'C槽';
  if (/carrying channel|double furring|cw-|u-channel/.test(n)) return '闊槽';
  if (/\bangle\b/.test(n)) return '鐵角';
  if (/deflection/.test(n)) return '地槽';
  return '其他';
}

export function buildPlistDisplayName({ type, thickness, width, height, length, fallbackName }) {
  const t = formatThicknessKey(thickness);
  const w = normalizeMaterialWidth(width);
  const h = normalizeMaterialWidth(height);
  const l = normalizeMaterialWidth(length);
  if (t && w && h && l && type) return `${t}x${w}x${h} ${type} ${l}mm`;
  return String(fallbackName ?? '').trim();
}

export function eofficeItemToPlistRow({ code, name, width }) {
  const fromCode = parseDimsFromCode(code) || {};
  const fromName = parseDimsFromText(name) || {};
  const type = inferProductType(code, name, fromCode);
  const thickness = fromCode.t || fromName.t || '';
  const w = fromCode.w || fromName.w || '';
  const h = fromCode.h || fromName.h || '';
  const length = fromCode.l || fromName.l || parseLengthFromText(name, code) || '';
  const materialWidth = normalizeMaterialWidth(width);
  const displayName = buildPlistDisplayName({
    type,
    thickness,
    width: w,
    height: h,
    length,
    fallbackName: name,
  });

  const complete = !!(type && thickness && w && h && length);
  return {
    code,
    row: [
      code,
      displayName,
      name,
      type,
      w,
      h,
      length,
      thickness,
      materialWidth,
    ],
    meta: {
      type,
      thickness,
      width: w,
      height: h,
      length,
      materialWidth,
      complete,
      englishName: name,
    },
  };
}

export function parseEofficeCsv(text) {
  const lines = String(text ?? '').trim().split('\n').filter(Boolean);
  const header = lines[0]?.toLowerCase().includes('code') ? lines.slice(1) : lines;
  return header.map((line) => parseCsvRow(line)).filter((item) => item.code);
}

export function buildPlistRowsFromEofficeItems(items) {
  return items.map((item) => eofficeItemToPlistRow(item));
}
