const DEFAULT_STEEL_DENSITY_G_CM3 = 7.85;
const ALUMINIUM_DENSITY_G_CM3 = 2.7;

const DENSITY_BY_TAB_HINT = [
  { pattern: /aluminium/i, density: ALUMINIUM_DENSITY_G_CM3 },
];

export function densityForMaterial({ tabName = '', thicknessKey = '' } = {}) {
  const text = `${tabName} ${thicknessKey}`;
  for (const rule of DENSITY_BY_TAB_HINT) {
    if (rule.pattern.test(text)) return rule.density;
  }
  return DEFAULT_STEEL_DENSITY_G_CM3;
}

export function parsePositiveNumber(value) {
  const text = String(value ?? '').replace(/,/g, '').trim();
  if (!text) return null;
  const num = parseFloat(text);
  return Number.isFinite(num) && num > 0 ? num : null;
}

/** Weight in kg for one piece: L(mm)×W(mm)×T(mm)×density(g/cm³) / 1,000,000 */
export function pieceWeightKg({ lengthMm, materialWidthMm, thicknessMm, densityGcm3 = DEFAULT_STEEL_DENSITY_G_CM3 }) {
  const length = parsePositiveNumber(lengthMm);
  const width = parsePositiveNumber(materialWidthMm);
  const thickness = parsePositiveNumber(thicknessMm);
  if (!length || !width || !thickness) return null;
  return (length * width * thickness * densityGcm3) / 1_000_000;
}

export function produciblePieces({
  availableKg,
  lengthMm,
  materialWidthMm,
  thicknessMm,
  densityGcm3 = DEFAULT_STEEL_DENSITY_G_CM3,
}) {
  const kg = parsePositiveNumber(availableKg);
  const pieceKg = pieceWeightKg({ lengthMm, materialWidthMm, thicknessMm, densityGcm3 });
  if (!kg || !pieceKg) return null;
  return Math.floor(kg / pieceKg);
}

export function materialAvailabilityStatus({ requestedQty, producibleQty }) {
  const req = parsePositiveNumber(requestedQty);
  const prod = parsePositiveNumber(producibleQty);
  if (!req || prod === null) return 'unknown';
  if (prod >= req) return 'ok';
  return 'short';
}
