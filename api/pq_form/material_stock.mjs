import {
  fetchMaterialStockMap,
  lookupMaterialStock,
  analyzeRawMaterialInventory,
} from '../../lib/rawMaterialInventorySheets.js';
import {
  produciblePieces,
  pieceWeightKg,
  materialAvailabilityStatus,
  densityForMaterial,
} from '../../lib/materialYield.js';

export const config = { runtime: 'nodejs' };

const CACHE_TTL_MS = 5 * 60 * 1000;
let cachedStock = null;
let cachedAt = 0;

async function getStockData() {
  const now = Date.now();
  if (cachedStock && now - cachedAt < CACHE_TTL_MS) {
    return cachedStock;
  }
  cachedStock = await fetchMaterialStockMap();
  cachedAt = now;
  return cachedStock;
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('Access-Control-Allow-Origin', '*');

  if (req.method !== 'GET') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const action = url.searchParams.get('action') || 'lookup';

    if (action === 'analyze') {
      const report = await analyzeRawMaterialInventory();
      return res.status(200).json({ success: true, action: 'analyze', report });
    }

    const stock = await getStockData();
    const thickness = url.searchParams.get('thickness') || '';
    const materialWidth = url.searchParams.get('materialWidth') || url.searchParams.get('mw') || '';
    const lengthMm = url.searchParams.get('length') || url.searchParams.get('l') || '';
    const requestedQty = url.searchParams.get('qty') || '';

    if (!thickness && !materialWidth) {
      return res.status(200).json({
        success: true,
        fetchedAt: stock.fetchedAt,
        stats: stock.stats,
        summaries: stock.summaries,
        byTab: stock.byTab,
      });
    }

    const hit = lookupMaterialStock({
      byTab: stock.byTab,
      thickness,
      materialWidth,
    });

    if (!hit) {
      return res.status(200).json({
        success: true,
        found: false,
        thickness,
        materialWidth,
        fetchedAt: stock.fetchedAt,
      });
    }

    const densityGcm3 = densityForMaterial({ tabName: hit.tabTitles[0] });
    const pieceKg = pieceWeightKg({
      lengthMm,
      materialWidthMm: hit.materialWidth,
      thicknessMm: thickness,
      densityGcm3,
    });
    const producibleQty = produciblePieces({
      availableKg: hit.totalKg,
      lengthMm,
      materialWidthMm: hit.materialWidth,
      thicknessMm: thickness,
      densityGcm3,
    });

    return res.status(200).json({
      success: true,
      found: true,
      fetchedAt: stock.fetchedAt,
      stock: hit,
      yield: {
        pieceWeightKg: pieceKg,
        producibleQty,
        requestedQty: requestedQty ? Number(requestedQty) : null,
        status: materialAvailabilityStatus({
          requestedQty,
          producibleQty,
        }),
        densityGcm3,
      },
    });
  } catch (e) {
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
