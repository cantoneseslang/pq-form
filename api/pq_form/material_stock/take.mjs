import { loadMaterialStockTakePayload } from '../../../lib/materialStockTake.js';

export const config = { runtime: 'nodejs' };

const TAKE_CACHE_TTL_MS = 5 * 60 * 1000;
let cachedTake = null;
let cachedTakeAt = 0;
let pendingTake = null;

async function getTakePayload() {
  const now = Date.now();
  if (cachedTake && now - cachedTakeAt < TAKE_CACHE_TTL_MS) {
    return cachedTake;
  }
  if (pendingTake) return pendingTake;

  pendingTake = loadMaterialStockTakePayload()
    .then((payload) => {
      cachedTake = payload;
      cachedTakeAt = Date.now();
      pendingTake = null;
      return cachedTake;
    })
    .catch((error) => {
      pendingTake = null;
      throw error;
    });

  return pendingTake;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'public, max-age=60, s-maxage=300, stale-while-revalidate=600');

  if (req.method !== 'GET') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const payload = await getTakePayload();
    return res.status(200).json({
      success: true,
      source: payload.source,
      fetchedAt: payload.fetchedAt,
      sourceSpreadsheetId: payload.sourceSpreadsheetId,
      stats: payload.stats,
      ageStats: payload.ageStats,
      lotCount: payload.lots.length,
      rowsHtml: payload.rowsHtml,
    });
  } catch (e) {
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
