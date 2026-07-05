import { fetchMaterialStockMap } from './rawMaterialInventorySheets.js';

const CACHE_TTL_MS = 10 * 60 * 1000;
let cachedStock = null;
let cachedAt = 0;
let pendingFetch = null;

export async function getCachedMaterialStock() {
  const now = Date.now();
  if (cachedStock && now - cachedAt < CACHE_TTL_MS) {
    return cachedStock;
  }
  if (pendingFetch) return pendingFetch;

  pendingFetch = fetchMaterialStockMap()
    .then((data) => {
      cachedStock = data;
      cachedAt = Date.now();
      pendingFetch = null;
      return cachedStock;
    })
    .catch((error) => {
      pendingFetch = null;
      throw error;
    });

  return pendingFetch;
}

export function clearMaterialStockCache() {
  cachedStock = null;
  cachedAt = 0;
  pendingFetch = null;
}
