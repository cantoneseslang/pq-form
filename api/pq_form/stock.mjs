export const config = { runtime: 'nodejs' };

const DEFAULT_INVENTORY_API_URL = 'https://qr-new-six.vercel.app/api/inventory';
const CACHE_TTL_MS = 5 * 60 * 1000;

let cachedData = null;
let cachedAt = 0;

function normalizeInventoryItem(item) {
  if (!item || !item.code) return null;
  return {
    code: String(item.code).trim(),
    name: String(item.name || item.category_detail || '').trim(),
    location: String(item.location || '').trim(),
    onHand: item.on_hand ?? null,
    withoutDn: item.without_dn ?? null,
    available: item.quantity ?? null,
    unit: String(item.unit || '').trim(),
    category: String(item.category || '').trim(),
    updated: String(item.updated || '').trim(),
    // Legacy fields used by main PQ-Form stock lookup
    thickness2: '',
    width2: '',
    height: '',
    length: '',
  };
}

function inventoryToCodeMap(raw) {
  const data = {};
  if (!raw || typeof raw !== 'object') return data;
  Object.values(raw).forEach((item) => {
    const normalized = normalizeInventoryItem(item);
    if (!normalized) return;
    data[normalized.code.toUpperCase()] = normalized;
  });
  return data;
}

async function fetchInventoryMap() {
  const now = Date.now();
  if (cachedData && now - cachedAt < CACHE_TTL_MS) {
    return cachedData;
  }

  const url = process.env.INVENTORY_API_URL || DEFAULT_INVENTORY_API_URL;
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`Inventory API failed (${res.status})`);
  }

  const raw = await res.json();
  cachedData = inventoryToCodeMap(raw);
  cachedAt = now;
  return cachedData;
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('Access-Control-Allow-Origin', '*');

  if (req.method !== 'GET') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const data = await fetchInventoryMap();
    return res.status(200).json({ success: true, data });
  } catch (e) {
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
