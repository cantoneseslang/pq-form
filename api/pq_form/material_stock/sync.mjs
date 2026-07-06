import { syncMaterialStockLotsToSheet } from '../../../lib/materialStockSync.js';

export const config = { runtime: 'nodejs', maxDuration: 300 };

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');

  if (req.method !== 'GET' && req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const isCron = req.headers['x-vercel-cron'] === '1';
    const apply = isCron || url.searchParams.get('apply') === '1' || url.searchParams.get('action') === 'sync';

    const result = await syncMaterialStockLotsToSheet({ dryRun: !apply });
    return res.status(200).json({
      success: true,
      action: apply ? 'sync' : 'preview',
      cron: isCron,
      result,
    });
  } catch (e) {
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
