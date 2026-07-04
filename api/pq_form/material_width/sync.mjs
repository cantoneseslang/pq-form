import { analyzeMaterialWidthSheets, syncMaterialWidthToPlist } from '../../../lib/materialWidthSheets.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'POST' && req.method !== 'GET') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const action = url.searchParams.get('action') || 'analyze';
    const apply = url.searchParams.get('apply') === '1' || action === 'sync';

    if (apply) {
      const result = await syncMaterialWidthToPlist({ dryRun: false });
      res.setHeader('Cache-Control', 'no-store');
      return res.status(200).json({ success: true, action: 'sync', result });
    }

    const report = await analyzeMaterialWidthSheets();
    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ success: true, action: 'analyze', report });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
