import { analyzeMaterialWidthSheets, syncMaterialWidthToPlist, addEofficeOnlyItemsToPlist } from '../../../lib/materialWidthSheets.js';
import { applyPlistDedupe } from '../../../lib/plistDuplicates.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'POST' && req.method !== 'GET') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const action = url.searchParams.get('action') || 'analyze';
    const apply = url.searchParams.get('apply') === '1' || action === 'sync';
    const addEoffice = action === 'add-eoffice' || url.searchParams.get('addEoffice') === '1';
    const dedupe = action === 'dedupe' || url.searchParams.get('dedupe') === '1';

    if (dedupe) {
      const result = await applyPlistDedupe({ dryRun: !apply });
      res.setHeader('Cache-Control', 'no-store');
      return res.status(200).json({ success: true, action: 'dedupe', result });
    }

    if (addEoffice) {
      const result = await addEofficeOnlyItemsToPlist({ dryRun: !apply });
      res.setHeader('Cache-Control', 'no-store');
      return res.status(200).json({ success: true, action: 'add-eoffice', result });
    }

    if (apply && action === 'sync') {
      const result = await syncMaterialWidthToPlist({ dryRun: false });
      res.setHeader('Cache-Control', 'no-store');
      return res.status(200).json({ success: true, action: 'sync', result });
    }

    const report = await analyzeMaterialWidthSheets();
    const format = url.searchParams.get('format') || '';

    if (format === 'eoffice-csv') {
      const items = report.crossRef?.eofficeOnlyItems || [];
      const lines = ['code,name,width_mm'];
      for (const item of items) {
        const name = String(item.name ?? '').replace(/"/g, '""');
        lines.push([item.code, `"${name}"`, item.width].join(','));
      }
      res.setHeader('Cache-Control', 'no-store');
      res.setHeader('Content-Type', 'text/csv; charset=utf-8');
      res.setHeader('Content-Disposition', 'attachment; filename="eoffice-not-in-plist.csv"');
      return res.status(200).send(`${lines.join('\n')}\n`);
    }

    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ success: true, action: 'analyze', report });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
