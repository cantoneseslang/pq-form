import {
  findLatestGoodPlistRevision,
  listSpreadsheetRevisions,
  previewPlistRevision,
  restorePlistFromRevision,
} from '../../../lib/plistRestore.js';

export const config = { runtime: 'nodejs', maxDuration: 300 };

export default async function handler(req, res) {
  if (req.method !== 'GET' && req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const action = url.searchParams.get('action') || 'preview';
    const apply = url.searchParams.get('apply') === '1' || action === 'restore';
    const revisionId = url.searchParams.get('revisionId') || '';

    if (action === 'revisions') {
      const result = await listSpreadsheetRevisions();
      res.setHeader('Cache-Control', 'no-store');
      return res.status(200).json({ success: true, action, result });
    }

    if (action === 'find') {
      const result = await findLatestGoodPlistRevision();
      res.setHeader('Cache-Control', 'no-store');
      return res.status(200).json({ success: true, action, result });
    }

    if (action === 'preview-revision' && revisionId) {
      const result = await previewPlistRevision(revisionId);
      res.setHeader('Cache-Control', 'no-store');
      return res.status(200).json({ success: true, action, result });
    }

    const result = await restorePlistFromRevision({
      revisionId: revisionId || undefined,
      dryRun: !apply,
    });

    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({
      success: !!result.restored || result.dryRun,
      action: apply ? 'restore' : 'preview-restore',
      result,
    });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
