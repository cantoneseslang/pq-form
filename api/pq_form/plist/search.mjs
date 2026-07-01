import { searchPlist, getPlistLengthHints } from '../../../lib/plist.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const type = url.searchParams.get('type') || '';
    const t = url.searchParams.get('t') || '';
    const w = url.searchParams.get('w') || '';
    const h = url.searchParams.get('h') || '';
    const l = url.searchParams.get('l') || '';
    const other = url.searchParams.get('other') || '';

    if (!type || !t || !w || !h || !l) {
      res.setHeader('Cache-Control', 'no-store');
      return res.status(400).json({
        success: false,
        error: 'type, t, w, h, l are required',
      });
    }

    const matches = await searchPlist({ type, t, w, h, l, other });
    let hint = '';
    if (matches.length === 0) {
      const lengths = await getPlistLengthHints({ type, t, w, h, other });
      hint = lengths.length
        ? `長度請改用: ${lengths.join(', ')}mm`
        : '此產品種類+厚度+闊度+高度在plist中無資料';
    }
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('Access-Control-Allow-Origin', '*');
    return res.status(200).json({ success: true, matches, hint });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
