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
    let hintType = '';
    if (matches.length === 0) {
      const lengths = await getPlistLengthHints({ type, t, w, h, other });
      if (lengths.length) {
        hint = `此規格以下長度有產品編碼：${lengths.join(', ')}mm`;
        hintType = 'length';
      } else {
        hint = '此產品種類+厚度+闊度+高度在plist中無資料';
        hintType = 'no_spec';
      }
    }
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('Access-Control-Allow-Origin', '*');
    return res.status(200).json({ success: true, matches, hint, hintType });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
