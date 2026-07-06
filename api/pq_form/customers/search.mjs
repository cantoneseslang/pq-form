import {
  isSalesSupabaseConfigured,
  searchCustomersByCode,
  searchCustomersByCnName,
} from '../../../lib/customerSearch.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    if (!isSalesSupabaseConfigured()) {
      res.setHeader('Cache-Control', 'no-store');
      return res.status(503).json({
        success: false,
        error: 'Sales customer database is not configured',
      });
    }

    const url = new URL(req.url, `http://${req.headers.host}`);
    const code = url.searchParams.get('code') || '';
    const name = url.searchParams.get('name') || '';

    if (!code && !name) {
      res.setHeader('Cache-Control', 'no-store');
      return res.status(400).json({
        success: false,
        error: 'code or name is required',
      });
    }

    const matches = code
      ? await searchCustomersByCode(code)
      : await searchCustomersByCnName(name);

    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('Access-Control-Allow-Origin', '*');
    return res.status(200).json({ success: true, matches });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
