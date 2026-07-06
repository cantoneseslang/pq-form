import { writeProductionOrder } from '../../../lib/productionOrderSheets.js';

export const config = { runtime: 'nodejs' };

async function readJsonBody(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', (chunk) => { data += chunk; });
    req.on('end', () => {
      try { resolve(JSON.parse(data || '{}')); } catch (e) { reject(e); }
    });
    req.on('error', reject);
  });
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const body = await readJsonBody(req);
    const result = await writeProductionOrder(body);

    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ success: true, ...result });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    if (e?.validationErrors) {
      return res.status(400).json({
        success: false,
        error: e.message,
        validationErrors: e.validationErrors,
      });
    }
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
