import { submitMaterialQcRecord } from '../../../lib/materialQcSheets.js';

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

function validateSubmitBody(body) {
  const errors = [];
  if (!String(body?.lotNo ?? '').trim()) errors.push('lotNo is required');
  if (!String(body?.inspector ?? '').trim()) errors.push('inspector is required');
  const thickness = parseFloat(body?.thicknessUm);
  if (!Number.isFinite(thickness)) errors.push('thicknessUm is required');
  if (!String(body?.substrate ?? '').trim()) errors.push('substrate is required');
  return errors;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const body = await readJsonBody(req);
    const validationErrors = validateSubmitBody(body);
    if (validationErrors.length) {
      res.setHeader('Cache-Control', 'no-store');
      return res.status(400).json({ success: false, error: validationErrors.join('; '), validationErrors });
    }

    const result = await submitMaterialQcRecord({
      receiptDate: String(body.receiptDate ?? '').trim(),
      lotNo: String(body.lotNo ?? '').trim(),
      materialSpec: String(body.materialSpec ?? '').trim(),
      supplier: String(body.supplier ?? '').trim(),
      inspector: String(body.inspector ?? '').trim(),
      substrate: String(body.substrate ?? '').trim(),
      thicknessUm: parseFloat(body.thicknessUm),
      imageBase64: body.imageBase64 || '',
      mimeType: body.mimeType || 'image/jpeg',
    });

    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ success: true, ...result });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
