import { getSheetsClient, writeRanges } from '../../../lib/sheets.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }
  try {
    const body = req.body ?? (await new Promise((resolve, reject) => {
      let data = '';
      req.on('data', chunk => { data += chunk; });
      req.on('end', () => { try { resolve(JSON.parse(data || '{}')); } catch (e) { reject(e); } });
      req.on('error', reject);
    }));

    const { sheetName } = getSheetsClient();

    const y = String(body?.date?.y||'');
    const m = String(body?.date?.m||'');
    const d = String(body?.date?.d||'');

    const t = body?.types || {};
    const mach = body?.machines || {};

    const updates = [];
    updates.push({ range: `${sheetName}!B4`, values: [y] });
    updates.push({ range: `${sheetName}!D4`, values: [m] });
    updates.push({ range: `${sheetName}!F4`, values: [d] });

    const typeMap = [
      { key: '企筒', cell: 'B2' }, { key: '地槽', cell: 'D2' }, { key: '鐵角', cell: 'F2' }, { key: '批灰角', cell: 'H2' },
      { key: 'W角', cell: 'J2' }, { key: '闊槽', cell: 'L2' }, { key: 'C槽', cell: 'N2' }, { key: '其他', cell: 'P2' },
    ];
    typeMap.forEach(({key, cell}) => updates.push({ range: `${sheetName}!${cell}`, values: [t[key] ? 'TRUE' : 'FALSE'] }));
    if (typeof t['其他入力'] === 'string') {
      updates.push({ range: `${sheetName}!R2`, values: [t['其他入力']] });
    }

    const machMap = [
      { key: '1號滾筒成形機', cell: 'B3' }, { key: '2號滾筒成形機', cell: 'D3' }, { key: '3號滾筒成形機', cell: 'F3' },
      { key: '4號滾筒成形機', cell: 'H3' }, { key: '5號滾筒成形機', cell: 'J3' },
    ];
    machMap.forEach(({key, cell}) => updates.push({ range: `${sheetName}!${cell}`, values: [mach[key] ? 'TRUE' : 'FALSE'] }));

    await writeRanges(updates);
    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ success: true });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}


