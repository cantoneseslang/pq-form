import { getSheetsClient, readRange } from '../../../lib/sheets.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const dateParam = url.searchParams.get('date') || '';
    const { sheetName } = getSheetsClient();
    const range = `${sheetName}!A8:T16`;
    const values = await readRange(range);
    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ success: true, date: dateParam, rows: values });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}


