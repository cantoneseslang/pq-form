import {
  getSheetsClient,
  readRange,
  writeRowA1,
  findFirstEmptyRowInBlock,
  dataBlockRange,
  prepareNewDataRow,
  ensureSheetRowCapacity,
  normalizeRowValues,
  DATA_START_ROW,
} from '../../lib/sheets.js';

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

    const rows = body?.rows || [];
    if (!Array.isArray(rows) || rows.length === 0) {
      res.setHeader('Cache-Control', 'no-store');
      return res.status(400).json({ success: false, error: 'rows is required' });
    }

    const { sheetName } = getSheetsClient();
    const requestedRow = parseInt(body?.targetRow, 10);
    let targetRow = Number.isFinite(requestedRow) ? requestedRow : null;
    const isOverwrite = Boolean(targetRow);

    if (targetRow && targetRow < DATA_START_ROW) {
      res.setHeader('Cache-Control', 'no-store');
      return res.status(400).json({ success: false, error: `targetRow must be >= ${DATA_START_ROW}` });
    }

    if (!targetRow) {
      const blockRange = dataBlockRange(sheetName);
      const block = await readRange(blockRange);
      targetRow = findFirstEmptyRowInBlock(block);
      await prepareNewDataRow(sheetName, targetRow);
    } else {
      await ensureSheetRowCapacity(sheetName, targetRow);
    }

    const values = normalizeRowValues(rows[0]);
    const writeRange = `${sheetName}!A${targetRow}:T${targetRow}`;
    await writeRowA1(writeRange, values);

    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ success: true, row: targetRow, range: writeRange, overwrite: isOverwrite });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
