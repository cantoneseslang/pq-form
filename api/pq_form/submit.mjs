import {
  getSheetsClient,
  readRange,
  writeRowA1,
  findFirstEmptyRowInBlock,
  dataBlockRange,
  copyRowTemplate,
  applyDataRowLayout,
  normalizeRowValues,
  DATA_START_ROW,
  DATA_END_ROW,
  STYLED_TEMPLATE_END_ROW,
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

    if (targetRow) {
      if (targetRow < DATA_START_ROW || targetRow > DATA_END_ROW) {
        res.setHeader('Cache-Control', 'no-store');
        return res.status(400).json({ success: false, error: `targetRow must be between ${DATA_START_ROW} and ${DATA_END_ROW}` });
      }
    } else {
      const blockRange = dataBlockRange(sheetName);
      const block = await readRange(blockRange);
      targetRow = findFirstEmptyRowInBlock(block);
      if (!targetRow) {
        res.setHeader('Cache-Control', 'no-store');
        return res.status(409).json({ success: false, error: `表の範囲(${DATA_START_ROW}-${DATA_END_ROW})が満杯です` });
      }

      if (targetRow > STYLED_TEMPLATE_END_ROW) {
        await copyRowTemplate(sheetName, targetRow, DATA_START_ROW);
      }
      await applyDataRowLayout(sheetName, targetRow);
    }

    const values = normalizeRowValues(rows[0]);
    const writeRange = `${sheetName}!A${targetRow}:T${targetRow}`;
    await writeRowA1(writeRange, values);

    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ success: true, row: targetRow, range: writeRange });
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}


