import { getSheetsClient, readRange, writeRowA1, findFirstEmptyRowInBlock } from '../../../lib/sheets.js';

export const config = { runtime: 'nodejs' };

function noStoreHeaders() {
  return {
    'Cache-Control': 'no-store',
    'Content-Type': 'application/json; charset=utf-8'
  };
}

export default async function handler(req) {
  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ success: false, error: 'Method not allowed' }), { status: 405, headers: noStoreHeaders() });
  }
  try {
    const body = await req.json();
    const rows = body?.rows || [];
    if (!Array.isArray(rows) || rows.length === 0) {
      return new Response(JSON.stringify({ success: false, error: 'rows is required' }), { status: 400, headers: noStoreHeaders() });
    }

    const { sheetName } = getSheetsClient();
    // Read A8:T16 block
    const blockRange = `${sheetName}!A8:T16`;
    const block = await readRange(blockRange);
    const targetRow = findFirstEmptyRowInBlock(block);
    if (!targetRow) {
      return new Response(JSON.stringify({ success: false, error: '表の範囲(8-16)が満杯です' }), { status: 409, headers: noStoreHeaders() });
    }

    const values = rows[0];
    const writeRange = `${sheetName}!A${targetRow}:T${targetRow}`;
    await writeRowA1(writeRange, values);

    return new Response(JSON.stringify({ success: true, row: targetRow, range: writeRange }), { status: 200, headers: noStoreHeaders() });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: e?.message || String(e) }), { status: 500, headers: noStoreHeaders() });
  }
}


