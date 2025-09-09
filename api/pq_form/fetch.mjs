import { getSheetsClient, readRange } from '../../../lib/sheets.js';

export const config = { runtime: 'nodejs' };

function noStoreHeaders() {
  return { 'Cache-Control': 'no-store', 'Content-Type': 'application/json; charset=utf-8' };
}

export default async function handler(req) {
  try {
    const url = new URL(req.url);
    const dateParam = url.searchParams.get('date') || '';
    const { sheetName } = getSheetsClient();
    // ひとまず日付によらずA8:T16を返却（クライアントでフィルタ可）。必要なら列A,Bでマッチングを追加
    const range = `${sheetName}!A8:T16`;
    const values = await readRange(range);
    return new Response(JSON.stringify({ success: true, date: dateParam, rows: values }), { headers: noStoreHeaders() });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: e?.message || String(e) }), { status: 500, headers: noStoreHeaders() });
  }
}


