import { getSheetsClient, writeRanges } from '../../../lib/sheets.js';

export const config = { runtime: 'nodejs' };

function noStoreHeaders() {
  return { 'Cache-Control': 'no-store', 'Content-Type': 'application/json; charset=utf-8' };
}

export default async function handler(req) {
  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ success: false, error: 'Method not allowed' }), { status: 405, headers: noStoreHeaders() });
  }
  try {
    const body = await req.json();
    const { sheetName } = getSheetsClient();

    // 期待payload: { date: {y,m,d}, types: {...}, machines: {...} }
    const y = String(body?.date?.y||'');
    const m = String(body?.date?.m||'');
    const d = String(body?.date?.d||'');

    // 製品種類・機械名は任意（TRUE/FALSE, その他文字列）
    const t = body?.types || {};
    const mach = body?.machines || {};

    // マッピング（ユーザー確定版に合わせる: B2..R2, B3..J3, B4,D4,F4）
    const updates = [];
    // 日付
    updates.push({ range: `${sheetName}!B4`, values: [y] });
    updates.push({ range: `${sheetName}!D4`, values: [m] });
    updates.push({ range: `${sheetName}!F4`, values: [d] });

    // 製品種類
    const typeMap = [
      { key: '企筒', cell: 'B2' }, { key: '地槽', cell: 'D2' }, { key: '鐵角', cell: 'F2' }, { key: '批灰角', cell: 'H2' },
      { key: 'W角', cell: 'J2' }, { key: '闊槽', cell: 'L2' }, { key: 'C槽', cell: 'N2' }, { key: '其他', cell: 'P2' },
    ];
    typeMap.forEach(({key, cell}) => updates.push({ range: `${sheetName}!${cell}`, values: [t[key] ? 'TRUE' : 'FALSE'] }));
    if (typeof t['其他入力'] === 'string') {
      updates.push({ range: `${sheetName}!R2`, values: [t['其他入力']] });
    }

    // 機械名
    const machMap = [
      { key: '1號滾筒成形機', cell: 'B3' }, { key: '2號滾筒成形機', cell: 'D3' }, { key: '3號滾筒成形機', cell: 'F3' },
      { key: '4號滾筒成形機', cell: 'H3' }, { key: '5號滾筒成形機', cell: 'J3' },
    ];
    machMap.forEach(({key, cell}) => updates.push({ range: `${sheetName}!${cell}`, values: [mach[key] ? 'TRUE' : 'FALSE'] }));

    await writeRanges(updates);
    return new Response(JSON.stringify({ success: true }), { headers: noStoreHeaders() });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: e?.message || String(e) }), { status: 500, headers: noStoreHeaders() });
  }
}


