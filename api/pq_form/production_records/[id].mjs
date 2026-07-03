import { getSupabaseAdmin, isSupabaseConfigured } from '../../../lib/supabase.js';
import {
  dbRowToClient,
  writeMainLinesToSheet,
  normalizeRecordDate,
  packMainData,
  packMaterialData,
  unpackMainData,
  unpackMaterialData,
} from '../../../lib/productionRecords.js';

export const config = { runtime: 'nodejs' };

function parseJsonBody(req) {
  return new Promise((resolve, reject) => {
    if (req.body && typeof req.body === 'object') {
      resolve(req.body);
      return;
    }
    let data = '';
    req.on('data', (chunk) => { data += chunk; });
    req.on('end', () => {
      try { resolve(JSON.parse(data || '{}')); } catch (e) { reject(e); }
    });
    req.on('error', reject);
  });
}

function extractId(req) {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const parts = url.pathname.split('/').filter(Boolean);
  return parts[parts.length - 1];
}

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('Access-Control-Allow-Origin', '*');

  if (req.method !== 'PATCH' && req.method !== 'DELETE') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  if (!isSupabaseConfigured()) {
    return res.status(503).json({ success: false, error: 'Supabase is not configured' });
  }

  const id = extractId(req);
  if (!id || id === 'production_records') {
    return res.status(400).json({ success: false, error: 'Record id is required' });
  }

  try {
    const supabase = getSupabaseAdmin();
    const { data: existing, error: fetchError } = await supabase
      .from('pq_production_records')
      .select('*')
      .eq('id', id)
      .single();

    if (fetchError) throw fetchError;
    if (existing.deleted_at) {
      return res.status(410).json({ success: false, error: 'Record already deleted' });
    }

    if (req.method === 'DELETE') {
      const now = new Date().toISOString();
      const { data, error } = await supabase
        .from('pq_production_records')
        .update({
          deleted_at: now,
          updated_at: now,
        })
        .eq('id', id)
        .select('*')
        .single();

      if (error) throw error;
      return res.status(200).json({ success: true, record: dbRowToClient(data) });
    }

    const body = await parseJsonBody(req);
    const {
      main,
      material,
      mainLines,
      materialLines,
      sheetRows,
      correction_note: correctionNote,
      recordDate,
    } = body;

    if (!main) {
      return res.status(400).json({ success: false, error: 'main is required' });
    }
    if (!String(correctionNote || '').trim()) {
      return res.status(400).json({ success: false, error: 'correction_note is required' });
    }

    const resolvedMainLines = Array.isArray(mainLines) && mainLines.length
      ? mainLines
      : unpackMainData(main).mainLines;
    const resolvedMaterialLines = Array.isArray(materialLines) && materialLines.length
      ? materialLines
      : unpackMaterialData(material || {}).materialLines;
    const resolvedSheetRows = Array.isArray(sheetRows) && sheetRows.length
      ? sheetRows
      : unpackMaterialData(material || {}).sheetRows;

    const now = new Date().toISOString();
    const updatePayload = {
      main_data: packMainData(main, resolvedMainLines),
      material_data: packMaterialData(material || {}, resolvedMaterialLines, resolvedSheetRows),
      correction_note: String(correctionNote).trim(),
      updated_at: now,
      corrected_at: now,
      sheet_row: resolvedSheetRows[0] || existing.sheet_row || null,
    };
    const normalizedDate = normalizeRecordDate(recordDate);
    if (normalizedDate) {
      updatePayload.record_date = normalizedDate;
    }

    const { data, error } = await supabase
      .from('pq_production_records')
      .update(updatePayload)
      .eq('id', id)
      .select('*')
      .single();

    if (error) throw error;

    const rowsToWrite = resolvedSheetRows.length
      ? resolvedSheetRows
      : (existing.sheet_row ? [existing.sheet_row] : []);
    if (rowsToWrite.length && resolvedMainLines.length) {
      await writeMainLinesToSheet(resolvedMainLines, rowsToWrite, existing.sheet_name);
    }

    return res.status(200).json({ success: true, record: dbRowToClient(data) });
  } catch (e) {
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
