import { getSupabaseAdmin, isSupabaseConfigured } from '../../../lib/supabase.js';
import {
  dbRowToClient,
  writeMainLinesToSheet,
  normalizeRecordDate,
  packMainData,
  packMaterialData,
  unpackMainData,
  unpackMaterialData,
  getMachineFromMachines,
  parseRecordDateForDaily,
} from '../../../lib/productionRecords.js';
import {
  updateDailyReportEntry,
  deleteDailyReportEntry,
  dailySheetTabName,
} from '../../../lib/dailyReport.js';

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

async function syncDailyReportForRecord(existing, mainLines, materialLines, dailyReportOverride = null) {
  const dailyReport = dailyReportOverride || unpackMaterialData(existing.material_data || {}).dailyReport;
  const recordDate = existing.record_date;
  const date = parseRecordDateForDaily(recordDate);
  if (!date) return { skipped: true, reason: 'invalid date' };

  const pageType = existing.page_type === 'auto' ? 'auto' : 'molding';
  const machine = dailyReport?.machine || getMachineFromMachines(existing.machines || {});
  if (pageType === 'molding' && !machine) {
    return { skipped: true, reason: 'machine not set' };
  }

  const tabFromDate = dailySheetTabName(date);
  const storedRow = dailyReport?.tabName === tabFromDate ? dailyReport?.row : null;

  return updateDailyReportEntry({
    date,
    pageType,
    machine,
    productTypes: existing.product_types || {},
    mainLines,
    materialLines,
    targetRow: storedRow,
    tabName: tabFromDate,
  });
}

async function clearDailyReportForRecord(existing) {
  const { dailyReport } = unpackMaterialData(existing.material_data || {});
  const { mainLines, materialLines } = unpackMainData(existing.main_data || {});
  const date = parseRecordDateForDaily(existing.record_date);
  if (!date) return { skipped: true, reason: 'invalid date' };

  const pageType = existing.page_type === 'auto' ? 'auto' : 'molding';
  const machine = dailyReport?.machine || getMachineFromMachines(existing.machines || {});
  const tabFromDate = dailySheetTabName(date);
  const storedRow = dailyReport?.tabName === tabFromDate ? dailyReport?.row : null;

  return deleteDailyReportEntry({
    date,
    pageType,
    machine,
    productTypes: existing.product_types || {},
    mainLines,
    materialLines,
    targetRow: storedRow,
    tabName: tabFromDate,
  });
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
      let dailyReportResult = null;
      try {
        dailyReportResult = await clearDailyReportForRecord(existing);
      } catch (dailyErr) {
        dailyReportResult = { error: dailyErr?.message || String(dailyErr) };
      }

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
      return res.status(200).json({
        success: true,
        record: dbRowToClient(data),
        dailyReport: dailyReportResult,
      });
    }

    const body = await parseJsonBody(req);

    if (body?.linkDailyReportOnly && body?.dailyReportLink) {
      const existingMaterial = unpackMaterialData(existing.material_data || {});
      const now = new Date().toISOString();
      const updatePayload = {
        material_data: packMaterialData(
          existingMaterial.material,
          existingMaterial.materialLines,
          existingMaterial.sheetRows,
          body.dailyReportLink,
        ),
        updated_at: now,
      };

      const { data, error } = await supabase
        .from('pq_production_records')
        .update(updatePayload)
        .eq('id', id)
        .select('*')
        .single();

      if (error) throw error;
      return res.status(200).json({ success: true, record: dbRowToClient(data) });
    }

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
    const existingDailyReport = unpackMaterialData(existing.material_data || {}).dailyReport;

    const now = new Date().toISOString();
    const updatePayload = {
      main_data: packMainData(main, resolvedMainLines),
      material_data: packMaterialData(
        material || {},
        resolvedMaterialLines,
        resolvedSheetRows,
        existingDailyReport,
      ),
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

    let dailyReportResult = null;
    try {
      dailyReportResult = await syncDailyReportForRecord(
        data,
        resolvedMainLines,
        resolvedMaterialLines,
      );
    } catch (dailyErr) {
      dailyReportResult = { error: dailyErr?.message || String(dailyErr) };
    }

    return res.status(200).json({
      success: true,
      record: dbRowToClient(data),
      dailyReport: dailyReportResult,
    });
  } catch (e) {
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
