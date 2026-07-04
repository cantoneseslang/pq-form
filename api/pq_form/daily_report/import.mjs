import { getSupabaseAdmin, isSupabaseConfigured } from '../../../lib/supabase.js';
import {
  clientToDbInsert,
  dbRowToClient,
  unpackMaterialData,
} from '../../../lib/productionRecords.js';
import {
  buildImportRecordsFromDailyTab,
  collectExistingDailyReportLinks,
  repairImportedProductNamesForTab,
  scanDailyReportTab,
} from '../../../lib/dailyReportImport.js';

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

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('Access-Control-Allow-Origin', '*');

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  if (!isSupabaseConfigured()) {
    return res.status(503).json({ success: false, error: 'Supabase is not configured' });
  }

  try {
    const body = await readJsonBody(req);
    const tabName = String(body?.tabName ?? body?.day ?? '').trim();
    if (!tabName || !/^\d{1,2}$/.test(tabName)) {
      return res.status(400).json({ success: false, error: 'tabName (1-31) is required' });
    }

    const preview = body?.preview === true;
    const repair = body?.repair === true;
    const supabase = getSupabaseAdmin();

    if (repair) {
      const { recordDateIso, repaired, skipped } = await repairImportedProductNamesForTab(tabName, supabase);
      return res.status(200).json({
        success: true,
        repair: true,
        tabName,
        recordDateIso,
        repaired,
        skipped,
      });
    }

    const scanned = await scanDailyReportTab(tabName);
    let existingLinks = new Set();

    if (scanned.recordDateIso) {
      const { data: existingRows, error: fetchError } = await supabase
        .from('pq_production_records')
        .select('material_data')
        .eq('record_date', scanned.recordDateIso)
        .is('deleted_at', null);

      if (fetchError) throw fetchError;

      existingLinks = collectExistingDailyReportLinks(
        (existingRows || []).map((row) => ({
          dailyReport: unpackMaterialData(row.material_data || {}).dailyReport,
        })),
      );
    }

    const { recordDateIso, imported, skipped, errors } = await buildImportRecordsFromDailyTab(
      tabName,
      existingLinks,
    );

    if (preview) {
      return res.status(200).json({
        success: true,
        preview: true,
        tabName,
        recordDateIso,
        wouldImport: imported,
        skipped,
        errors,
      });
    }

    const saved = [];
    for (const record of imported) {
      const insertRow = clientToDbInsert(record);
      insertRow.correction_note = record.correctionNote || '生產日報取込';
      const { data, error } = await supabase
        .from('pq_production_records')
        .insert(insertRow)
        .select('*')
        .single();
      if (error) throw error;
      saved.push(dbRowToClient(data));
    }

    return res.status(200).json({
      success: true,
      tabName,
      recordDateIso,
      imported: saved,
      skipped,
      errors,
    });
  } catch (e) {
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
