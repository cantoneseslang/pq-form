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
import {
  normalizeSheetMonth,
  normalizeSheetYear,
  resolveDailyReportSpreadsheetId,
} from '../../../lib/dailyReportSheetMap.js';
import { describeDailyReportSpreadsheet, readRange } from '../../../lib/dailyReportSheets.js';

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

function resolveImportOptions(body) {
  const sheetMonth = normalizeSheetMonth(body?.month ?? body?.sheetMonth ?? '');
  const sheetYear = normalizeSheetYear(body?.year ?? body?.sheetYear ?? '');
  return {
    sheetMonth,
    sheetYear,
    spreadsheetId: sheetMonth ? resolveDailyReportSpreadsheetId(sheetMonth, sheetYear) : null,
  };
}

function buildImportContext(options = {}) {
  const { sheetMonth, sheetYear } = options;
  if (!sheetMonth) return {};
  return {
    month: sheetMonth,
    ...(sheetYear ? { year: sheetYear } : {}),
  };
}

function sleep(ms) {
  return new Promise((resolve) => { setTimeout(resolve, ms); });
}

function recordPartsFromIso(isoDate) {
  const m = String(isoDate ?? '').match(/^(\d{4})-(\d{2})-/);
  if (!m) return { year: '', month: '' };
  return { year: m[1], month: String(parseInt(m[2], 10)) };
}

function isImportedDailyReportRow(row) {
  const { dailyReport, importedFromDailyReport } = unpackMaterialData(row.material_data || {});
  if (!dailyReport?.tabName || !dailyReport?.row) return false;
  if (importedFromDailyReport) return true;
  return String(row.correction_note || '').includes('日報取込');
}

function isMislabeledDailyImportRow(row, targetYear) {
  if (!isImportedDailyReportRow(row)) return false;
  const { dailyReport } = unpackMaterialData(row.material_data || {});
  const { year: dateYear, month: dateMonth } = recordPartsFromIso(row.record_date);
  if (!dateYear || !dateMonth) return false;
  if (targetYear && dateYear !== String(targetYear)) return false;

  const linkYear = String(dailyReport.year || dateYear);
  if (targetYear && linkYear !== String(targetYear)) return false;

  const linkMonth = String(dailyReport.month || '');
  if (!linkMonth || linkMonth === dateMonth) return false;
  return true;
}

async function findMislabeledDailyImports(supabase, targetYear) {
  const { data: rows, error: fetchError } = await supabase
    .from('pq_production_records')
    .select('id, record_date, material_data, correction_note')
    .is('deleted_at', null);
  if (fetchError) throw fetchError;

  return (rows || []).filter((row) => isMislabeledDailyImportRow(row, targetYear)).map((row) => {
    const { dailyReport } = unpackMaterialData(row.material_data || {});
    const { month: dateMonth } = recordPartsFromIso(row.record_date);
    return {
      id: row.id,
      recordDate: row.record_date,
      linkMonth: String(dailyReport?.month || ''),
      dateMonth,
      tabName: dailyReport?.tabName,
      rowNum: dailyReport?.row,
    };
  });
}

async function importDailyReportTab(tabName, supabase, options = {}) {
  const { sheetMonth, sheetYear } = options;
  const importOptions = buildImportContext(options);

  const scanned = await scanDailyReportTab(tabName, importOptions);
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
    importOptions,
  );

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

  return {
    tabName,
    sheetMonth: scanned.sheetMonth,
    sheetYear: scanned.sheetYear,
    recordDateIso,
    imported: saved,
    skipped,
    errors,
    wouldImport: imported,
  };
}

async function importDailyReportMonth(sheetMonth, supabase, options = {}) {
  const tabs = [];
  let totalImported = 0;
  let totalSkipped = 0;
  let totalErrors = 0;

  for (let day = 1; day <= 31; day += 1) {
    const tabName = String(day);
    try {
      if (day > 1) await sleep(2000);
      const result = await importDailyReportTab(tabName, supabase, { ...options, sheetMonth });
      if (!result.imported.length && !result.skipped.length && !result.recordDateIso) {
        tabs.push({ tabName, skipped: true, reason: 'empty or missing tab' });
        continue;
      }
      tabs.push({
        tabName,
        recordDateIso: result.recordDateIso,
        imported: result.imported.length,
        skipped: result.skipped.length,
        errors: result.errors.length,
      });
      totalImported += result.imported.length;
      totalSkipped += result.skipped.length;
      totalErrors += result.errors.length;
    } catch (tabErr) {
      tabs.push({ tabName, error: tabErr?.message || String(tabErr) });
      totalErrors += 1;
    }
  }

  return { tabs, totalImported, totalSkipped, totalErrors };
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
    const { sheetMonth, sheetYear, spreadsheetId } = resolveImportOptions(body);
    const preview = body?.preview === true;
    const repair = body?.repair === true;
    const importMonth = body?.importMonth === true;
    const importYear = body?.importYear === true;
    const inspect = body?.inspect === true;
    const cleanupMonth = body?.cleanupMonth === true;
    const cleanupMislabeled = body?.cleanupMislabeled === true;
    const supabase = getSupabaseAdmin();

    if (cleanupMislabeled) {
      const targetYear = sheetYear || '2025';
      const targets = await findMislabeledDailyImports(supabase, targetYear);

      if (preview) {
        const byLinkMonth = {};
        for (const row of targets) {
          byLinkMonth[row.linkMonth] = (byLinkMonth[row.linkMonth] || 0) + 1;
        }
        return res.status(200).json({
          success: true,
          cleanupMislabeled: true,
          preview: true,
          year: targetYear,
          wouldDelete: targets.length,
          byLinkMonth,
          samples: targets.slice(0, 10),
        });
      }

      const now = new Date().toISOString();
      const deleted = [];
      for (const row of targets) {
        const { error: updateError } = await supabase
          .from('pq_production_records')
          .update({ deleted_at: now, updated_at: now })
          .eq('id', row.id);
        if (updateError) throw updateError;
        deleted.push(row);
      }

      const byLinkMonth = {};
      for (const row of deleted) {
        byLinkMonth[row.linkMonth] = (byLinkMonth[row.linkMonth] || 0) + 1;
      }

      return res.status(200).json({
        success: true,
        cleanupMislabeled: true,
        year: targetYear,
        deletedCount: deleted.length,
        byLinkMonth,
        deleted,
      });
    }

    if (cleanupMonth) {
      if (!sheetMonth) {
        return res.status(400).json({ success: false, error: 'month (1-12) is required for cleanupMonth' });
      }
      const { data: rows, error: fetchError } = await supabase
        .from('pq_production_records')
        .select('id, record_date, material_data, correction_note')
        .is('deleted_at', null);
      if (fetchError) throw fetchError;

      const targets = (rows || []).filter((row) => {
        const { dailyReport } = unpackMaterialData(row.material_data || {});
        if (!isImportedDailyReportRow(row)) return false;
        if (String(dailyReport?.month || '') !== String(sheetMonth)) return false;
        if (sheetYear) {
          const rowYear = String(dailyReport?.year || recordPartsFromIso(row.record_date).year || '2026');
          return rowYear === String(sheetYear);
        }
        return true;
      });

      const now = new Date().toISOString();
      const deleted = [];
      for (const row of targets) {
        const { error: updateError } = await supabase
          .from('pq_production_records')
          .update({ deleted_at: now, updated_at: now })
          .eq('id', row.id);
        if (updateError) throw updateError;
        deleted.push({ id: row.id, recordDate: row.record_date });
      }

      return res.status(200).json({
        success: true,
        cleanupMonth: true,
        year: sheetYear || null,
        month: sheetMonth,
        deletedCount: deleted.length,
        deleted,
      });
    }

    if (inspect) {
      if (!sheetMonth) {
        return res.status(400).json({ success: false, error: 'month (1-12) is required for inspect' });
      }
      const info = await describeDailyReportSpreadsheet(sheetMonth, sheetYear);
      const tabSamples = {};
      for (const name of (info.sheetNames || []).slice(0, 8)) {
        try {
          tabSamples[name] = await readRange(`'${name.replace(/'/g, "''")}'!A1:B2`, sheetMonth, sheetYear);
        } catch (sampleErr) {
          tabSamples[name] = { error: sampleErr?.message || String(sampleErr) };
        }
      }
      return res.status(200).json({ success: true, inspect: true, ...info, tabSamples });
    }

    if (importYear) {
      if (!sheetYear) {
        return res.status(400).json({ success: false, error: 'year (e.g. 2025) is required for importYear' });
      }

      const months = [];
      let grandImported = 0;
      let grandSkipped = 0;
      let grandErrors = 0;

      for (let monthNum = 1; monthNum <= 12; monthNum += 1) {
        const month = String(monthNum);
        try {
          if (monthNum > 1) await sleep(3000);
          if (preview) {
            let totalImported = 0;
            let totalSkipped = 0;
            let totalErrors = 0;
            for (let day = 1; day <= 31; day += 1) {
              const tabName = String(day);
              const scanned = await scanDailyReportTab(tabName, { month, year: sheetYear });
              if (!scanned.entries.length && !scanned.recordDateIso) continue;
              const { imported, skipped, errors } = await buildImportRecordsFromDailyTab(
                tabName,
                new Set(),
                { month, year: sheetYear },
              );
              totalImported += imported.length;
              totalSkipped += skipped.length;
              totalErrors += errors.length;
            }
            months.push({ month, wouldImport: totalImported, skipped: totalSkipped, errors: totalErrors });
            grandImported += totalImported;
            grandSkipped += totalSkipped;
            grandErrors += totalErrors;
          } else {
            const result = await importDailyReportMonth(month, supabase, { sheetMonth: month, sheetYear });
            months.push({
              month,
              imported: result.totalImported,
              skipped: result.totalSkipped,
              errors: result.totalErrors,
            });
            grandImported += result.totalImported;
            grandSkipped += result.totalSkipped;
            grandErrors += result.totalErrors;
          }
        } catch (monthErr) {
          months.push({ month, error: monthErr?.message || String(monthErr) });
          grandErrors += 1;
        }
      }

      return res.status(200).json({
        success: true,
        importYear: true,
        preview,
        year: sheetYear,
        totalImported: grandImported,
        totalSkipped: grandSkipped,
        totalErrors: grandErrors,
        months,
      });
    }

    if (importMonth) {
      if (!sheetMonth) {
        return res.status(400).json({ success: false, error: 'month (1-12) is required for importMonth' });
      }

      if (preview) {
        const tabs = [];
        let totalImported = 0;
        let totalSkipped = 0;
        let totalErrors = 0;
        for (let day = 1; day <= 31; day += 1) {
          const tabName = String(day);
          try {
            const scanned = await scanDailyReportTab(tabName, buildImportContext({ sheetMonth, sheetYear }));
            if (!scanned.entries.length && !scanned.recordDateIso) {
              tabs.push({ tabName, skipped: true, reason: 'empty or missing tab' });
              continue;
            }
            const { recordDateIso, imported, skipped, errors } = await buildImportRecordsFromDailyTab(
              tabName,
              new Set(),
              buildImportContext({ sheetMonth, sheetYear }),
            );
            tabs.push({
              tabName,
              recordDateIso,
              wouldImport: imported.length,
              skipped: skipped.length,
              errors: errors.length,
            });
            totalImported += imported.length;
            totalSkipped += skipped.length;
            totalErrors += errors.length;
          } catch (tabErr) {
            tabs.push({ tabName, error: tabErr?.message || String(tabErr) });
            totalErrors += 1;
          }
        }
        return res.status(200).json({
          success: true,
          importMonth: true,
          preview: true,
          year: sheetYear || null,
          month: sheetMonth,
          spreadsheetId,
          totalImported,
          totalSkipped,
          totalErrors,
          tabs,
        });
      }

      const result = await importDailyReportMonth(sheetMonth, supabase, { sheetMonth, sheetYear });
      return res.status(200).json({
        success: true,
        importMonth: true,
        preview: false,
        year: sheetYear || null,
        month: sheetMonth,
        spreadsheetId,
        totalImported: result.totalImported,
        totalSkipped: result.totalSkipped,
        totalErrors: result.totalErrors,
        tabs: result.tabs,
      });
    }

    const tabName = String(body?.tabName ?? body?.day ?? '').trim();
    if (!tabName || !/^\d{1,2}$/.test(tabName)) {
      return res.status(400).json({ success: false, error: 'tabName (1-31) is required' });
    }

    const importOptions = buildImportContext({ sheetMonth, sheetYear });

    if (repair) {
      const {
        recordDateIso,
        repaired,
        skipped,
        sheetMonth: resolvedMonth,
        sheetYear: resolvedYear,
      } = await repairImportedProductNamesForTab(
        tabName,
        supabase,
        importOptions,
      );
      return res.status(200).json({
        success: true,
        repair: true,
        tabName,
        year: resolvedYear,
        month: resolvedMonth,
        spreadsheetId: resolveDailyReportSpreadsheetId(resolvedMonth, resolvedYear),
        recordDateIso,
        repaired,
        skipped,
      });
    }

    if (preview) {
      const scanned = await scanDailyReportTab(tabName, importOptions);
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
        importOptions,
      );

      return res.status(200).json({
        success: true,
        preview: true,
        tabName,
        year: scanned.sheetYear,
        month: scanned.sheetMonth,
        spreadsheetId: resolveDailyReportSpreadsheetId(scanned.sheetMonth, scanned.sheetYear),
        recordDateIso,
        wouldImport: imported,
        skipped,
        errors,
      });
    }

    const result = await importDailyReportTab(tabName, supabase, { sheetMonth, sheetYear });
    return res.status(200).json({
      success: true,
      tabName,
      year: result.sheetYear,
      month: result.sheetMonth,
      spreadsheetId: resolveDailyReportSpreadsheetId(result.sheetMonth, result.sheetYear),
      recordDateIso: result.recordDateIso,
      imported: result.imported,
      skipped: result.skipped,
      errors: result.errors,
    });
  } catch (e) {
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
