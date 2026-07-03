import { getSupabaseAdmin, isSupabaseConfigured } from '../../../lib/supabase.js';
import {
  clientToDbInsert,
  dbRowToClient,
  normalizeRecordDate,
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

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('Access-Control-Allow-Origin', '*');

  if (!isSupabaseConfigured()) {
    return res.status(503).json({ success: false, error: 'Supabase is not configured' });
  }

  try {
    const supabase = getSupabaseAdmin();

    if (req.method === 'GET') {
      const url = new URL(req.url, `http://${req.headers.host}`);
      const date = normalizeRecordDate(url.searchParams.get('date') || '');
      const pageType = url.searchParams.get('page') || '';

      let query = supabase
        .from('pq_production_records')
        .select('*')
        .is('deleted_at', null)
        .order('created_at', { ascending: false });

      if (date) query = query.eq('record_date', date);
      if (pageType) query = query.eq('page_type', pageType);

      const { data, error } = await query;
      if (error) throw error;

      return res.status(200).json({
        success: true,
        records: (data || []).map(dbRowToClient),
      });
    }

    if (req.method === 'POST') {
      const body = await parseJsonBody(req);
      const insertRow = clientToDbInsert(body);
      if (!insertRow.record_date || !insertRow.page_type || !insertRow.main_data) {
        return res.status(400).json({
          success: false,
          error: 'recordDate, pageType, and main are required',
        });
      }

      const { data, error } = await supabase
        .from('pq_production_records')
        .insert(insertRow)
        .select('*')
        .single();

      if (error) throw error;
      return res.status(200).json({ success: true, record: dbRowToClient(data) });
    }

    return res.status(405).json({ success: false, error: 'Method not allowed' });
  } catch (e) {
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
}
