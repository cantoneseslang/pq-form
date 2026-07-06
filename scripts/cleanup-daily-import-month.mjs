#!/usr/bin/env node
import { readFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { getSupabaseAdmin } from '../lib/supabase.js';
import { unpackMaterialData } from '../lib/productionRecords.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
for (const envPath of [
  resolve(__dirname, '../.env.production.local'),
  resolve(__dirname, '../.env.local'),
]) {
  if (!existsSync(envPath)) continue;
  for (const line of readFileSync(envPath, 'utf8').split('\n')) {
    const m = line.match(/^([^#=]+)=(.*)$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  }
}

const months = (process.argv[2] || '4,5').split(',').map((v) => v.trim()).filter(Boolean);
const dryRun = process.argv.includes('--dry-run');

const supabase = getSupabaseAdmin();
const { data: rows, error } = await supabase
  .from('pq_production_records')
  .select('id, record_date, material_data')
  .is('deleted_at', null);

if (error) throw error;

const targets = (rows || []).filter((row) => {
  const { dailyReport, importedFromDailyReport } = unpackMaterialData(row.material_data || {});
  if (!importedFromDailyReport || !dailyReport?.month) return false;
  return months.includes(String(dailyReport.month));
});

console.log(`Found ${targets.length} imported records for month(s): ${months.join(', ')}`);
if (dryRun) {
  for (const row of targets.slice(0, 5)) {
    const { dailyReport } = unpackMaterialData(row.material_data || {});
    console.log(' sample', row.id, row.record_date, dailyReport);
  }
  process.exit(0);
}

const now = new Date().toISOString();
let deleted = 0;
for (const row of targets) {
  const { error: updateError } = await supabase
    .from('pq_production_records')
    .update({ deleted_at: now, updated_at: now })
    .eq('id', row.id);
  if (updateError) throw updateError;
  deleted += 1;
}
console.log(`Soft-deleted ${deleted} records.`);
