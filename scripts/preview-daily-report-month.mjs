#!/usr/bin/env node
/** Preview or import daily report tabs for a given month. */
import { readFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { scanDailyReportTab, buildImportRecordsFromDailyTab } from '../lib/dailyReportImport.js';
import { resolveDailyReportSpreadsheetId } from '../lib/dailyReportSheetMap.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const envPaths = [
  resolve(__dirname, '../.env.production.local'),
  resolve(__dirname, '../.env.local'),
];
for (const envPath of envPaths) {
  if (!existsSync(envPath)) continue;
  for (const line of readFileSync(envPath, 'utf8').split('\n')) {
    const m = line.match(/^([^#=]+)=(.*)$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  }
}

const month = process.argv[2] || '1';
const tabName = process.argv[3] || '1';
const previewOnly = process.argv.includes('--preview');

const spreadsheetId = resolveDailyReportSpreadsheetId(month);
console.log(`Month ${month}, tab ${tabName}, sheet ${spreadsheetId}`);

try {
  const scanned = await scanDailyReportTab(tabName, { month });
  console.log('B1 date:', scanned.recordDateIso);
  console.log('Entries:', scanned.entries.length);

  const { imported, skipped, errors } = await buildImportRecordsFromDailyTab(tabName, new Set(), { month });
  console.log(`Would import: ${imported.length}, skip: ${skipped.length}, errors: ${errors.length}`);
  if (previewOnly && imported.length) {
    console.log('Sample:', JSON.stringify(imported[0], null, 2));
  }
  if (errors.length) console.log('Errors:', errors);
} catch (e) {
  console.error('FAILED:', e.message);
  process.exit(1);
}
