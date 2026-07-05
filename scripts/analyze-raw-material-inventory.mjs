import { mkdirSync, writeFileSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import {
  analyzeRawMaterialInventory,
  buildLotsCsvRows,
  buildSummaryCsvRows,
} from '../lib/rawMaterialInventorySheets.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const envPaths = [
  resolve(__dirname, '../.env.production.local'),
  resolve(__dirname, '../.env.local'),
  '/tmp/pq-form-vercel-env.tmp',
];

for (const envPath of envPaths) {
  try {
    const { readFileSync } = await import('fs');
    const content = readFileSync(envPath, 'utf8');
    for (const line of content.split('\n')) {
      const m = line.match(/^([^#=]+)=(.*)$/);
      if (m && !process.env[m[1].trim()]) {
        process.env[m[1].trim()] = m[2].replace(/^["']|["']$/g, '');
      }
    }
  } catch {
    // ignore
  }
}

const jsonOnly = process.argv.includes('--json');
const writeCsv = process.argv.includes('--csv');
const tabArg = process.argv.find((a) => a.startsWith('--tab='));
const tabFilter = tabArg ? [tabArg.slice(6)] : null;

try {
  const report = await analyzeRawMaterialInventory({ tabFilter });

  if (jsonOnly) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log('Service account:', report.serviceAccountEmail);
    console.log('Spreadsheet:', report.spreadsheetTitle, report.spreadsheetId);
    console.log('Inventory tabs:', report.inventoryTabs.length);
    console.log('Stats:', report.stats);
    console.log('Errors:', report.access.errors);
    for (const tr of report.tabReports) {
      console.log(`\n[${tr.tabTitle}] lots=${tr.lotCount} apAy=${tr.apAy.interpretation}`);
      console.log('  AP-AY headers:', tr.apAy.headers.join(' | '));
      console.log('  AP-AY non-empty:', tr.apAy.nonEmptyCounts.join(', '));
    }
    console.log('\nTop summaries (by kg):');
    const top = [...report.summaries].sort((a, b) => b.totalKg - a.totalKg).slice(0, 15);
    for (const s of top) {
      console.log(`  ${s.tabTitle} / ${s.materialWidth}mm → ${s.totalKg} kg (${s.lotCount} lots)`);
    }
  }

  if (writeCsv) {
    const outDir = resolve(__dirname, '../output');
    mkdirSync(outDir, { recursive: true });
    const allLots = report.tabReports.flatMap((t) => t.lots);
    writeFileSync(
      resolve(outDir, 'raw-material-inventory-summary.csv'),
      `${buildSummaryCsvRows(report.summaries).join('\n')}\n`,
      'utf8',
    );
    writeFileSync(
      resolve(outDir, 'raw-material-inventory-lots.csv'),
      `${buildLotsCsvRows(allLots).join('\n')}\n`,
      'utf8',
    );
    writeFileSync(
      resolve(outDir, 'raw-material-inventory-report.json'),
      `${JSON.stringify(report, null, 2)}\n`,
      'utf8',
    );
    console.log('Wrote output/raw-material-inventory-*.csv/json');
  }

  if (report.access.errors.length) process.exitCode = 2;
} catch (e) {
  console.error(e.message || e);
  process.exit(1);
}
