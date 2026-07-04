import { writeFileSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { analyzeMaterialWidthSheets } from '../lib/materialWidthSheets.js';

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

const outPath = process.argv.includes('--csv')
  ? resolve(__dirname, '../output/material-width-recommendations.csv')
  : null;
const jsonOnly = process.argv.includes('--json');

try {
  const report = await analyzeMaterialWidthSheets();
  if (jsonOnly) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log('Service account:', report.serviceAccountEmail);
    console.log('Access:', report.access);
    console.log('Sheets:', JSON.stringify(report.sheets, null, 2));
    console.log('Stats:', report.stats);
    console.log('Cross-ref:', report.crossRef);
    console.log('Unresolved sample:', report.unresolvedSample.length);
    console.log('Conflict sample:', report.conflictSample.length);
  }

  if (outPath) {
    const { mkdirSync } = await import('fs');
    mkdirSync(resolve(__dirname, '../output'), { recursive: true });
    const lines = ['row,code,name,existing,recommended,source'];
    for (const r of report.recommendations) {
      lines.push([
        r.rowIndex,
        r.code,
        `"${String(r.name).replace(/"/g, '""')}"`,
        r.existing,
        r.recommended,
        r.source,
      ].join(','));
    }
    writeFileSync(outPath, `${lines.join('\n')}\n`, 'utf8');
    console.log('Wrote', outPath);
  }

  if (!report.access.monthlyDes) process.exitCode = 2;
} catch (e) {
  console.error(e.message || e);
  process.exit(1);
}
