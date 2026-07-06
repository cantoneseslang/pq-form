import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { syncMaterialStockLotsToSheet } from '../lib/materialStockSync.js';

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

const dryRun = !process.argv.includes('--apply');

try {
  const result = await syncMaterialStockLotsToSheet({ dryRun });
  console.log(JSON.stringify(result, null, 2));
  if (dryRun) {
    console.log('\nDry run only. Re-run with --apply to write the sheet.');
  }
} catch (error) {
  console.error(error?.message || error);
  process.exitCode = 1;
}
