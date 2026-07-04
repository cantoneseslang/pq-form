import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { syncMaterialWidthToPlist } from '../lib/materialWidthSheets.js';

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

const apply = process.argv.includes('--apply');

try {
  const result = await syncMaterialWidthToPlist({ dryRun: !apply });
  console.log(JSON.stringify(result, null, 2));
  if (!apply) {
    console.log('\nDry run only. Re-run with --apply to write PQ-Form-plist column I.');
  }
} catch (e) {
  console.error(e.message || e);
  process.exit(1);
}
