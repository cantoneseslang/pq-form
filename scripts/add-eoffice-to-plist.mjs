import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { addEofficeOnlyItemsToPlist } from '../lib/materialWidthSheets.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const envPaths = [
  resolve(__dirname, '../.env.production.local'),
  resolve(__dirname, '../.env.local'),
  '/tmp/pq-vercel-env.prod',
  '/tmp/pq-form-vercel-env.tmp',
];

for (const envPath of envPaths) {
  try {
    const { readFileSync } = await import('fs');
    const content = readFileSync(envPath, 'utf8');
    for (const line of content.split('\n')) {
      const m = line.match(/^([^#=]+)=(.*)$/);
      if (!m) continue;
      let v = m[2].trim().replace(/^["']|["']$/g, '').replace(/\\n$/g, '');
      if (!process.env[m[1].trim()] || process.env[m[1].trim()] === '""') {
        process.env[m[1].trim()] = v;
      }
    }
  } catch {
    // ignore
  }
}

const apply = process.argv.includes('--apply');

try {
  const result = await addEofficeOnlyItemsToPlist({ dryRun: !apply });
  console.log(JSON.stringify(result, null, 2));
  if (!apply) {
    console.log('\nDry run only. Re-run with --apply to append rows to PQ-Form-plist.');
  }
} catch (e) {
  console.error(e.message || e);
  process.exit(1);
}
