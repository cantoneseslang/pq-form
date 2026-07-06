import { writeFileSync, mkdirSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { applyPlistDedupe } from '../lib/plistDuplicates.js';

const __dirname = dirname(fileURLToPath(import.meta.url));

async function loadEnv() {
  const { readFileSync } = await import('fs');
  for (const envPath of [
    '/tmp/pq-vercel-env.prod',
    resolve(__dirname, '../.env.production.local'),
    resolve(__dirname, '../.env.local'),
  ]) {
    try {
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
}

async function main() {
  await loadEnv();
  const apply = process.argv.includes('--apply');
  const jsonOnly = process.argv.includes('--json');

  const result = await applyPlistDedupe({ dryRun: !apply });

  if (jsonOnly) {
    console.log(JSON.stringify(result, null, 2));
  } else {
    console.log('total rows:', result.totalRows);
    console.log('unique codes:', result.uniqueCodes);
    console.log('duplicate codes:', result.duplicateCodeCount);
    console.log('rows to remove:', result.plan.deleteCount);
    console.log('merge material width:', result.plan.mergeCount);
    console.log('categories:', result.categories);
    console.log('\nTop duplicates:');
    for (const g of result.duplicateGroups.slice(0, 15)) {
      console.log(`  ${g.code} x${g.count} [${g.category}] keep row ${g.keep.rowIndex}`);
      for (const v of g.variants) {
        console.log(`    variant row ${v.sample.rowIndex}: ${v.sample.pqFormDesc || v.sample.pdesc1} | I=${v.sample.materialWidth || '-'}`);
        if (v.diffsFromKeep.length) {
          console.log(`      diffs: ${v.diffsFromKeep.map((d) => `${d.column}:${d.a}≠${d.b}`).join(', ')}`);
        }
      }
    }
  }

  mkdirSync(resolve(__dirname, '../output'), { recursive: true });
  writeFileSync(
    resolve(__dirname, '../output/plist-duplicates-report.json'),
    JSON.stringify(result, null, 2),
  );

  const csvLines = ['code,category,count,keep_row,remove_rows,conflict_columns,keep_desc,variant_descs'];
  for (const g of result.duplicateGroups) {
    const removeRows = g.removeRows.map((r) => r.rowIndex).join(';');
    const conflictCols = [...new Set(g.variants.flatMap((v) => v.diffsFromKeep.map((d) => d.column)))].join(';');
    const variantDescs = g.variants.map((v) => `row${v.sample.rowIndex}:${v.sample.pqFormDesc || v.sample.pdesc1}`).join(' | ');
    csvLines.push([
      g.code,
      g.category,
      g.count,
      g.keep.rowIndex,
      `"${removeRows}"`,
      `"${conflictCols}"`,
      `"${String(g.keep.pqFormDesc).replace(/"/g, '""')}"`,
      `"${variantDescs.replace(/"/g, '""')}"`,
    ].join(','));
  }
  writeFileSync(resolve(__dirname, '../output/plist-duplicates.csv'), `${csvLines.join('\n')}\n`);

  if (!apply) {
    console.log('\nWrote output/plist-duplicates-report.json and output/plist-duplicates.csv');
    console.log('Re-run with --apply to delete duplicate rows (keeps best row per code).');
  } else {
    console.log('\nApplied:', result.deleted, 'rows deleted,', result.merged, 'material widths merged.');
    console.log('Rows after:', result.rowsAfter);
  }
}

main().catch((e) => {
  console.error(e.message || e);
  process.exit(1);
});
