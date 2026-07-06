import { readFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import {
  pickTargetRow,
  resolveMoldingMachineBlock,
  formatDailyProductName,
} from '../lib/dailyReport.js';
import { readRange } from '../lib/dailyReportSheets.js';

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

const tab = process.argv[2] || '26';
const machine = process.argv[3] || '3號滾壓成型機';

const mainLine = {
  thickness: '0.4',
  width: '64',
  height: '32',
  length: '2440',
};
const productTypes = { 批灰角: true };
const productLabel = formatDailyProductName(mainLine, productTypes);

const rows = await readRange(`${tab}!A1:L100`);
const block = resolveMoldingMachineBlock(rows, machine);

console.log('machine:', machine);
console.log('product B:', productLabel);
console.log('block headerRow:', block?.headerRow);
console.log('block slotRows:', block?.slotRows);

for (const rowNum of block?.slotRows || []) {
  const row = rows[rowNum - 1] || [];
  const a = row[0] ?? '';
  const b = row[1] ?? '';
  const c = row[2] ?? '';
  const d = row[3] ?? '';
  const f = row[5] ?? '';
  const hasData = a || b || d || f;
  console.log(`  row ${rowNum}: A=${JSON.stringify(String(a))} B=${JSON.stringify(String(b).slice(0, 40))} C=${JSON.stringify(String(c).slice(0, 20))} ${hasData ? 'USED' : 'empty'}`);
}

const target = pickTargetRow(rows, block, productLabel);
console.log('pickTargetRow →', target);
