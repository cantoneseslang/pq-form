import { readFileSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { google } from 'googleapis';

const __dirname = dirname(fileURLToPath(import.meta.url));
const envPaths = [
  resolve(__dirname, '../.env.production.local'),
  resolve(__dirname, '../.env.local'),
  '/tmp/pq-form-vercel-env.tmp',
];

for (const envPath of envPaths) {
  try {
    const content = readFileSync(envPath, 'utf8');
    for (const line of content.split('\n')) {
      const m = line.match(/^([^#=]+)=(.*)$/);
      if (m && !process.env[m[1].trim()]) {
        process.env[m[1].trim()] = m[2].replace(/^["']|["']$/g, '');
      }
    }
  } catch {
    // ignore missing files
  }
}

const raw = process.env.GOOGLE_SA_JSON || '';
if (!raw) {
  console.error('GOOGLE_SA_JSON not found');
  process.exit(1);
}

const sa = JSON.parse(raw);
if (sa.private_key) sa.private_key = sa.private_key.replace(/\\n/g, '\n');

const jwt = new google.auth.JWT(
  sa.client_email,
  undefined,
  sa.private_key,
  ['https://www.googleapis.com/auth/spreadsheets.readonly'],
);

const sheets = google.sheets({ version: 'v4', auth: jwt });
const spreadsheetId = '1ivqlw58PKeyXWwap-lddnlYiWpWG1nrn4EneNRbjlEU';
const targetGid = 863784501;

const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
const tabs = (res.data.sheets || []).map((s) => s.properties).filter(Boolean);
const match = tabs.find((p) => p.sheetId === targetGid);

console.log(JSON.stringify({
  spreadsheetId,
  serviceAccountEmail: sa.client_email,
  tabs: tabs.map((p) => ({ title: p.title, sheetId: p.sheetId })),
  matchedTab: match ? match.title : null,
}, null, 2));

if (!match) process.exit(2);
