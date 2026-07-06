/**
 * 生產紀錄の全 mainLines を plist 照合し、カバレッジをレポートする。
 * Usage: node scripts/test-production-records-plist.mjs [--api https://pq-form.vercel.app]
 */
import { readFileSync, writeFileSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { createClient } from '@supabase/supabase-js';
import { inferProductTypeKeyFromName } from '../lib/dailyReportImport.js';
import { thicknessForProductLookup, NOT_FOUND_PRODUCT_CODE } from '../lib/plist.js';
import { dbRowToClient } from '../lib/productionRecords.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const API_BASE = (process.argv.find((a) => a.startsWith('--api='))?.slice(6))
  || process.env.PQFORM_API_BASE
  || 'https://pq-form.vercel.app';

function loadEnvFile(name) {
  try {
    const env = readFileSync(resolve(__dirname, '..', name), 'utf8');
    for (const line of env.split('\n')) {
      const m = line.match(/^([^#=]+)=(.*)$/);
      if (!m) continue;
      let v = m[2].trim();
      if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) v = v.slice(1, -1);
      v = v.replace(/\\n$/, '');
      process.env[m[1].trim()] = v;
    }
  } catch {
    // optional
  }
}

loadEnvFile('.env.production.local');
loadEnvFile('.env.local');
loadEnvFile('/tmp/pq-vercel-env.prod');

const PRODUCT_TYPE_KEYS = ['企筒', '地槽', '鐵角', '批灰角', 'W角', '闊槽', 'C槽', 'CT企筒打孔', '其他'];

function resolveType(record, line) {
  const types = record.productTypes || {};
  for (const key of PRODUCT_TYPE_KEYS) {
    if (types[key]) return key === 'CT企筒打孔' ? '企筒' : key;
  }
  return inferProductTypeKeyFromName(line.name || '') || '';
}

function specKey({ type, thickness, width, height, length }) {
  return [type, thickness, width, height, length].join('|');
}

function hasSpec(line) {
  return line.thickness && line.width && line.height && line.length;
}

async function fetchRecords() {
  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (url && key) {
    const supabase = createClient(url, key);
    const all = [];
    let from = 0;
    const page = 1000;
    while (true) {
      const { data, error } = await supabase
        .from('pq_production_records')
        .select('*')
        .is('deleted_at', null)
        .order('created_at', { ascending: false })
        .range(from, from + page - 1);
      if (error) throw error;
      all.push(...(data || []).map(dbRowToClient));
      if (!data || data.length < page) break;
      from += page;
    }
    return all;
  }

  const res = await fetch(`${API_BASE}/api/pq_form/production_records`, { cache: 'no-store' });
  const json = await res.json();
  if (!json.success) throw new Error(json.error || 'records fetch failed');
  return json.records || [];
}

async function searchPlist({ type, thickness, width, height, length }) {
  const params = new URLSearchParams({
    type: type || '企筒',
    t: thicknessForProductLookup(thickness),
    w: String(width),
    h: String(height),
    l: String(length),
  });
  const res = await fetch(`${API_BASE}/api/pq_form/plist/search?${params}`, { cache: 'no-store' });
  const data = await res.json();
  if (!data.success) {
    return { matches: [], hint: data.error || 'search failed' };
  }
  return data;
}

async function main() {
  console.log('API:', API_BASE);
  const records = await fetchRecords();
  console.log('records:', records.length);

  const specMap = new Map();
  let lineCount = 0;
  let skipNoType = 0;
  let skipIncomplete = 0;

  for (const record of records) {
    const lines = record.mainLines?.length ? record.mainLines : (record.main ? [record.main] : []);
    for (const line of lines) {
      lineCount += 1;
      if (!hasSpec(line)) {
        skipIncomplete += 1;
        continue;
      }
      const type = resolveType(record, line);
      if (!type || type === '其他') {
        skipNoType += 1;
        continue;
      }
      const spec = {
        type,
        thickness: String(line.thickness).trim(),
        width: String(line.width).trim(),
        height: String(line.height).trim(),
        length: String(line.length).trim(),
      };
      const key = specKey(spec);
      if (!specMap.has(key)) {
        specMap.set(key, {
          ...spec,
          storedCodes: new Set(),
          names: new Set(),
          recordCount: 0,
        });
      }
      const bucket = specMap.get(key);
      bucket.recordCount += 1;
      if (line.productNo) bucket.storedCodes.add(String(line.productNo).trim());
      if (line.name) bucket.names.add(String(line.name).trim());
    }
  }

  const uniqueSpecs = [...specMap.values()];
  console.log('main lines:', lineCount);
  console.log('unique specs:', uniqueSpecs.length);
  console.log('skip incomplete:', skipIncomplete, 'skip no type:', skipNoType);

  const results = {
    testedAt: new Date().toISOString(),
    apiBase: API_BASE,
    recordCount: records.length,
    lineCount,
    skipIncomplete,
    skipNoType,
    testableLines: lineCount - skipIncomplete - skipNoType,
    uniqueSpecCount: uniqueSpecs.length,
    hit: 0,
    miss: 0,
    hitLines: 0,
    missLines: 0,
    codeMismatch: 0,
    codeMatch: 0,
    codeMatchLines: 0,
    codeMismatchLines: 0,
    storedCodeLines: 0,
    storedCodeMissing: 0,
    nameMismatch: 0,
    nameMatch: 0,
    nameMismatchLines: 0,
    nameMatchLines: 0,
    notFoundCodeLines: 0,
    typeAdjusted: 0,
    ambiguous: 0,
    misses: [],
    mismatches: [],
    nameMismatches: [],
    typeAdjustments: [],
  };

  function normalizeName(text) {
    return String(text ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
  }

  function namesAlign(recordName, plistName) {
    const a = normalizeName(recordName);
    const b = normalizeName(plistName);
    if (!a || !b) return false;
    if (a === b) return true;
    // 0.8x50x25 地槽 300mm vs 0.8x50x25 地槽 300mm
    const dims = a.match(/^([\d.]+)x([\d.]+)x([\d.]+)\s+(.+?)\s+([\d.]+)mm/);
    const dimsB = b.match(/^([\d.]+)x([\d.]+)x([\d.]+)\s+(.+?)\s+([\d.]+)mm/);
    if (dims && dimsB) {
      return dims[1] === dimsB[1] && dims[2] === dimsB[2] && dims[3] === dimsB[3]
        && dims[4] === dimsB[4] && dims[5] === dimsB[5];
    }
    return a.includes(b) || b.includes(a);
  }

  let i = 0;
  for (const spec of uniqueSpecs) {
    i += 1;
    if (i % 20 === 0) process.stderr.write(`testing ${i}/${uniqueSpecs.length}\n`);
    const data = await searchPlist(spec);
    const matches = data.matches || [];
    const stored = [...spec.storedCodes].filter((c) => c && c !== NOT_FOUND_PRODUCT_CODE);

    if (!matches.length) {
      results.miss += 1;
      results.missLines += spec.recordCount;
      results.misses.push({
        ...spec,
        storedCodes: [...spec.storedCodes],
        sampleName: [...spec.names][0],
        recordCount: spec.recordCount,
        hint: data.hint,
      });
      continue;
    }

    results.hit += 1;
    results.hitLines += spec.recordCount;
    if (data.typeAdjusted) {
      results.typeAdjusted += 1;
      results.typeAdjustments.push({
        requestedType: spec.type,
        resolvedType: data.resolvedType,
        spec: `${spec.thickness}x${spec.width}x${spec.height} ${spec.type} ${spec.length}mm`,
        code: matches[0]?.code,
        recordCount: spec.recordCount,
      });
    }
    if (matches.length > 1) results.ambiguous += 1;

    if (stored.length) {
      const codes = new Set(matches.map((m) => m.code.toUpperCase()));
      const plistNames = matches.map((m) => m.name);
      const ok = stored.some((c) => codes.has(c.toUpperCase()));
      results.storedCodeLines += spec.recordCount;
      if (ok) {
        results.codeMatch += 1;
        results.codeMatchLines += spec.recordCount;
      } else {
        results.codeMismatch += 1;
        results.codeMismatchLines += spec.recordCount;
        results.mismatches.push({
          ...spec,
          storedCodes: stored,
          plistCodes: matches.map((m) => m.code),
          plistNames,
          sampleName: [...spec.names][0],
          recordCount: spec.recordCount,
        });
      }
      const sampleName = [...spec.names][0] || '';
      const nameOk = plistNames.some((pn) => namesAlign(sampleName, pn));
      if (nameOk) {
        results.nameMatch += 1;
        results.nameMatchLines += spec.recordCount;
      } else if (sampleName) {
        results.nameMismatch += 1;
        results.nameMismatchLines += spec.recordCount;
        results.nameMismatches.push({
          ...spec,
          storedCodes: stored,
          recordName: sampleName,
          plistNames,
          plistCodes: matches.map((m) => m.code),
          recordCount: spec.recordCount,
        });
      }
    } else {
      const onlyNotFound = [...spec.storedCodes].every((c) => !c || c === NOT_FOUND_PRODUCT_CODE);
      if (onlyNotFound) results.notFoundCodeLines += spec.recordCount;
    }
  }

  // 保存コードがある行のうち plist 未ヒット分
  for (const m of results.misses) {
    const stored = m.storedCodes.filter((c) => c && c !== NOT_FOUND_PRODUCT_CODE);
    if (stored.length) results.storedCodeMissing += m.recordCount;
  }

  const outPath = resolve(__dirname, '../output/production-records-plist-test.json');
  writeFileSync(outPath, JSON.stringify(results, null, 2));

  console.log('\n=== RESULT ===');
  console.log('plist hit (unique spec):', results.hit, '/', uniqueSpecs.length,
    `(${((results.hit / uniqueSpecs.length) * 100).toFixed(1)}%)`);
  console.log('plist hit (line-weighted):', results.hitLines, '/', results.testableLines,
    `(${((results.hitLines / results.testableLines) * 100).toFixed(1)}%)`);
  console.log('plist miss:', results.miss, 'specs,', results.missLines, 'lines');
  console.log('stored code lines (has code):', results.storedCodeLines);
  console.log('code match (spec with stored code vs plist):', results.codeMatch, 'specs,', results.codeMatchLines, 'lines');
  console.log('code mismatch:', results.codeMismatch, 'specs,', results.codeMismatchLines, 'lines');
  console.log('name match:', results.nameMatch, 'specs,', results.nameMatchLines, 'lines');
  console.log('name mismatch:', results.nameMismatch, 'specs,', results.nameMismatchLines, 'lines');
  console.log('not-found code lines:', results.notFoundCodeLines);
  console.log('stored code but plist miss:', results.storedCodeMissing, 'lines');
  console.log('type auto-adjusted:', results.typeAdjusted);
  console.log('ambiguous multi-match:', results.ambiguous);
  console.log('report:', outPath);
  if (results.mismatches.length) {
    console.log('\nCode mismatches (stored ≠ plist search):');
    for (const m of [...results.mismatches].sort((a, b) => b.recordCount - a.recordCount).slice(0, 10)) {
      console.log(`  ${m.recordCount}x ${m.thickness}x${m.width}x${m.height} ${m.type} ${m.length}mm`);
      console.log(`    record: ${m.storedCodes.join(',')} | plist: ${m.plistCodes.join(',')}`);
    }
  }
  if (results.nameMismatches.length) {
    console.log('\nName mismatches (code OK but name differs):');
    for (const m of [...results.nameMismatches].sort((a, b) => b.recordCount - a.recordCount).slice(0, 10)) {
      console.log(`  ${m.recordCount}x stored=${m.storedCodes.join(',')}`);
      console.log(`    record name: ${m.recordName}`);
      console.log(`    plist name:  ${m.plistNames[0]}`);
    }
  }
  if (results.misses.length) {
    console.log('\nTop plist misses (by record count):');
    for (const m of [...results.misses].sort((a, b) => b.recordCount - a.recordCount).slice(0, 15)) {
      console.log(`  ${m.recordCount}x ${m.thickness}x${m.width}x${m.height} ${m.type} ${m.length}mm | stored: ${m.storedCodes.join(',') || '-'} | ${m.hint || ''}`);
    }
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
