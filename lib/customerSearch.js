import { createClient } from '@supabase/supabase-js';

const DEFAULT_SALES_TABLE = process.env.SALES_DATA_TABLE || 'sales';
const DEFAULT_LIMIT = 12;

export function isSalesSupabaseConfigured() {
  return !!(process.env.SALES_SUPABASE_URL && process.env.SALES_SUPABASE_SERVICE_ROLE_KEY);
}

function getSalesSupabase() {
  const url = process.env.SALES_SUPABASE_URL;
  const key = process.env.SALES_SUPABASE_SERVICE_ROLE_KEY;
  if (!url || !key) {
    throw new Error('SALES_SUPABASE_URL and SALES_SUPABASE_SERVICE_ROLE_KEY must be set');
  }
  return createClient(url, key, { auth: { persistSession: false } });
}

function trimText(value) {
  return String(value ?? '').trim();
}

export function dedupeCustomerMatches(rows = []) {
  const map = new Map();
  for (const row of rows) {
    const code = trimText(row.customer_code);
    const cnName = trimText(row.customer_cn_name);
    if (!code || !cnName) continue;
    const key = `${code}\0${cnName}`;
    if (!map.has(key)) map.set(key, { code, cnName });
  }
  return [...map.values()];
}

function rankCustomerCnNameMatch(name, query) {
  const lowerName = name.toLowerCase();
  const lowerQuery = query.toLowerCase();
  if (lowerName === lowerQuery) return 0;
  if (lowerName.startsWith(lowerQuery)) return 1;
  if (lowerName.includes(lowerQuery)) return 2;
  return 3;
}

export function sortCustomerMatchesByName(matches, query) {
  const term = trimText(query);
  if (!term) return matches;
  return [...matches].sort((a, b) => {
    const rankDiff = rankCustomerCnNameMatch(a.cnName, term) - rankCustomerCnNameMatch(b.cnName, term);
    if (rankDiff !== 0) return rankDiff;
    return a.cnName.localeCompare(b.cnName, 'zh-Hans');
  });
}

async function loadCustomerRows(supabase, pattern, limit = 200) {
  const { data, error } = await supabase
    .from(DEFAULT_SALES_TABLE)
    .select('customer_code,customer_cn_name')
    .ilike('customer_cn_name', pattern)
    .not('customer_cn_name', 'is', null)
    .neq('customer_cn_name', '')
    .not('customer_code', 'is', null)
    .neq('customer_code', '')
    .limit(limit);

  if (error) throw new Error(error.message);
  return data || [];
}

async function loadCustomerRowsByCode(supabase, pattern, limit = 200) {
  const { data, error } = await supabase
    .from(DEFAULT_SALES_TABLE)
    .select('customer_code,customer_cn_name')
    .ilike('customer_code', pattern)
    .not('customer_cn_name', 'is', null)
    .neq('customer_cn_name', '')
    .not('customer_code', 'is', null)
    .neq('customer_code', '')
    .limit(limit);

  if (error) throw new Error(error.message);
  return data || [];
}

export async function searchCustomersByCode(code, { limit = DEFAULT_LIMIT } = {}) {
  const term = trimText(code);
  if (!term) return [];

  const supabase = getSalesSupabase();
  const rows = await loadCustomerRowsByCode(supabase, `%${term}%`);
  const matches = dedupeCustomerMatches(rows);

  const exact = matches.filter((m) => m.code.toUpperCase() === term.toUpperCase());
  if (exact.length) return exact.slice(0, limit);

  const prefix = matches.filter((m) => m.code.toUpperCase().startsWith(term.toUpperCase()));
  if (prefix.length) return prefix.slice(0, limit);

  return matches.slice(0, limit);
}

export async function searchCustomersByCnName(name, { limit = DEFAULT_LIMIT } = {}) {
  const term = trimText(name);
  if (!term) return [];

  const supabase = getSalesSupabase();
  const prefixRows = await loadCustomerRows(supabase, `${term}%`);
  let rows = prefixRows;

  if (rows.length < limit) {
    const containsRows = await loadCustomerRows(supabase, `%${term}%`);
    rows = [...prefixRows, ...containsRows];
  }

  const matches = sortCustomerMatchesByName(dedupeCustomerMatches(rows), term);
  return matches.slice(0, limit);
}
