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

function compactCnName(value) {
  return trimText(value).replace(/\s+/g, '');
}

function buildCnNameIlikePattern(term, { prefix = true } = {}) {
  const chars = [...compactCnName(term)];
  if (!chars.length) return null;
  const core = chars.join('%');
  return prefix ? `${core}%` : `%${core}%`;
}

function rankCustomerCnNameMatch(name, query) {
  const compactName = compactCnName(name);
  const compactQuery = compactCnName(query);
  if (!compactQuery) return 3;
  if (compactName === compactQuery) return 0;
  if (compactName.startsWith(compactQuery)) return 1;
  if (compactName.includes(compactQuery)) return 2;
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
  const prefixPattern = buildCnNameIlikePattern(term, { prefix: true });
  const prefixRows = prefixPattern ? await loadCustomerRows(supabase, prefixPattern) : [];
  let rows = prefixRows;

  if (rows.length < limit) {
    const containsPattern = buildCnNameIlikePattern(term, { prefix: false });
    if (containsPattern && containsPattern !== prefixPattern) {
      const containsRows = await loadCustomerRows(supabase, containsPattern);
      rows = [...prefixRows, ...containsRows];
    }
  }

  const matches = sortCustomerMatchesByName(dedupeCustomerMatches(rows), term);
  return matches.slice(0, limit);
}
