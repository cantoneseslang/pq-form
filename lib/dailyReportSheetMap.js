/** 年別・月別生產日報スプレッドシート ID（month: 1=1月 … 12=12月） */
export const DAILY_REPORT_SHEETS_BY_YEAR = {
  2026: {
    1: '1AOR_xp0pQGEIVaw05qvk_Hu6zItqO6DB6sIn-wet9TY',
    2: '1B7FQ16NSR-ilNnS3_GseTSMFv0A5r_b3yVhHJiKOW8E',
    3: '17e4jrh0G0_MWVxJs4S3FxoreXwGRWyYMwO3lDd22LL8',
    4: '1mX6fQhvMYvKN9-vbdiDgAs30J5Oz4jnV-F5ou_tlaME',
    5: '1yK0dj8uvTBitJPy-SMTJLl6HlCk5lSNkgd5w28RuiHg',
    6: '14R8GVayR_Uu6zx-yBVUTTBQib_rpJo22c_nz53oNbWw',
  },
  2025: {
    1: '1-SgRkRObuUA2oUXd5hq7WH-fD3aFZF-SwUh_p2Cg_UU',
    2: '1xmVydoePmvT3Q7SbprFpj3SbkcAK1Pe6SfQfOXfstBg',
    3: '17OgZmGUCAzkGw2t5jr4qspWwzz9825kSA1y_R9WvK3M',
    4: '1Wk6Ugpv3OdP_96OteM98HM2__7XRc-kptnmRrB0WsO8',
    5: '1LgowJi1M2zFhdnM3PrA0TKO_C8YG73nz9lArZqFWpeE',
    6: '1hTKaP0Yy19NU9gQHypVFE4c9hUTGCFJfO_d3gXkjkI4',
    7: '1b7f9M730l1R2T04aPk3D22Ax-kkX4lsLTqwyMTHPKaI',
    8: '1MY9lZj0mF3BHwXJSsSe7Lb2uodIpY1qEKpO2xmLyUhc',
    9: '1vHiMjE7Gw-FigW5ClkT7T5dhr9psoIV4rvinQio0MA0',
    10: '1bh9dfGZ6Ocu23HqrOx7kabsM_HpKxo7qvje-uSvnlso',
    11: '1gs7X7tS7qECBhiolD6S_g_7y_8E9KM0VOlZxCBKKM_o',
    12: '11DOwKl9JB4zThKn6Vp-bX8Vt-y-4p4XmaAncCT8f614',
  },
};

/** @deprecated use DAILY_REPORT_SHEETS_BY_YEAR[2026] */
export const DAILY_REPORT_SHEETS_BY_MONTH = DAILY_REPORT_SHEETS_BY_YEAR[2026];

export const DEFAULT_DAILY_REPORT_YEAR = '2026';
export const DEFAULT_DAILY_REPORT_MONTH = '6';

export function normalizeSheetYear(value) {
  const y = String(value ?? '').trim();
  if (!/^\d{4}$/.test(y)) return '';
  return y;
}

export function normalizeSheetMonth(value) {
  const m = String(value ?? '').trim();
  if (!m) return '';
  const n = parseInt(m, 10);
  if (!Number.isFinite(n) || n < 1 || n > 12) return '';
  return String(n);
}

export function resolveDailyReportSpreadsheetId(month, year) {
  const normalizedMonth = normalizeSheetMonth(month);
  const normalizedYear = normalizeSheetYear(year) || DEFAULT_DAILY_REPORT_YEAR;

  if (normalizedMonth) {
    const envKey = `PQFORM_DAILY_REPORT_SHEET_${normalizedYear}_${normalizedMonth}`;
    const fromEnv = process.env[envKey];
    if (fromEnv) return fromEnv;

    const legacyEnvKey = normalizedYear === DEFAULT_DAILY_REPORT_YEAR
      ? `PQFORM_DAILY_REPORT_SHEET_${normalizedMonth}`
      : '';
    if (legacyEnvKey && process.env[legacyEnvKey]) return process.env[legacyEnvKey];

    const yearMap = DAILY_REPORT_SHEETS_BY_YEAR[normalizedYear];
    if (yearMap?.[normalizedMonth]) return yearMap[normalizedMonth];
  }

  return process.env.PQFORM_DAILY_REPORT_SHEET_ID
    || DAILY_REPORT_SHEETS_BY_YEAR[DEFAULT_DAILY_REPORT_YEAR][DEFAULT_DAILY_REPORT_MONTH];
}

export function resolveDailyReportYearFromIso(isoDate) {
  const m = String(isoDate ?? '').match(/^(\d{4})-/);
  return m ? m[1] : '';
}

export function resolveDailyReportMonthFromIso(isoDate) {
  const m = String(isoDate ?? '').match(/^(\d{4})-(\d{2})-/);
  if (!m) return '';
  return String(parseInt(m[2], 10));
}

export function resolveDailyReportMonth(dailyReport, recordDateIso = '') {
  const fromLink = normalizeSheetMonth(dailyReport?.month);
  if (fromLink) return fromLink;
  const fromDate = normalizeSheetMonth(resolveDailyReportMonthFromIso(recordDateIso));
  if (fromDate) return fromDate;
  return DEFAULT_DAILY_REPORT_MONTH;
}

export function resolveDailyReportYear(dailyReport, recordDateIso = '') {
  const fromLink = normalizeSheetYear(dailyReport?.year);
  if (fromLink) return fromLink;
  const fromDate = normalizeSheetYear(resolveDailyReportYearFromIso(recordDateIso));
  if (fromDate) return fromDate;
  return DEFAULT_DAILY_REPORT_YEAR;
}
