(() => {
  const isLocal = ['localhost', '127.0.0.1'].includes(location.hostname);
  const isStaticDev = location.port === '5508';
  const PROD_API = 'https://pq-form.vercel.app';
  const API_BASE = (isLocal && isStaticDev) ? 'http://localhost:5013' : '';

  const FALLBACK_SOURCE_SHEET_ID = '1R-xjzmki0pzMlXJzhzVee_y6XWqFbrUSlSL04osBhbc';

  const totalLotsEl = document.getElementById('totalLots');
  const totalKgEl = document.getElementById('totalKg');
  const ageKg365El = document.getElementById('ageKg365');
  const ageKg730El = document.getElementById('ageKg730');
  const ageKg731El = document.getElementById('ageKg731');
  const fetchedDateEl = document.getElementById('fetchedDate');
  const fetchedTimeEl = document.getElementById('fetchedTime');
  const sourceSheetIdEl = document.getElementById('sourceSheetId');
  const stockBody = document.getElementById('stockBody');
  const loadingMsg = document.getElementById('loadingMsg');
  const errorMsg = document.getElementById('errorMsg');
  const printBtn = document.getElementById('printBtn');

  function formatInteger(value) {
    const n = Number(value);
    if (!Number.isFinite(n)) return '—';
    return Math.round(n).toLocaleString('en-US');
  }

  function formatFetchedTimestamp(iso) {
    if (!iso) return { date: '—', time: '—' };
    const dt = new Date(iso);
    if (Number.isNaN(dt.getTime())) return { date: '—', time: '—' };

    const parts = new Intl.DateTimeFormat('en-GB', {
      timeZone: 'Asia/Hong_Kong',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false,
    }).formatToParts(dt);

    const pick = (type) => parts.find((p) => p.type === type)?.value ?? '';
    return {
      date: `${pick('day')}/${pick('month')}/${pick('year')}`,
      time: `${pick('hour')}:${pick('minute')}:${pick('second')}`,
    };
  }

  function setReportMeta(data) {
    const { date, time } = formatFetchedTimestamp(data?.fetchedAt);
    fetchedDateEl.textContent = date;
    fetchedTimeEl.textContent = time;
    totalLotsEl.textContent = formatInteger(data?.stats?.totalLots ?? data?.lotCount ?? 0);
    totalKgEl.textContent = formatInteger(data?.stats?.totalKg ?? 0);
    ageKg365El.textContent = formatInteger(data?.ageStats?.within365Kg ?? 0);
    ageKg730El.textContent = formatInteger(data?.ageStats?.within730Kg ?? 0);
    ageKg731El.textContent = formatInteger(data?.ageStats?.over730Kg ?? 0);
    sourceSheetIdEl.textContent = data?.sourceSpreadsheetId || FALLBACK_SOURCE_SHEET_ID;
  }

  async function fetchTakePayload() {
    const urls = [`${API_BASE}/api/pq_form/material_stock/take`];
    if (isLocal && API_BASE !== PROD_API) {
      urls.push(`${PROD_API}/api/pq_form/material_stock/take`);
    }

    let lastError = null;
    for (const url of urls) {
      try {
        const res = await fetch(url, { cache: 'default' });
        const json = await res.json();
        if (json.success) return json;
        lastError = new Error(json.error || '材料在庫 take API 錯誤');
      } catch (err) {
        lastError = err;
      }
    }
    throw lastError || new Error('材料在庫 take API 錯誤');
  }

  async function init() {
    try {
      const data = await fetchTakePayload();
      loadingMsg.hidden = true;
      errorMsg.hidden = true;

      if (!data.rowsHtml) {
        errorMsg.textContent = '沒有原材料庫存資料。';
        errorMsg.hidden = false;
        return;
      }

      setReportMeta(data);
      stockBody.innerHTML = data.rowsHtml;
    } catch (err) {
      loadingMsg.hidden = true;
      errorMsg.textContent = `載入失敗: ${err?.message || String(err)}`;
      errorMsg.hidden = false;
    }
  }

  printBtn.addEventListener('click', () => window.print());

  init();
})();
