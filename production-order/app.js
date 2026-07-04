(() => {
  const isLocal = ['localhost', '127.0.0.1'].includes(location.hostname);
  const isStaticDev = location.port === '5508';
  const API_BASE = (isLocal && isStaticDev) ? 'http://localhost:5013' : '';
  const STORAGE_KEY = 'production-order-ui-v8';
  const ITEM_COUNT = 6;
  const STOCK_CARD_GAP = 10;

  const FALLBACK_THICKNESS_OPTIONS = ['0.3', '0.4', '0.4D', '0.5', '0.6', '0.8', '0.8A', '1.0', '1.2', '1.5', '3.0'];
  const NOT_FOUND_CODE = '暫時未搵到產品編碼';
  const NOT_FOUND_NAME = '暫時未搵到產品名稱';

  const form = document.getElementById('productionOrderForm');
  const submitBtn = document.getElementById('submitBtn');
  const clearBtn = document.getElementById('clearBtn');
  const formMessage = document.getElementById('formMessage');
  const productResolveHint = document.getElementById('productResolveHint');
  const productMatchPicker = document.getElementById('productMatchPicker');
  const productItemsBody = document.getElementById('productItemsBody');
  const stockCheckItems = document.getElementById('stockCheckItems');

  let thicknessOptionsHtml = '';
  let activeResolveItem = null;
  const resolveTimers = new Map();
  let stockByCode = null;
  let stockAlignFrame = null;
  let stockAlignObserver = null;

  function formatThicknessValue(value) {
    const text = String(value ?? '').trim();
    if (!text) return '';
    if (/[A-Za-z]/.test(text)) return text;
    const num = parseFloat(text);
    return Number.isFinite(num) ? num.toFixed(1) : text;
  }

  function thicknessForProductLookup(value) {
    const t = formatThicknessValue(value);
    if (t === '0.8A') return '0.8';
    if (t === '0.4D') return '0.4';
    return t;
  }

  function normalizeThicknessOptions(items) {
    const seen = new Set();
    const out = [];
    items.forEach((item) => {
      const v = formatThicknessValue(item);
      if (!v || seen.has(v)) return;
      seen.add(v);
      out.push(v);
    });
    return out.sort((a, b) => {
      const na = parseFloat(a);
      const nb = parseFloat(b);
      if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
      return a.localeCompare(b);
    });
  }

  function buildThicknessSelectOptions(options, selected = '') {
    const normalized = normalizeThicknessOptions(options);
    const html = ['<option value="">厚度</option>'];
    normalized.forEach((t) => {
      const sel = t === selected ? ' selected' : '';
      html.push(`<option value="${t.replace(/"/g, '&quot;')}"${sel}>${t}</option>`);
    });
    return html.join('');
  }

  function productTypeOptions(selected = '') {
    const types = ['企筒', '地槽', '鐵角', '批灰角', 'W角', '闊槽', 'C槽'];
    return ['<option value="">請選擇</option>']
      .concat(types.map((t) => `<option value="${t}"${t === selected ? ' selected' : ''}>${t}</option>`))
      .join('');
  }

  function machineOptions(selected = '') {
    const machines = ['1號滾壓成型機', '2號滾壓成型機', '3號滾壓成型機', '4號滾壓成型機', '5號滾壓成型機'];
    return ['<option value="">請選擇</option>']
      .concat(machines.map((m) => `<option value="${m}"${m === selected ? ' selected' : ''}>${m}</option>`))
      .join('');
  }

  function packagingOptions(selected = '') {
    const options = ['(每扎2支)', '(每扎4支)', '(每扎6支)', '(每扎8支)', '(每扎10支)'];
    return ['<option value="">請選擇</option>']
      .concat(options.map((o) => `<option value="${o}"${o === selected ? ' selected' : ''}>${o}</option>`))
      .join('');
  }

  function bandOptions(selected = '') {
    const options = ['(黃帶)', '(粉紅)', '(綠帶)', '(藍帶)'];
    return ['<option value="">請選擇</option>']
      .concat(options.map((o) => `<option value="${o}"${o === selected ? ' selected' : ''}>${o}</option>`))
      .join('');
  }

  function renderProductItems() {
    const thickOpts = thicknessOptionsHtml || buildThicknessSelectOptions(FALLBACK_THICKNESS_OPTIONS);

    const rows = Array.from({ length: ITEM_COUNT }, (_, i) => {
      const n = i + 1;
      const startClass = ' pos-item-line--start';
      const endRowspanClass = n === ITEM_COUNT ? ' pos-item-line--end-rowspan' : '';
      const endClass = n === ITEM_COUNT ? ' pos-item-line--end' : '';
      return `
        <tr class="pos-item-line pos-item-line--1${startClass}${endRowspanClass}" data-item="${n}">
          <td class="pos-item-no" rowspan="5">${n}</td>
          <td class="pos-grid-label">產品編碼：</td>
          <td colspan="2"><input data-field="productCode" class="cell-input cell-input--code" type="text" readonly tabindex="-1" /></td>
          <td class="pos-grid-label">產品名稱：</td>
          <td colspan="2"><input data-field="productName" class="cell-input cell-input--name" type="text" readonly tabindex="-1" /></td>
          <td class="pos-qty" rowspan="5">
            <span class="pos-qty-inner">
              <input data-field="quantity" class="cell-input cell-input--qty" type="text" inputmode="numeric" autocomplete="off" />
              <span class="cell-suffix">支</span>
            </span>
          </td>
        </tr>
        <tr class="pos-item-line pos-item-line--2" data-item="${n}">
          <td colspan="3"><select data-field="productType" class="cell-select cell-select--type">${productTypeOptions()}</select></td>
          <td colspan="3" class="pos-pair-row">
            <div class="pos-pair-inner">
              <div class="pos-pair-cell pos-grid-length">
                <input data-field="length" class="cell-input cell-input--num cell-input--length" type="text" inputmode="decimal" autocomplete="off" />
                <span class="cell-suffix">mm長</span>
              </div>
              <div class="pos-pair-cell">
                <select data-field="machine" class="cell-select cell-select--machine">${machineOptions()}</select>
              </div>
            </div>
          </td>
        </tr>
        <tr class="pos-item-line pos-item-line--3" data-item="${n}">
          <td class="pos-grid-dim" colspan="2">
            <input data-field="width" class="cell-input cell-input--num" type="text" inputmode="decimal" autocomplete="off" />
            <span class="cell-suffix">mm</span>
          </td>
          <td class="pos-grid-x">x</td>
          <td colspan="3" class="pos-pair-row">
            <div class="pos-pair-inner">
              <div class="pos-pair-cell pos-grid-dim">
                <input data-field="height" class="cell-input cell-input--num" type="text" inputmode="decimal" autocomplete="off" />
                <span class="cell-suffix">mm</span>
              </div>
              <div class="pos-pair-cell pos-grid-thick">
                <span class="cell-prefix">(</span>
                <select data-field="thickness" class="cell-select cell-select--thick">${thickOpts}</select>
                <span class="cell-suffix">mm厚)</span>
              </div>
            </div>
          </td>
        </tr>
        <tr class="pos-item-line pos-item-line--4" data-item="${n}">
          <td class="pos-grid-label">用料闊度：</td>
          <td colspan="2" class="pos-grid-dim">
            <input data-field="materialWidth" class="cell-input cell-input--num" type="text" inputmode="decimal" autocomplete="off" />
            <span class="cell-suffix">mm</span>
          </td>
          <td colspan="3" class="pos-pair-row">
            <div class="pos-pair-inner">
              <div class="pos-pair-cell">
                <select data-field="packagingNote" class="cell-select cell-select--pack">${packagingOptions()}</select>
              </div>
              <div class="pos-pair-cell">
                <select data-field="bandNote" class="cell-select cell-select--band">${bandOptions()}</select>
              </div>
            </div>
          </td>
        </tr>
        <tr class="pos-item-line pos-item-line--5 pos-item-line--status${endClass}" data-item="${n}">
          <td colspan="6" class="pos-status-row">
            <div class="pos-status-inner">
              <div class="pos-status-group">
                <span class="pos-grid-label">開料完成：</span>
                <span class="pos-chk"><input type="checkbox" data-field="cuttingComplete" aria-label="開料完成" /></span>
              </div>
              <div class="pos-status-group">
                <span class="pos-grid-label">包裝完成：</span>
                <span class="pos-chk"><input type="checkbox" data-field="packagingComplete" aria-label="包裝完成" /></span>
              </div>
              <div class="pos-status-group">
                <span class="pos-grid-label">打孔完成：</span>
                <span class="pos-chk"><input type="checkbox" data-field="punchingComplete" aria-label="打孔完成" /></span>
              </div>
            </div>
          </td>
        </tr>`;
    }).join('');

    const footer = `
      <tr class="pos-footer-row">
        <td class="pos-grid-label">送：</td>
        <td class="pos-footer-value" colspan="6">
          <input id="deliveryDestination" class="cell-input" type="text" autocomplete="off" />
        </td>
        <td class="pos-footer-qty"></td>
      </tr>
      <tr class="pos-footer-instruction-row">
        <td colspan="8" class="pos-footer-instruction">(起貨後請將此單交回寫字樓)</td>
      </tr>
      <tr class="pos-footer-sign-row">
        <td colspan="8" class="pos-footer-sign-row-cell">
          <div class="pos-footer-sign-inner">
            <div class="pos-footer-sign-label">制單人簽名：</div>
            <div class="pos-footer-sign-value">
              <input id="preparerSignature" class="cell-input" type="text" autocomplete="off" />
            </div>
            <div class="pos-footer-date" id="orderSheetDateDisplay"></div>
          </div>
        </td>
      </tr>`;

    productItemsBody.innerHTML = rows + footer;
    refreshFieldFillStates(productItemsBody);
    renderStockCheckSlots();
    scheduleStockCardAlignment();
  }

  function getItemRow(itemNo) {
    return productItemsBody.querySelector(`tr[data-item="${itemNo}"]`);
  }

  function getItemBlockRows(itemNo) {
    const rows = productItemsBody.querySelectorAll(`tr[data-item="${itemNo}"]`);
    if (!rows.length) return null;
    return { first: rows[0], last: rows[rows.length - 1] };
  }

  function isStockPanelAlignedLayout() {
    return window.matchMedia('(min-width: 901px)').matches;
  }

  function resetStockCardAlignment() {
    if (!stockCheckItems) return;
    stockCheckItems.classList.remove('is-aligned');
    stockCheckItems.style.top = '';
    stockCheckItems.style.left = '';
    stockCheckItems.style.width = '';
    stockCheckItems.style.height = '';
    stockCheckItems.style.marginTop = '';
    stockCheckItems.querySelectorAll('.stock-check-slot').forEach((slot) => {
      slot.style.top = '';
      slot.style.height = '';
    });
  }

  function applyStockSlotPositions() {
    if (!stockCheckItems?.classList.contains('is-aligned')) return;
    const containerRect = stockCheckItems.getBoundingClientRect();
    for (let n = 1; n <= ITEM_COUNT; n += 1) {
      const block = getItemBlockRows(n);
      const slot = stockCheckItems.querySelector(`[data-item="${n}"]`);
      if (!block || !slot) continue;
      const blockTop = block.first.getBoundingClientRect().top;
      const blockBottom = block.last.getBoundingClientRect().bottom;
      const blockHeight = blockBottom - blockTop;
      slot.style.top = `${blockTop - containerRect.top + STOCK_CARD_GAP / 2}px`;
      slot.style.height = `${Math.max(0, blockHeight - STOCK_CARD_GAP)}px`;
    }
  }

  function syncStockCardAlignment() {
    if (!stockCheckItems) return;
    if (!isStockPanelAlignedLayout()) {
      resetStockCardAlignment();
      return;
    }

    const workspace = document.querySelector('.pos-workspace');
    const block1 = getItemBlockRows(1);
    const blockLast = getItemBlockRows(ITEM_COUNT);
    if (!workspace || !block1 || !blockLast) return;

    const wsRect = workspace.getBoundingClientRect();
    const productTop = block1.first.getBoundingClientRect().top;
    const productBottom = blockLast.last.getBoundingClientRect().bottom;

    stockCheckItems.classList.add('is-aligned');
    stockCheckItems.style.top = `${productTop - wsRect.top}px`;
    stockCheckItems.style.height = `${productBottom - productTop}px`;

    applyStockSlotPositions();
    requestAnimationFrame(() => {
      applyStockSlotPositions();
      requestAnimationFrame(applyStockSlotPositions);
    });
  }

  function scheduleStockCardAlignment() {
    if (stockAlignFrame) cancelAnimationFrame(stockAlignFrame);
    stockAlignFrame = requestAnimationFrame(() => {
      stockAlignFrame = null;
      syncStockCardAlignment();
    });
  }

  function bindStockCardAlignment() {
    window.addEventListener('resize', scheduleStockCardAlignment);
    window.addEventListener('load', scheduleStockCardAlignment);
    document.fonts?.ready?.then(scheduleStockCardAlignment);

    const observeTargets = [
      document.querySelector('.a4-sheet'),
      document.querySelector('.pos-product'),
      productItemsBody,
    ].filter(Boolean);

    if (typeof ResizeObserver !== 'undefined') {
      stockAlignObserver = new ResizeObserver(scheduleStockCardAlignment);
      observeTargets.forEach((el) => stockAlignObserver.observe(el));
    }
  }

  function getField(row, name) {
    const itemNo = row?.dataset?.item;
    if (!itemNo) return null;
    return productItemsBody.querySelector(`tr[data-item="${itemNo}"] [data-field="${name}"]`);
  }

  const BAND_TONE_CLASSES = [
    'band-tone--default',
    'band-tone--yellow',
    'band-tone--pink',
    'band-tone--blue',
    'band-tone--green',
  ];

  function getBandToneClass(value) {
    const text = String(value ?? '').trim();
    if (text === '(黃帶)') return 'band-tone--yellow';
    if (text === '(粉紅)') return 'band-tone--pink';
    if (text === '(藍帶)') return 'band-tone--blue';
    if (text === '(綠帶)') return 'band-tone--green';
    return 'band-tone--default';
  }

  function updateBandNoteTone(el) {
    if (!el || el.dataset.field !== 'bandNote') return;
    BAND_TONE_CLASSES.forEach((cls) => el.classList.remove(cls));
    el.classList.add(getBandToneClass(el.value));
  }

  function isFieldFilled(el) {
    if (!el) return false;
    return String(el.value ?? '').trim().length > 0;
  }

  function updateFieldFillState(el) {
    if (!el) return;
    if (!el.matches('.cell-input, .cell-select')) return;
    const filled = isFieldFilled(el);
    el.classList.toggle('cell-is-filled', filled);
    el.classList.toggle('cell-is-empty', !filled);
    updateBandNoteTone(el);
  }

  function refreshFieldFillStates(root = form) {
    root.querySelectorAll('.pos-product .cell-input, .pos-product .cell-select, .pos-info .cell-input').forEach(updateFieldFillState);
  }

  function getSpecValues(row) {
    return {
      type: getField(row, 'productType')?.value.trim() || '',
      thickness: getField(row, 'thickness')?.value.trim() || '',
      width: getField(row, 'width')?.value.trim() || '',
      height: getField(row, 'height')?.value.trim() || '',
      length: getField(row, 'length')?.value.trim() || '',
    };
  }

  function hasAllSpecValues(spec) {
    return spec.type && spec.thickness && spec.width && spec.height && spec.length;
  }

  function buildProvisionalProductName(type, spec) {
    return `${formatThicknessValue(spec.thickness)}x${spec.width}x${spec.height} ${type} ${spec.length}mm`;
  }

  function applyRecordedThicknessToProductName(name, recordedThickness) {
    const displayT = formatThicknessValue(recordedThickness);
    if (!displayT) return name;
    return String(name).replace(/^[\d.]+(?=[x×])/i, displayT);
  }

  function hideProductMatchPicker() {
    productMatchPicker.hidden = true;
    productMatchPicker.innerHTML = '';
    activeResolveItem = null;
  }

  function showProductMatchPicker(itemNo, matches, onPick) {
    activeResolveItem = itemNo;
    productMatchPicker.innerHTML = `<span class="product-match-picker__label">項目 ${itemNo}：</span>`;
    matches.forEach((match) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = `${match.code} · ${match.name}`;
      btn.addEventListener('click', () => {
        onPick(match);
        hideProductMatchPicker();
      });
      productMatchPicker.appendChild(btn);
    });
    productMatchPicker.hidden = false;
  }

  function showProductResolveHint(text, isError = false) {
    productResolveHint.hidden = !text;
    productResolveHint.textContent = text || '';
    productResolveHint.classList.toggle('form-message--error', isError);
    productResolveHint.classList.toggle('form-message--hint', !isError);
  }

  async function ensureStockData() {
    if (stockByCode) return stockByCode;
    try {
      const res = await fetch(`${API_BASE}/api/pq_form/stock`, { cache: 'no-store' });
      const json = await res.json();
      stockByCode = json.success && json.data ? json.data : {};
    } catch (error) {
      console.error('ensureStockData failed', error);
      stockByCode = {};
    }
    return stockByCode;
  }

  function escapeHtml(text) {
    return String(text ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function formatStockQuantity(value, unit) {
    if (value === null || value === undefined || value === '') return '';
    return `${value}${unit || ''}`;
  }

  function buildStockDetailsHtml(stock) {
    const parts = [];
    if (stock.location) parts.push(`📍 ${escapeHtml(stock.location)}`);
    if (stock.onHand !== null && stock.onHand !== undefined) {
      parts.push(`📦 OH ${escapeHtml(formatStockQuantity(stock.onHand, stock.unit))}`);
    }
    if (stock.withoutDn !== null && stock.withoutDn !== undefined) {
      parts.push(`📃 w/o ${escapeHtml(formatStockQuantity(stock.withoutDn, stock.unit))}`);
    }
    if (stock.available !== null && stock.available !== undefined) {
      parts.push(`Avail 📊 ${escapeHtml(formatStockQuantity(stock.available, stock.unit))}`);
    }
    if (stock.category) parts.push(`🏷️ ${escapeHtml(stock.category)}`);
    return parts.join(' | ');
  }

  function renderStockCardSlot(itemNo, state, payload = {}) {
    const slot = stockCheckItems?.querySelector(`[data-item="${itemNo}"]`);
    if (!slot) return;

    const { code = '', name = '', stock = null } = payload;
    let cardHtml = '';

    if (state === 'empty') {
      cardHtml = `
        <div class="stock-card stock-card--empty">
          <div class="stock-card__placeholder">項目 ${itemNo}<br />輸入規格後顯示在庫</div>
        </div>`;
    } else if (state === 'loading') {
      cardHtml = `
        <div class="stock-card stock-card--loading">
          <div class="stock-card__placeholder">項目 ${itemNo}<br />載入中…</div>
        </div>`;
    } else if (state === 'plist-miss') {
      cardHtml = `
        <div class="stock-card stock-card--not-found">
          <div class="stock-card__code">產品編碼 | ${escapeHtml(code || NOT_FOUND_CODE)}</div>
          <div class="stock-card__name">${escapeHtml(name || NOT_FOUND_NAME)}</div>
          <div class="stock-card__details">plist 未登録 — 在庫查詢略過</div>
        </div>`;
    } else if (state === 'not-found') {
      cardHtml = `
        <div class="stock-card stock-card--not-found">
          <div class="stock-card__code">產品編碼 | ${escapeHtml(code)}</div>
          <div class="stock-card__name">${escapeHtml(name)}</div>
          <div class="stock-card__details">在庫資料未找到</div>
        </div>`;
    } else if (state === 'ready' && stock) {
      const negativeClass = Number(stock.available) < 0 ? ' stock-card--negative' : '';
      cardHtml = `
        <div class="stock-card${negativeClass}">
          <div class="stock-card__code">產品編碼 | ${escapeHtml(stock.code || code)}</div>
          <div class="stock-card__name">${escapeHtml(stock.name || name)}</div>
          <div class="stock-card__details">${buildStockDetailsHtml(stock)}</div>
        </div>`;
    }

    slot.innerHTML = cardHtml;
    scheduleStockCardAlignment();
  }

  function renderStockCheckSlots() {
    if (!stockCheckItems) return;
    stockCheckItems.innerHTML = Array.from({ length: ITEM_COUNT }, (_, i) => {
      const n = i + 1;
      return `<div class="stock-check-slot" data-item="${n}"></div>`;
    }).join('');
    for (let i = 1; i <= ITEM_COUNT; i += 1) {
      renderStockCardSlot(i, 'empty');
    }
    scheduleStockCardAlignment();
  }

  async function updateStockCardForItem(itemNo) {
    const row = getItemRow(itemNo);
    if (!row) return;

    const code = getField(row, 'productCode')?.value.trim() || '';
    const name = getField(row, 'productName')?.value.trim() || '';

    if (!code && !name) {
      renderStockCardSlot(itemNo, 'empty');
      return;
    }

    if (!code || code === NOT_FOUND_CODE) {
      renderStockCardSlot(itemNo, 'plist-miss', { code, name });
      return;
    }

    renderStockCardSlot(itemNo, 'loading', { code, name });

    const data = await ensureStockData();
    const stock = data[code.toUpperCase()] || null;
    if (!stock) {
      renderStockCardSlot(itemNo, 'not-found', { code, name });
      return;
    }

    renderStockCardSlot(itemNo, 'ready', { code, name, stock });
  }

  function syncAllStockCards() {
    for (let i = 1; i <= ITEM_COUNT; i += 1) {
      updateStockCardForItem(i);
    }
  }

  function setProductOutputs(row, code, name, notFound = false, materialWidth = null) {
    const codeInput = getField(row, 'productCode');
    const nameInput = getField(row, 'productName');
    const materialInput = getField(row, 'materialWidth');
    if (!codeInput || !nameInput) return;
    codeInput.value = code;
    nameInput.value = name;
    codeInput.classList.toggle('product-not-found', notFound || code === NOT_FOUND_CODE);
    nameInput.classList.toggle('product-not-found', notFound || name === NOT_FOUND_NAME);
    if (materialInput && materialWidth !== null) {
      materialInput.value = materialWidth;
      updateFieldFillState(materialInput);
    }
    updateFieldFillState(codeInput);
    updateFieldFillState(nameInput);
    if (row?.dataset?.item) {
      updateStockCardForItem(Number(row.dataset.item));
    }
  }

  function clearProductOutputs(row) {
    setProductOutputs(row, '', '', false, '');
    if (row?.dataset?.item) {
      renderStockCardSlot(Number(row.dataset.item), 'empty');
    }
  }

  async function tryResolveProduct(itemNo) {
    const row = getItemRow(itemNo);
    if (!row) return;

    const spec = getSpecValues(row);
    const codeInput = getField(row, 'productCode');
    const nameInput = getField(row, 'productName');

    if (!hasAllSpecValues(spec)) {
      clearProductOutputs(row);
      if (activeResolveItem === itemNo) hideProductMatchPicker();
      return;
    }

    if (activeResolveItem === itemNo) hideProductMatchPicker();
    showProductResolveHint('');

    const params = new URLSearchParams({
      type: spec.type,
      t: thicknessForProductLookup(spec.thickness),
      w: spec.width,
      h: spec.height,
      l: spec.length,
    });

    try {
      const res = await fetch(`${API_BASE}/api/pq_form/plist/search?${params.toString()}`, { cache: 'no-store' });
      const data = await res.json();
      const displayName = (name) => applyRecordedThicknessToProductName(name, spec.thickness);

      if (!data.success) {
        setProductOutputs(row, NOT_FOUND_CODE, NOT_FOUND_NAME, true);
        return;
      }

      if (data.matches.length === 1) {
        const match = data.matches[0];
        setProductOutputs(row, match.code, displayName(match.name), false, match.materialWidth || '');
        persistLocal();
      } else if (data.matches.length > 1) {
        codeInput.value = '';
        const uniqueNames = [...new Set(data.matches.map((m) => displayName(m.name)))];
        nameInput.value = uniqueNames.length === 1 ? uniqueNames[0] : '';
        codeInput.classList.remove('product-not-found');
        nameInput.classList.remove('product-not-found');
        updateStockCardForItem(itemNo);
        showProductMatchPicker(itemNo, data.matches, (match) => {
          setProductOutputs(row, match.code, displayName(match.name), false, match.materialWidth || '');
          persistLocal();
        });
      } else {
        setProductOutputs(row, NOT_FOUND_CODE, buildProvisionalProductName(spec.type, spec), true, '');
        if (data.hint) showProductResolveHint(`項目 ${itemNo}：${data.hint}`, true);
        persistLocal();
      }
    } catch (error) {
      console.error('plist search failed', error);
      setProductOutputs(row, NOT_FOUND_CODE, NOT_FOUND_NAME, true);
    }
  }

  function scheduleResolveProduct(itemNo) {
    clearTimeout(resolveTimers.get(itemNo));
    resolveTimers.set(itemNo, setTimeout(() => tryResolveProduct(itemNo), 300));
  }

  function collectProductFromRow(row) {
    return {
      type: getField(row, 'productType')?.value || '',
      machine: getField(row, 'machine')?.value || '',
      thickness: getField(row, 'thickness')?.value || '',
      width: getField(row, 'width')?.value || '',
      height: getField(row, 'height')?.value || '',
      length: getField(row, 'length')?.value || '',
      quantity: getField(row, 'quantity')?.value || '',
      productCode: getField(row, 'productCode')?.value || '',
      productName: getField(row, 'productName')?.value || '',
      materialWidth: getField(row, 'materialWidth')?.value || '',
      packagingNote: getField(row, 'packagingNote')?.value || '',
      bandNote: getField(row, 'bandNote')?.value || '',
      cuttingComplete: !!getField(row, 'cuttingComplete')?.checked,
      packagingComplete: !!getField(row, 'packagingComplete')?.checked,
      punchingComplete: !!getField(row, 'punchingComplete')?.checked,
    };
  }

  function collectFormState() {
    return {
      header: {
        deliveryNoteNo: buildDeliveryNoteNo(document.getElementById('deliveryNoteNoSuffix')?.value),
        customerNo: document.getElementById('customerNo').value,
        orderingCompany: document.getElementById('orderingCompany').value,
        deliveryDate: document.getElementById('deliveryDate').value,
        orderDate: document.getElementById('orderDate').value,
        estimatedProductionPeriod: document.getElementById('estimatedProductionPeriod').value,
        completionDate: document.getElementById('completionDate').value,
        personInCharge: document.getElementById('personInCharge').value,
        signature: document.getElementById('signature').value,
        deliveryDestination: document.getElementById('deliveryDestination')?.value || '',
        preparerSignature: document.getElementById('preparerSignature')?.value || '',
      },
      products: Array.from({ length: ITEM_COUNT }, (_, i) => collectProductFromRow(getItemRow(i + 1))),
    };
  }

  function applyProductToRow(row, product = {}) {
    if (!row) return;
    getField(row, 'productType').value = product.type || '';
    getField(row, 'machine').value = product.machine || '';
    getField(row, 'thickness').value = product.thickness || '';
    getField(row, 'width').value = product.width || '';
    getField(row, 'height').value = product.height || '';
    getField(row, 'length').value = product.length || '';
    getField(row, 'quantity').value = product.quantity || '';
    getField(row, 'productCode').value = product.productCode || '';
    getField(row, 'productName').value = product.productName || '';
    getField(row, 'materialWidth').value = product.materialWidth || '';
    getField(row, 'packagingNote').value = product.packagingNote || '';
    getField(row, 'bandNote').value = product.bandNote || '';
    getField(row, 'cuttingComplete').checked = !!product.cuttingComplete;
    getField(row, 'packagingComplete').checked = !!product.packagingComplete;
    getField(row, 'punchingComplete').checked = !!product.punchingComplete;
    productItemsBody.querySelectorAll(`tr[data-item="${row.dataset.item}"] [data-field]`).forEach((el) => {
      if (el.matches('.cell-input, .cell-select')) updateFieldFillState(el);
    });
  }

  function applyFormState(state) {
    if (!state) return;
    const header = state.header || {};
    const parsedNoteNo = parseDeliveryNoteNo(header.deliveryNoteNo);
    const suffixInput = document.getElementById('deliveryNoteNoSuffix');
    if (suffixInput) suffixInput.value = parsedNoteNo.suffix;
    updateDeliveryNoteNoPrefix(parsedNoteNo.yyyyMM);
    document.getElementById('customerNo').value = header.customerNo || '';
    document.getElementById('orderingCompany').value = header.orderingCompany || '';
    document.getElementById('deliveryDate').value = header.deliveryDate || '';
    document.getElementById('orderDate').value = header.orderDate || '';
    document.getElementById('estimatedProductionPeriod').value = header.estimatedProductionPeriod || '';
    document.getElementById('completionDate').value = header.completionDate || '';
    document.getElementById('personInCharge').value = header.personInCharge || '';
    document.getElementById('signature').value = header.signature || '';
    const dest = document.getElementById('deliveryDestination');
    if (dest) dest.value = header.deliveryDestination || '';
    const preparer = document.getElementById('preparerSignature');
    if (preparer) preparer.value = header.preparerSignature || '';
    updateOrderSheetDateDisplay();

    const products = state.products || (state.product ? [state.product] : []);
    for (let i = 0; i < ITEM_COUNT; i += 1) {
      applyProductToRow(getItemRow(i + 1), products[i] || {});
    }
    syncAllStockCards();
  }

  async function loadThicknessOptions() {
    let options = FALLBACK_THICKNESS_OPTIONS;
    try {
      const res = await fetch(`${API_BASE}/api/pq_form/plist/thicknesses`, { cache: 'no-store' });
      const data = await res.json();
      if (data.success && Array.isArray(data.thicknesses)) {
        options = normalizeThicknessOptions([...data.thicknesses, ...FALLBACK_THICKNESS_OPTIONS]);
      }
    } catch (error) {
      console.error('loadThicknessOptions failed', error);
    }
    thicknessOptionsHtml = buildThicknessSelectOptions(options);
    productItemsBody.querySelectorAll('[data-field="thickness"]').forEach((sel) => {
      const current = sel.value;
      sel.innerHTML = buildThicknessSelectOptions(options, current);
    });
  }

  function persistLocal() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(collectFormState()));
    } catch (error) {
      console.warn('persistLocal failed', error);
    }
  }

  function restoreLocal() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY)
        || localStorage.getItem('production-order-ui-v7')
        || localStorage.getItem('production-order-ui-v2')
        || localStorage.getItem('production-order-ui-v1');
      if (!raw) return;
      applyFormState(JSON.parse(raw));
    } catch (error) {
      console.warn('restoreLocal failed', error);
    }
  }

  function showFormMessage(text, type = 'hint') {
    formMessage.hidden = !text;
    formMessage.textContent = text || '';
    formMessage.classList.remove('form-message--error', 'form-message--success', 'form-message--hint');
    formMessage.classList.add(`form-message--${type}`);
  }

  function parseFormDate(value) {
    const text = String(value ?? '').trim();
    if (!text) return null;
    const match = text.match(/^(\d{4})[\/\-年.](\d{1,2})[\/\-月.](\d{1,2})/);
    if (!match) return null;
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    return { year, month, day };
  }

  function formatDeliveryNoteYearMonth(dateValue) {
    const parsed = parseFormDate(dateValue);
    if (!parsed) return '';
    return `${parsed.year}${String(parsed.month).padStart(2, '0')}`;
  }

  function getDeliveryNoteDateSource() {
    const deliveryDate = document.getElementById('deliveryDate')?.value.trim() || '';
    const orderDate = document.getElementById('orderDate')?.value.trim() || '';
    return deliveryDate || orderDate;
  }

  function parseDeliveryNoteNo(value) {
    const text = String(value ?? '').trim();
    const match = text.match(/^SC\/(\d{6})\/(\d{1,3})$/i);
    if (!match) return { yyyyMM: '', suffix: '' };
    return { yyyyMM: match[1], suffix: match[2] };
  }

  function updateDeliveryNoteNoPrefix(fallbackYyyyMM = '') {
    const prefixEl = document.getElementById('deliveryNoteNoPrefix');
    if (!prefixEl) return;
    const yyyyMM = formatDeliveryNoteYearMonth(getDeliveryNoteDateSource()) || fallbackYyyyMM;
    prefixEl.textContent = yyyyMM ? `SC/${yyyyMM}/` : 'SC/';
  }

  function buildDeliveryNoteNo(suffix) {
    const yyyyMM = formatDeliveryNoteYearMonth(getDeliveryNoteDateSource());
    const tail = String(suffix ?? '').replace(/\D/g, '').slice(0, 3);
    if (!yyyyMM || !tail) return '';
    return `SC/${yyyyMM}/${tail}`;
  }

  function normalizeDeliveryNoteSuffixInput() {
    const input = document.getElementById('deliveryNoteNoSuffix');
    if (!input) return;
    input.value = input.value.replace(/\D/g, '').slice(0, 3);
  }

  function formatDisplayDateChinese(dateValue) {
    const parsed = parseFormDate(dateValue);
    if (!parsed) return '';
    return `${parsed.year}年${parsed.month}月${parsed.day}日`;
  }

  function updateOrderSheetDateDisplay() {
    const el = document.getElementById('orderSheetDateDisplay');
    if (!el) return;
    el.textContent = formatDisplayDateChinese(getDeliveryNoteDateSource());
  }

  function setDefaultOrderDate() {
    const orderDate = document.getElementById('orderDate');
    if (orderDate.value) return;
    const now = new Date();
    orderDate.value = `${now.getFullYear()}/${now.getMonth() + 1}/${now.getDate()}`;
  }

  function bindEvents() {
    const specFields = ['productType', 'thickness', 'width', 'height', 'length'];

    productItemsBody.addEventListener('input', (e) => {
      const row = e.target.closest('tr[data-item]');
      updateFieldFillState(e.target);
      if (row) {
        persistLocal();
        if (specFields.includes(e.target.dataset.field)) {
          scheduleResolveProduct(Number(row.dataset.item));
        }
        return;
      }
      if (e.target.id === 'deliveryDestination') {
        updateFieldFillState(e.target);
        persistLocal();
      }
    });

    productItemsBody.addEventListener('change', (e) => {
      updateFieldFillState(e.target);
      const row = e.target.closest('tr[data-item]');
      if (row || e.target.type === 'checkbox') persistLocal();
      if (!row) return;
      if (specFields.includes(e.target.dataset.field)) {
        scheduleResolveProduct(Number(row.dataset.item));
      }
    });

    form.querySelectorAll('#deliveryNoteNoSuffix, #customerNo, #orderingCompany, #deliveryDate, #orderDate, #estimatedProductionPeriod, #completionDate, #personInCharge, #signature, #preparerSignature').forEach((el) => {
      const handler = () => {
        if (el.id === 'deliveryNoteNoSuffix') normalizeDeliveryNoteSuffixInput();
        if (el.id === 'deliveryDate' || el.id === 'orderDate') {
          updateDeliveryNoteNoPrefix();
          updateOrderSheetDateDisplay();
        }
        updateFieldFillState(el);
        persistLocal();
      };
      el.addEventListener('input', handler);
      el.addEventListener('change', handler);
    });

    clearBtn.addEventListener('click', () => {
      form.reset();
      for (let i = 1; i <= ITEM_COUNT; i += 1) {
        clearProductOutputs(getItemRow(i));
        renderStockCardSlot(i, 'empty');
      }
      hideProductMatchPicker();
      showProductResolveHint('');
      setDefaultOrderDate();
      updateDeliveryNoteNoPrefix();
      updateOrderSheetDateDisplay();
      refreshFieldFillStates();
      ['production-order-ui-v3', 'production-order-ui-v2', 'production-order-ui-v1'].forEach((k) => localStorage.removeItem(k));
      showFormMessage('');
    });

    form.addEventListener('submit', async (event) => {
      event.preventDefault();
      showFormMessage('');
      submitBtn.disabled = true;
      submitBtn.textContent = '送出中…';

      try {
        const res = await fetch(`${API_BASE}/api/pq_form/production_order/submit`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(collectFormState()),
        });
        const data = await res.json();
        if (!data.success) {
          showFormMessage(data.error || '送出失敗', 'error');
          return;
        }
        showFormMessage(`已成功寫入 Google Sheet（${data.updatedCells} 個儲存格）`, 'success');
        persistLocal();
      } catch (error) {
        console.error('submit failed', error);
        showFormMessage('送出失敗，請稍後再試', 'error');
      } finally {
        submitBtn.disabled = false;
        submitBtn.textContent = '送出至 Google Sheet';
      }
    });
  }

  async function initApp() {
    renderProductItems();
    await loadThicknessOptions();
    restoreLocal();
    setDefaultOrderDate();
    updateDeliveryNoteNoPrefix();
    updateOrderSheetDateDisplay();
    refreshFieldFillStates();
    bindEvents();
    bindStockCardAlignment();
    for (let i = 1; i <= ITEM_COUNT; i += 1) scheduleResolveProduct(i);
    syncAllStockCards();
    scheduleStockCardAlignment();
  }

  initApp();
})();
