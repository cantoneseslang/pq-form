(() => {
  const isLocal = ['localhost', '127.0.0.1'].includes(location.hostname);
  const isStaticDev = location.port === '5508';
  const API_BASE = (isLocal && isStaticDev) ? 'http://localhost:5013' : '';
  const STORAGE_KEY = 'production-order-ui-v8';
  const ITEM_COUNT = 6;
  const STOCK_CARD_GAP = 10;

  const FALLBACK_THICKNESS_OPTIONS = ['0.3', '0.4', '0.4D', '0.5', '0.6', '0.8', '0.8A', '1.0', '1.0A', '1.2', '1.5', '1.5A', '3.0'];
  const NOT_FOUND_CODE = '暫時未搵到產品編碼';
  const NOT_FOUND_NAME = '暫時未搵到產品名稱';

  const form = document.getElementById('productionOrderForm');
  const submitBtn = document.getElementById('submitBtn');
  const clearBtn = document.getElementById('clearBtn');
  const printBtn = document.getElementById('printBtn');
  const formMessage = document.getElementById('formMessage');
  const productResolveHint = document.getElementById('productResolveHint');
  const productMatchPicker = document.getElementById('productMatchPicker');
  const productItemsBody = document.getElementById('productItemsBody');
  const stockCheckItems = document.getElementById('stockCheckItems');
  const materialCheckItems = document.getElementById('materialCheckItems');
  const customerNoInput = document.getElementById('customerNo');
  const orderingCompanyInput = document.getElementById('orderingCompany');
  const customerCnNameSuggest = document.getElementById('customerCnNameSuggest');
  const customerMatchPicker = document.getElementById('customerMatchPicker');

  let thicknessOptionsHtml = '';
  let activeResolveItem = null;
  const resolveTimers = new Map();
  const materialStockTimers = new Map();
  const materialProducibleByItem = new Map();
  let stockByCode = null;
  let materialStockData = null;
  let materialStockLoadPromise = null;
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
    if (!t) return '';
    const m = t.toUpperCase().match(/^([\d.]+)([AD])$/);
    if (m) return formatThicknessValue(m[1]);
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
    renderMaterialCheckSlots();
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
    [stockCheckItems, materialCheckItems].forEach((container) => {
      if (!container) return;
      container.classList.remove('is-aligned');
      container.style.top = '';
      container.style.left = '';
      container.style.width = '';
      container.style.height = '';
      container.style.marginTop = '';
      container.querySelectorAll('[data-item]').forEach((slot) => {
        slot.style.top = '';
        slot.style.height = '';
      });
    });
  }

  function applySlotPositions(container, slotClass) {
    if (!container?.classList.contains('is-aligned')) return;
    const containerRect = container.getBoundingClientRect();
    for (let n = 1; n <= ITEM_COUNT; n += 1) {
      const block = getItemBlockRows(n);
      const slot = container.querySelector(`${slotClass}[data-item="${n}"]`);
      if (!block || !slot) continue;
      const blockTop = block.first.getBoundingClientRect().top;
      const blockBottom = block.last.getBoundingClientRect().bottom;
      const blockHeight = blockBottom - blockTop;
      slot.style.top = `${blockTop - containerRect.top + STOCK_CARD_GAP / 2}px`;
      slot.style.height = `${Math.max(0, blockHeight - STOCK_CARD_GAP)}px`;
    }
  }

  function applyStockSlotPositions() {
    applySlotPositions(stockCheckItems, '.stock-check-slot');
    applySlotPositions(materialCheckItems, '.material-check-slot');
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
    const alignStyle = {
      top: `${productTop - wsRect.top}px`,
      height: `${productBottom - productTop}px`,
    };

    if (stockCheckItems) {
      stockCheckItems.classList.add('is-aligned');
      Object.assign(stockCheckItems.style, alignStyle);
    }
    if (materialCheckItems) {
      materialCheckItems.classList.add('is-aligned');
      Object.assign(materialCheckItems.style, alignStyle);
    }

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

  function displayPlistProductName(name, spec) {
    return applyRecordedThicknessToProductName(name, spec.thickness);
  }

  function applyPlistMaterialWidthIfEmpty(row, materialWidth) {
    const materialInput = getField(row, 'materialWidth');
    const mw = String(materialWidth ?? '').trim();
    if (!materialInput || !mw || materialInput.value.trim()) return;
    materialInput.value = mw;
    updateFieldFillState(materialInput);
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

  let customerSyncLock = false;
  let customerCodeLookupTimer = null;
  let customerNameLookupTimer = null;
  let customerNameBlurTimer = null;
  let activeCustomerLookupSource = null;

  function hideCustomerCnNameSuggest() {
    if (!customerCnNameSuggest) return;
    customerCnNameSuggest.hidden = true;
    customerCnNameSuggest.innerHTML = '';
  }

  function hideCustomerMatchPicker() {
    if (!customerMatchPicker) return;
    customerMatchPicker.hidden = true;
    customerMatchPicker.innerHTML = '';
  }

  function setCustomerFields({ code, cnName } = {}, { persist = true } = {}) {
    if (!customerNoInput || !orderingCompanyInput) return;
    customerSyncLock = true;
    try {
      if (code !== undefined) customerNoInput.value = code;
      if (cnName !== undefined) orderingCompanyInput.value = cnName;
      updateFieldFillState(customerNoInput);
      updateFieldFillState(orderingCompanyInput);
      if (persist) persistLocal();
    } finally {
      customerSyncLock = false;
    }
  }

  function showCustomerMatchPicker(label, matches, onPick) {
    if (!customerMatchPicker) return;
    customerMatchPicker.innerHTML = `<span class="customer-match-picker__label">${label}</span>`;
    matches.forEach((match) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = `${match.code} · ${match.cnName}`;
      btn.addEventListener('click', () => {
        onPick(match);
        hideCustomerMatchPicker();
      });
      customerMatchPicker.appendChild(btn);
    });
    customerMatchPicker.hidden = false;
  }

  function renderCustomerCnNameSuggest(matches, onPick) {
    if (!customerCnNameSuggest) return;
    customerCnNameSuggest.innerHTML = '';
    if (!matches.length) {
      hideCustomerCnNameSuggest();
      return;
    }
    matches.forEach((match) => {
      const li = document.createElement('li');
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = `${match.cnName} (${match.code})`;
      btn.addEventListener('mousedown', (event) => {
        event.preventDefault();
        onPick(match);
      });
      li.appendChild(btn);
      customerCnNameSuggest.appendChild(li);
    });
    customerCnNameSuggest.hidden = false;
  }

  async function fetchCustomerMatches(params) {
    const qs = new URLSearchParams(params);
    const res = await fetch(`${API_BASE}/api/pq_form/customers/search?${qs.toString()}`, { cache: 'no-store' });
    const data = await res.json();
    if (!data.success) return { ok: false, matches: [], error: data.error || 'lookup failed' };
    return { ok: true, matches: data.matches || [] };
  }

  async function resolveCustomerByCode() {
    if (customerSyncLock || !customerNoInput) return;
    const code = customerNoInput.value.trim();
    hideCustomerMatchPicker();
    if (!code) return;

    activeCustomerLookupSource = 'code';
    try {
      const { ok, matches } = await fetchCustomerMatches({ code });
      if (!ok || activeCustomerLookupSource !== 'code') return;

      if (matches.length === 1) {
        setCustomerFields({ cnName: matches[0].cnName });
        hideCustomerCnNameSuggest();
        return;
      }
      if (matches.length > 1) {
        showCustomerMatchPicker('客戶編號候選：', matches, (match) => {
          setCustomerFields({ code: match.code, cnName: match.cnName });
          hideCustomerCnNameSuggest();
        });
      }
    } catch (error) {
      console.error('customer code lookup failed', error);
    }
  }

  async function resolveCustomerByCnName() {
    if (customerSyncLock || !orderingCompanyInput) return;
    const name = orderingCompanyInput.value.trim();
    hideCustomerMatchPicker();
    hideCustomerCnNameSuggest();
    if (!name) return;

    activeCustomerLookupSource = 'name';
    try {
      const { ok, matches } = await fetchCustomerMatches({ name });
      if (!ok || activeCustomerLookupSource !== 'name') return;

      if (!matches.length) return;

      showCustomerMatchPicker('訂貨公司候選：', matches, (match) => {
        setCustomerFields({ code: match.code, cnName: match.cnName });
        hideCustomerMatchPicker();
      });
    } catch (error) {
      console.error('customer name lookup failed', error);
      hideCustomerMatchPicker();
    }
  }

  function scheduleCustomerCodeLookup() {
    if (customerSyncLock) return;
    clearTimeout(customerCodeLookupTimer);
    customerCodeLookupTimer = setTimeout(() => resolveCustomerByCode(), 300);
  }

  function scheduleCustomerNameLookup() {
    if (customerSyncLock) return;
    clearTimeout(customerNameLookupTimer);
    customerNameLookupTimer = setTimeout(() => resolveCustomerByCnName(), 200);
  }

  let printBlankSelectRestore = [];

  function blankEmptySelectsForPrint() {
    printBlankSelectRestore = [];
    document.querySelectorAll('.pos-product .cell-select').forEach((sel) => {
      if (String(sel.value ?? '').trim()) return;
      const option = sel.options[sel.selectedIndex];
      if (!option) return;
      printBlankSelectRestore.push({ option, text: option.textContent });
      option.textContent = '';
    });
  }

  function restoreEmptySelectsAfterPrint() {
    printBlankSelectRestore.forEach(({ option, text }) => {
      option.textContent = text;
    });
    printBlankSelectRestore = [];
  }

  function bindPrintHandlers() {
    window.addEventListener('beforeprint', blankEmptySelectsForPrint);
    window.addEventListener('afterprint', restoreEmptySelectsAfterPrint);
  }

    if (!customerNoInput || !orderingCompanyInput) return;

    customerNoInput.addEventListener('input', () => {
      if (customerSyncLock) return;
      activeCustomerLookupSource = 'code';
      hideCustomerCnNameSuggest();
      scheduleCustomerCodeLookup();
    });

    orderingCompanyInput.addEventListener('input', () => {
      if (customerSyncLock) return;
      activeCustomerLookupSource = 'name';
      hideCustomerMatchPicker();
      scheduleCustomerNameLookup();
    });

    orderingCompanyInput.addEventListener('focus', () => {
      if (customerSyncLock) return;
      if (orderingCompanyInput.value.trim()) scheduleCustomerNameLookup();
    });

    orderingCompanyInput.addEventListener('blur', () => {
      clearTimeout(customerNameBlurTimer);
      customerNameBlurTimer = setTimeout(() => hideCustomerCnNameSuggest(), 150);
    });
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

  function buildStockDetailsHtml(stock, itemNo) {
    const parts = [];
    if (stock.onHand !== null && stock.onHand !== undefined) {
      parts.push(`OH ${escapeHtml(formatStockQuantity(stock.onHand, stock.unit))}`);
    }
    if (stock.withoutDn !== null && stock.withoutDn !== undefined) {
      parts.push(`w/o ${escapeHtml(formatStockQuantity(stock.withoutDn, stock.unit))}`);
    }
    if (stock.available !== null && stock.available !== undefined) {
      const availText = `Avail ${escapeHtml(formatStockQuantity(stock.available, stock.unit))}`;
      const { status, hasComparison } = getQtyHighlightForItem(itemNo);
      const availClass = hasComparison
        ? (status === 'ok' ? 'stock-card__avail--ok'
          : status === 'partial' ? 'stock-card__avail--partial'
            : status === 'short' ? 'stock-card__avail--short' : '')
        : '';
      parts.push(availClass
        ? `<span class="stock-card__avail ${availClass}">${availText}</span>`
        : availText);
    }
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
          <div class="stock-card__details">${buildStockDetailsHtml(stock, itemNo)}</div>
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
      refreshQtyHighlight(itemNo);
      return;
    }

    if (!code || code === NOT_FOUND_CODE) {
      renderStockCardSlot(itemNo, 'plist-miss', { code, name });
      refreshQtyHighlight(itemNo);
      return;
    }

    renderStockCardSlot(itemNo, 'loading', { code, name });

    const data = await ensureStockData();
    const stock = data[code.toUpperCase()] || null;
    if (!stock) {
      renderStockCardSlot(itemNo, 'not-found', { code, name });
      refreshQtyHighlight(itemNo);
      return;
    }

    renderStockCardSlot(itemNo, 'ready', { code, name, stock });
    refreshQtyHighlight(itemNo);
  }

  function syncAllMaterialStockCards() {
    for (let i = 1; i <= ITEM_COUNT; i += 1) {
      scheduleMaterialStockUpdate(i);
    }
  }

  function formatNumber(value, digits = 0) {
    const num = Number(value);
    if (!Number.isFinite(num)) return '';
    return num.toLocaleString('en-HK', {
      minimumFractionDigits: digits,
      maximumFractionDigits: digits,
    });
  }

  const MATERIAL_STEEL_DENSITY = 7.85;

  const THICKNESS_TAB_CANDIDATES = {
    '0.8A': ['0.8-A mm'],
    '0.4D': ['0.4mm (B)'],
    '0.4B': ['0.4mm (B)'],
    '0.4W': ['0.4mm (W)'],
    '0.4AL': ['0.4mm (Aluminium)'],
    '0.8C': ['0.8mm C (Z-120)'],
    '0.3': ['0.3mm'],
    '0.4': ['0.4mm'],
    '0.45': ['0.45mm'],
    '0.5': ['0.5mm'],
    '0.6': ['0.6mm'],
    '0.8': ['0.8mm'],
    '1.0': ['1.0mm'],
    '1.2': ['1.2mm'],
    '1.5': ['1.5mm', '1.5mm ASTM (G90)'],
    '3.0': ['3.0mm'],
  };

  function normalizeMaterialWidthValue(value) {
    const text = String(value ?? '').trim();
    if (!text) return '';
    const num = parseFloat(text.replace(/mm$/i, ''));
    return Number.isFinite(num) ? String(num) : text;
  }

  function isAluminiumCornerBead({ productType = '', productName = '' } = {}) {
    const type = String(productType).trim();
    if (type !== '批灰角') return false;
    const name = String(productName).trim();
    if (!name) return false;
    return /鋁|aluminium|aluminum/i.test(name)
      || /批灰角鋁|批灰角\(鋁\)|批灰角（鋁）/i.test(name);
  }

  function resolveMaterialInventoryThicknessKey({
    thickness = '',
    productType = '',
    productName = '',
  } = {}) {
    const raw = String(thickness ?? '').trim();
    if (!raw) return '';

    const upper = raw.toUpperCase();
    const aluminium = isAluminiumCornerBead({ productType, productName });

    if (upper === '0.4AL') return '0.4AL';
    if (upper === '0.4D' || upper === '0.4B') return aluminium ? '0.4AL' : '0.4D';
    if (upper === '0.4W' || upper === '0.8A' || upper === '0.8C') return upper;

    if (aluminium && (upper === '0.4' || parseFloat(upper) === 0.4)) return '0.4AL';

    return raw;
  }

  function thicknessTabCandidates(thickness) {
    const key = String(thickness ?? '').trim().toUpperCase();
    if (THICKNESS_TAB_CANDIDATES[key]) return THICKNESS_TAB_CANDIDATES[key];
    const num = parseFloat(key);
    return Number.isFinite(num) ? [`${num}mm`] : [];
  }

  function parsePositiveNumber(value) {
    const text = String(value ?? '').replace(/,/g, '').trim();
    if (!text) return null;
    const num = parseFloat(text);
    return Number.isFinite(num) && num > 0 ? num : null;
  }

  function calcPieceWeightKg({ lengthMm, materialWidthMm, thicknessMm, densityGcm3 = MATERIAL_STEEL_DENSITY }) {
    const length = parsePositiveNumber(lengthMm);
    const width = parsePositiveNumber(materialWidthMm);
    const thickness = parsePositiveNumber(thicknessMm);
    if (!length || !width || !thickness) return null;
    return (length * width * thickness * densityGcm3) / 1_000_000;
  }

  function calcProduciblePieces({ availableKg, lengthMm, materialWidthMm, thicknessMm, densityGcm3 = MATERIAL_STEEL_DENSITY }) {
    const kg = parsePositiveNumber(availableKg);
    const pieceKg = calcPieceWeightKg({ lengthMm, materialWidthMm, thicknessMm, densityGcm3 });
    if (!kg || !pieceKg) return null;
    return Math.floor(kg / pieceKg);
  }

  function densityFromMaterialStock(stock) {
    const tab = stock?.tabTitles?.[0] || '';
    if (/aluminium/i.test(tab)) return 2.7;
    const lotDensity = stock?.lots?.[0]?.densityGcm3;
    if (Number.isFinite(lotDensity) && lotDensity > 0) return lotDensity;
    return MATERIAL_STEEL_DENSITY;
  }

  function materialStockDisplayLabel({ thickness, materialWidth, stock, productType = '', productName = '' }) {
    const tab = stock?.tabTitles?.[0];
    const widthPart = materialWidth ? `${materialWidth}mm` : '';
    if (tab && widthPart) return `${tab} × ${widthPart}`;
    if (tab) return tab;
    const thicknessKey = resolveMaterialInventoryThicknessKey({ thickness, productType, productName });
    const candidateTab = thicknessTabCandidates(thicknessKey)[0];
    if (candidateTab && widthPart) return `${candidateTab} × ${widthPart}`;
    if (candidateTab) return candidateTab;
    return [thickness, widthPart].filter(Boolean).join(' × ');
  }

  function materialAvailabilityStatus({ requestedQty, producibleQty }) {
    const req = parsePositiveNumber(requestedQty);
    const prod = parsePositiveNumber(producibleQty);
    if (!req || prod === null) return 'unknown';
    return prod >= req ? 'ok' : 'short';
  }

  function lookupLocalMaterialStock({ byTab, thickness, materialWidth, productType = '', productName = '' }) {
    const mw = normalizeMaterialWidthValue(materialWidth);
    if (!mw || !byTab) return null;

    const thicknessKey = resolveMaterialInventoryThicknessKey({ thickness, productType, productName });
    const candidates = thicknessTabCandidates(thicknessKey);
    let totalKg = 0;
    let totalRolls = 0;
    const lots = [];
    const matchedTabs = [];

    for (const tabTitle of candidates) {
      const key = `${tabTitle}|${mw}`;
      const hit = byTab[key];
      if (!hit) continue;
      matchedTabs.push(tabTitle);
      totalKg += hit.totalKg || 0;
      totalRolls += hit.totalRolls || 0;
      lots.push(...(hit.lots || []));
    }

    if (!matchedTabs.length) return null;

    return {
      thickness: String(thickness ?? '').trim(),
      thicknessKey,
      materialWidth: mw,
      tabTitles: matchedTabs,
      totalKg: Math.round(totalKg * 1000) / 1000,
      totalRolls,
      lotCount: lots.length,
      lots,
    };
  }

  async function ensureMaterialStockData() {
    if (materialStockData) return materialStockData;
    if (!materialStockLoadPromise) {
      materialStockLoadPromise = fetch(`${API_BASE}/api/pq_form/material_stock`, { cache: 'no-store' })
        .then(async (res) => {
          const json = await res.json();
          if (!json.success) throw new Error(json.error || '材料在庫 API 錯誤');
          materialStockData = json;
          return materialStockData;
        })
        .catch((error) => {
          materialStockLoadPromise = null;
          throw error;
        });
    }
    return materialStockLoadPromise;
  }

  function formatReceiptDate(value) {
    const text = String(value ?? '').trim();
    const m = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (m) return `${parseInt(m[1], 10)}/${parseInt(m[2], 10)}/${m[3]}`;
    return text;
  }

  function receiptDateSortKey(value) {
    const text = String(value ?? '').trim();
    const m = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (!m) return 0;
    return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])).getTime();
  }

  function buildMaterialSummaryHtml({ stock, yieldInfo, requestedQty }) {
    const parts = [];
    if (stock?.totalKg !== null && stock?.totalKg !== undefined) {
      parts.push(`${escapeHtml(formatNumber(stock.totalKg, 1))} kg 可用`);
    }
    if (stock?.totalRolls) {
      parts.push(`${escapeHtml(formatNumber(stock.totalRolls))} 卷`);
    }
    if (yieldInfo?.producibleQty !== null && yieldInfo?.producibleQty !== undefined) {
      parts.push(`約 ${escapeHtml(formatNumber(yieldInfo.producibleQty))} 支可製造`);
    }
    const hasComparison = Boolean(requestedQty)
      && yieldInfo?.producibleQty !== null
      && yieldInfo?.producibleQty !== undefined;
    const summaryClass = hasComparison
      ? (yieldInfo?.status === 'short' ? 'stock-card__summary--short' : 'stock-card__summary--ok')
      : '';
    return `<div class="stock-card__summary ${summaryClass}">${parts.join(' | ')}</div>`;
  }

  function buildMaterialLotsHtml(lots = []) {
    const sorted = [...lots].sort((a, b) => receiptDateSortKey(a.receiptDate) - receiptDateSortKey(b.receiptDate));
    if (!sorted.length) return '';
    const items = sorted.map((lot) => {
      const date = formatReceiptDate(lot.receiptDate);
      const lotNo = escapeHtml(lot.lotNo || '');
      const rolls = lot.rolls !== null && lot.rolls !== undefined
        ? `${escapeHtml(formatNumber(lot.rolls))}卷`
        : '';
      const kg = lot.availableKg !== null && lot.availableKg !== undefined
        ? `${escapeHtml(formatNumber(lot.availableKg, 1))}kg`
        : '';
      return `<li class="stock-card__lot">${escapeHtml(date)}　${lotNo}${rolls ? `　${rolls}` : ''}${kg ? `　${kg}` : ''}</li>`;
    }).join('');
    return `<ul class="stock-card__lots">${items}</ul>`;
  }

  function getProductAvailForItem(itemNo) {
    const row = getItemRow(itemNo);
    const code = getField(row, 'productCode')?.value.trim() || '';
    if (!code || code === NOT_FOUND_CODE || !stockByCode) return 0;
    const stock = stockByCode[code.toUpperCase()];
    if (!stock || stock.available === null || stock.available === undefined) return 0;
    const avail = parsePositiveNumber(stock.available);
    return avail !== null ? avail : 0;
  }

  function resolveQtyHighlightStatus({ requestedQty, productAvail, producibleQty }) {
    const req = parsePositiveNumber(requestedQty);
    if (!req) return { status: '', hasComparison: false };

    const avail = Number.isFinite(productAvail) ? productAvail : 0;
    if (avail >= req) return { status: 'ok', hasComparison: true };

    const shortfall = req - avail;
    const prod = parsePositiveNumber(producibleQty);
    if (prod !== null && prod >= shortfall) return { status: 'partial', hasComparison: true };
    return { status: 'short', hasComparison: true };
  }

  function getQtyHighlightForItem(itemNo) {
    const row = getItemRow(itemNo);
    if (!row) return { status: '', hasComparison: false };
    const requestedQty = getField(row, 'quantity')?.value.trim() || '';
    const productAvail = getProductAvailForItem(itemNo);
    const producibleQty = materialProducibleByItem.get(itemNo);
    return resolveQtyHighlightStatus({
      requestedQty,
      productAvail,
      producibleQty: producibleQty ?? null,
    });
  }

  const AVAIL_HIGHLIGHT_CLASSES = [
    'stock-card__avail--ok',
    'stock-card__avail--partial',
    'stock-card__avail--short',
  ];

  function applyAvailHighlightToDetails(details, itemNo, stock) {
    const { status, hasComparison } = getQtyHighlightForItem(itemNo);
    let availEl = details.querySelector('.stock-card__avail');

    if (!availEl) {
      details.innerHTML = buildStockDetailsHtml(stock, itemNo);
      return;
    }

    availEl.classList.remove(...AVAIL_HIGHLIGHT_CLASSES);
    if (!hasComparison) {
      availEl.replaceWith(document.createTextNode(availEl.textContent));
      return;
    }

    if (status === 'ok') availEl.classList.add('stock-card__avail--ok');
    else if (status === 'partial') availEl.classList.add('stock-card__avail--partial');
    else if (status === 'short') availEl.classList.add('stock-card__avail--short');
  }

  function refreshStockCardAvailDisplay(itemNo) {
    const slot = stockCheckItems?.querySelector(`[data-item="${itemNo}"]`);
    const details = slot?.querySelector('.stock-card__details');
    if (!details) return;
    const row = getItemRow(itemNo);
    const code = getField(row, 'productCode')?.value.trim() || '';
    if (!code || code === NOT_FOUND_CODE || !stockByCode) return;
    const stock = stockByCode[code.toUpperCase()];
    if (!stock) return;
    applyAvailHighlightToDetails(details, itemNo, stock);
  }

  function refreshQtyHighlight(itemNo) {
    const { status, hasComparison } = getQtyHighlightForItem(itemNo);
    syncQtyHighlight(itemNo, { status, hasComparison });
    refreshStockCardAvailDisplay(itemNo);
  }

  function syncQtyHighlight(itemNo, { status = '', hasComparison = false } = {}) {
    const row = getItemRow(itemNo);
    const qtyInput = row ? getField(row, 'quantity') : null;
    const qtyCell = qtyInput?.closest('.pos-qty');
    if (!qtyCell) return;
    qtyCell.classList.remove('material-qty--ok', 'material-qty--partial', 'material-qty--short');
    if (!hasComparison) return;
    if (status === 'ok') qtyCell.classList.add('material-qty--ok');
    else if (status === 'partial') qtyCell.classList.add('material-qty--partial');
    else if (status === 'short') qtyCell.classList.add('material-qty--short');
  }

  function renderMaterialStockCardSlot(itemNo, state, payload = {}) {
    const slot = materialCheckItems?.querySelector(`[data-item="${itemNo}"]`);
    if (!slot) return;

    const {
      thickness = '',
      materialWidth = '',
      stock = null,
      yieldInfo = null,
      requestedQty = '',
      productType = '',
      productName = '',
    } = payload;
    const label = materialStockDisplayLabel({ thickness, materialWidth, stock, productType, productName });
    let cardHtml = '';

    if (state === 'empty') {
      cardHtml = `
        <div class="stock-card stock-card--material stock-card--empty">
          <div class="stock-card__placeholder">項目 ${itemNo}<br />輸入規格後顯示材料在庫</div>
        </div>`;
    } else if (state === 'loading') {
      cardHtml = `
        <div class="stock-card stock-card--material stock-card--loading">
          <div class="stock-card__placeholder">項目 ${itemNo}<br />材料在庫載入中…</div>
        </div>`;
    } else if (state === 'no-mw') {
      cardHtml = `
        <div class="stock-card stock-card--material stock-card--not-found">
          <div class="stock-card__code">材料在庫 | ${escapeHtml(label || '規格未齊')}</div>
          <div class="stock-card__details">用料闊度未設定</div>
        </div>`;
    } else if (state === 'error') {
      const message = payload.error || '材料在庫 API 錯誤';
      cardHtml = `
        <div class="stock-card stock-card--material stock-card--error">
          <div class="stock-card__code">材料在庫 | ${escapeHtml(label || '規格未齊')}</div>
          <div class="stock-card__details">${escapeHtml(message)}</div>
        </div>`;
    } else if (state === 'not-found') {
      cardHtml = `
        <div class="stock-card stock-card--material stock-card--not-found">
          <div class="stock-card__code">材料在庫 | ${escapeHtml(label)}</div>
          <div class="stock-card__details">此厚度+材料闊度無庫存</div>
        </div>`;
    } else if (state === 'ready' && stock) {
      cardHtml = `
        <div class="stock-card stock-card--material">
          <div class="stock-card__code">材料在庫 | ${escapeHtml(label)}</div>
          <div class="stock-card__scroll">
            ${buildMaterialSummaryHtml({ stock, yieldInfo, requestedQty })}
            ${buildMaterialLotsHtml(stock.lots)}
          </div>
        </div>`;
    }

    slot.innerHTML = cardHtml;
    scheduleStockCardAlignment();
  }

  function renderMaterialCheckSlots() {
    if (!materialCheckItems) return;
    materialCheckItems.innerHTML = Array.from({ length: ITEM_COUNT }, (_, i) => {
      const n = i + 1;
      return `<div class="material-check-slot" data-item="${n}"></div>`;
    }).join('');
    for (let i = 1; i <= ITEM_COUNT; i += 1) {
      renderMaterialStockCardSlot(i, 'empty');
    }
    scheduleStockCardAlignment();
  }

  async function updateMaterialStockCardForItem(itemNo) {
    const row = getItemRow(itemNo);
    if (!row) return;

    const thickness = getField(row, 'thickness')?.value.trim() || '';
    const materialWidth = getField(row, 'materialWidth')?.value.trim() || '';
    const length = getField(row, 'length')?.value.trim() || '';
    const requestedQty = getField(row, 'quantity')?.value.trim() || '';
    const productType = getField(row, 'productType')?.value.trim() || '';
    const productName = getField(row, 'productName')?.value.trim() || '';

    if (!thickness && !materialWidth) {
      materialProducibleByItem.delete(itemNo);
      renderMaterialStockCardSlot(itemNo, 'empty');
      refreshQtyHighlight(itemNo);
      return;
    }

    if (!materialWidth) {
      materialProducibleByItem.delete(itemNo);
      renderMaterialStockCardSlot(itemNo, 'no-mw', { thickness, productType, productName });
      refreshQtyHighlight(itemNo);
      return;
    }

    renderMaterialStockCardSlot(itemNo, 'loading', { thickness, materialWidth, productType, productName });

    try {
      const stockPayload = await ensureMaterialStockData();
      const stock = lookupLocalMaterialStock({
        byTab: stockPayload.byTab,
        thickness,
        materialWidth,
        productType,
        productName,
      });

      if (!stock) {
        materialProducibleByItem.set(itemNo, 0);
        renderMaterialStockCardSlot(itemNo, 'not-found', { thickness, materialWidth, productType, productName });
        refreshQtyHighlight(itemNo);
        return;
      }

      const producibleQty = calcProduciblePieces({
        availableKg: stock.totalKg,
        lengthMm: length,
        materialWidthMm: stock.materialWidth,
        thicknessMm: thickness,
        densityGcm3: densityFromMaterialStock(stock),
      });
      materialProducibleByItem.set(itemNo, producibleQty ?? 0);
      const yieldInfo = {
        producibleQty,
        status: materialAvailabilityStatus({ requestedQty, producibleQty }),
      };

      renderMaterialStockCardSlot(itemNo, 'ready', {
        thickness,
        materialWidth,
        stock,
        yieldInfo,
        requestedQty,
        productType,
        productName,
      });
      refreshQtyHighlight(itemNo);
    } catch (error) {
      console.error('updateMaterialStockCardForItem failed', error);
      renderMaterialStockCardSlot(itemNo, 'error', {
        thickness,
        materialWidth,
        error: error?.message || '材料在庫載入失敗',
        productType,
        productName,
      });
      refreshQtyHighlight(itemNo);
    }
  }

  function scheduleMaterialStockUpdate(itemNo) {
    clearTimeout(materialStockTimers.get(itemNo));
    materialStockTimers.set(itemNo, setTimeout(() => updateMaterialStockCardForItem(itemNo), 300));
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
      const itemNo = Number(row.dataset.item);
      updateStockCardForItem(itemNo);
      scheduleMaterialStockUpdate(itemNo);
    }
  }

  function clearProductOutputs(row) {
    setProductOutputs(row, '', '', false, null);
    if (row?.dataset?.item) {
      const itemNo = Number(row.dataset.item);
      renderStockCardSlot(itemNo, 'empty');
      renderMaterialStockCardSlot(itemNo, 'empty');
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
      strictType: '1',
    });

    try {
      const res = await fetch(`${API_BASE}/api/pq_form/plist/search?${params.toString()}`, { cache: 'no-store' });
      const data = await res.json();

      if (!data.success) {
        setProductOutputs(row, NOT_FOUND_CODE, NOT_FOUND_NAME, true, null);
        return;
      }

      if (data.matches.length === 1) {
        const match = data.matches[0];
        setProductOutputs(row, match.code, displayPlistProductName(match.name, spec), false, null);
        applyPlistMaterialWidthIfEmpty(row, match.materialWidth);
        persistLocal();
      } else if (data.matches.length > 1) {
        codeInput.value = '';
        const uniqueNames = [...new Set(data.matches.map((m) => displayPlistProductName(m.name, spec)))];
        nameInput.value = uniqueNames.length === 1 ? uniqueNames[0] : '';
        codeInput.classList.remove('product-not-found');
        nameInput.classList.remove('product-not-found');
        updateStockCardForItem(itemNo);
        showProductMatchPicker(itemNo, data.matches, (match) => {
          setProductOutputs(row, match.code, displayPlistProductName(match.name, spec), false, null);
          applyPlistMaterialWidthIfEmpty(row, match.materialWidth);
          persistLocal();
        });
      } else {
        setProductOutputs(
          row,
          NOT_FOUND_CODE,
          buildProvisionalProductName(spec.type, spec),
          false,
          null,
        );
        if (data.hint) {
          showProductResolveHint(`項目 ${itemNo}：${data.hint}`, data.hintType === 'no_spec');
        }
        persistLocal();
      }
    } catch (error) {
      console.error('plist search failed', error);
      setProductOutputs(row, NOT_FOUND_CODE, NOT_FOUND_NAME, true, null);
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
    const materialStockFields = ['thickness', 'materialWidth', 'length', 'quantity'];

    productItemsBody.addEventListener('input', (e) => {
      const row = e.target.closest('tr[data-item]');
      updateFieldFillState(e.target);
      if (row) {
        persistLocal();
        const itemNo = Number(row.dataset.item);
        if (specFields.includes(e.target.dataset.field)) {
          scheduleResolveProduct(itemNo);
        }
        if (materialStockFields.includes(e.target.dataset.field)) {
          if (e.target.dataset.field === 'quantity') {
            refreshQtyHighlight(itemNo);
          }
          scheduleMaterialStockUpdate(itemNo);
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
      const itemNo = Number(row.dataset.item);
      if (specFields.includes(e.target.dataset.field)) {
        scheduleResolveProduct(itemNo);
      }
      if (materialStockFields.includes(e.target.dataset.field)) {
        if (e.target.dataset.field === 'quantity') {
          refreshQtyHighlight(itemNo);
        }
        scheduleMaterialStockUpdate(itemNo);
      }
    });

    form.querySelectorAll('#deliveryNoteNoSuffix, #deliveryDate, #orderDate, #estimatedProductionPeriod, #completionDate, #personInCharge, #signature, #preparerSignature').forEach((el) => {
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

    if (customerNoInput && orderingCompanyInput) {
      const customerHandler = () => {
        updateFieldFillState(customerNoInput);
        updateFieldFillState(orderingCompanyInput);
        persistLocal();
      };
      customerNoInput.addEventListener('input', customerHandler);
      customerNoInput.addEventListener('change', customerHandler);
      orderingCompanyInput.addEventListener('input', customerHandler);
      orderingCompanyInput.addEventListener('change', customerHandler);
    }
    bindCustomerLookupEvents();
    bindPrintHandlers();

    if (printBtn) {
      printBtn.addEventListener('click', () => window.print());
    }

    clearBtn.addEventListener('click', () => {
      form.reset();
      for (let i = 1; i <= ITEM_COUNT; i += 1) {
        clearProductOutputs(getItemRow(i));
      }
      hideProductMatchPicker();
      hideCustomerCnNameSuggest();
      hideCustomerMatchPicker();
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
    renderMaterialCheckSlots();
    for (let i = 1; i <= ITEM_COUNT; i += 1) scheduleResolveProduct(i);
    syncAllStockCards();
    syncAllMaterialStockCards();
    scheduleStockCardAlignment();
  }

  initApp();
})();
