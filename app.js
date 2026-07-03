(() => {
  // 同一オリジン運用: ローカル(5508)のみ Flask(5013) を指し、Vercel 等は空(=同一オリジン)
  const isLocal = ['localhost','127.0.0.1'].includes(location.hostname);
  const isStaticDev = location.port === '5508';
  const API_BASE = (isLocal && isStaticDev) ? 'http://localhost:5013' : '';
  const tableBody = document.getElementById('tableBody');
  const materialTableBody = document.getElementById('materialTableBody');
  const tableBody2 = document.getElementById('tableBody2');
  const materialTableBody2 = document.getElementById('materialTableBody2');
  const STORAGE_KEY = 'pq-form-ui-v3';
  const addRowBtn = document.getElementById('addRowBtn');
  const removeRowBtn = document.getElementById('removeRowBtn');
  const clearBtn = document.getElementById('clearBtn');
  const addMaterialRowBtn = document.getElementById('addMaterialRowBtn');
  const removeMaterialRowBtn = document.getElementById('removeMaterialRowBtn');
  const clearMaterialBtn = document.getElementById('clearMaterialBtn');
  // 第2ページ用の要素
  const addRowBtn2 = document.getElementById('addRowBtn2');
  const removeRowBtn2 = document.getElementById('removeRowBtn2');
  const clearBtn2 = document.getElementById('clearBtn2');
  const addMaterialRowBtn2 = document.getElementById('addMaterialRowBtn2');
  const removeMaterialRowBtn2 = document.getElementById('removeMaterialRowBtn2');
  const clearMaterialBtn2 = document.getElementById('clearMaterialBtn2');

  function pad(n){return n.toString().padStart(2,'0');}

  function toTF(v){ return v ? 'TRUE' : 'FALSE'; }

  const FALLBACK_THICKNESS_OPTIONS = ['0.3', '0.4', '0.4D', '0.5', '0.6', '0.8', '0.8A', '1.0', '1.2', '1.5', '3.0'];
  let thicknessOptions = [...FALLBACK_THICKNESS_OPTIONS];

  function formatThicknessValue(value) {
    const text = String(value ?? '').trim();
    if (!text) return '';
    if (/[A-Za-z]/.test(text)) return text;
    const num = parseFloat(text);
    return Number.isFinite(num) ? num.toFixed(1) : text;
  }

  /** plist 照合のみ 0.8A→0.8、0.4D→0.4。材料厚度・產品名稱表示はそのまま */
  function thicknessForProductLookup(value) {
    const formatted = formatThicknessValue(value);
    if (!formatted) return '';
    const upper = formatted.toUpperCase();
    if (upper === '0.8A') return '0.8';
    if (upper === '0.4D') return '0.4';
    return formatted;
  }

  function getRecordedThickness(row) {
    if (isMaterialRow(row)) {
      return getThicknessValue(row, 5) || getThicknessValue(row, 2);
    }
    return getThicknessValue(row, 4);
  }

  function applyRecordedThicknessToProductName(name, recordedThickness) {
    const displayT = formatThicknessValue(recordedThickness);
    const lookupT = thicknessForProductLookup(recordedThickness);
    if (!name || !displayT || displayT === lookupT) return name;
    const escaped = lookupT.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return String(name).replace(new RegExp(`^${escaped}x`, 'i'), `${displayT}x`);
  }

  function thicknessMatches(a, b) {
    if (!a && !b) return true;
    if (!a || !b) return false;
    const sa = String(a).trim();
    const sb = String(b).trim();
    if (sa === sb) return true;
    if (/[A-Za-z]/.test(sa) || /[A-Za-z]/.test(sb)) return false;
    const na = parseFloat(sa);
    const nb = parseFloat(sb);
    if (Number.isFinite(na) && Number.isFinite(nb)) return na === nb;
    return false;
  }

  function normalizeThicknessOptions(list) {
    const seen = new Set();
    const out = [];
    for (const item of list) {
      const formatted = formatThicknessValue(item);
      if (!formatted) continue;
      const key = /[A-Za-z]/.test(formatted) ? formatted : parseFloat(formatted);
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(formatted);
    }
    return out.sort((a, b) => {
      const na = parseFloat(a);
      const nb = parseFloat(b);
      if (na !== nb) return na - nb;
      return String(a).localeCompare(String(b));
    });
  }

  function buildThicknessSelectOptions(selected = '') {
    const options = ['<option value=""></option>'];
    thicknessOptions.forEach((t) => {
      const display = formatThicknessValue(t);
      const isSelected = thicknessMatches(selected, display) ? ' selected' : '';
      options.push(`<option value="${display}"${isSelected}>${display}</option>`);
    });
    return options.join('');
  }

  function thicknessSelectHtml(selected = '') {
    return `<select class="thickness-select">${buildThicknessSelectOptions(selected)}</select>`;
  }

  function getThicknessSelect(row, colIndex) {
    return row.querySelector(`td:nth-child(${colIndex}) select.thickness-select`);
  }

  function getThicknessValue(row, colIndex = 4) {
    return getThicknessSelect(row, colIndex)?.value?.trim() || '';
  }

  function setThicknessSelectValue(select, value) {
    if (!select) return;
    const v = formatThicknessValue(value);
    if (!v) {
      select.value = '';
      return;
    }
    const hasOption = [...select.options].some((opt) => thicknessMatches(opt.value, v));
    const inList = thicknessOptions.some((t) => thicknessMatches(t, v));
    if (!hasOption && !inList) {
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = v;
      select.appendChild(opt);
    }
    const matched = [...select.options].find((opt) => thicknessMatches(opt.value, v));
    select.value = matched ? matched.value : v;
  }

  function refreshThicknessSelects() {
    document.querySelectorAll('select.thickness-select').forEach((sel) => {
      const current = sel.value;
      sel.innerHTML = buildThicknessSelectOptions();
      setThicknessSelectValue(sel, current);
    });
  }

  async function loadThicknessOptions() {
    let fromApi = [];
    try {
      const res = await fetch(`${API_BASE}/api/pq_form/plist/thicknesses`, { cache: 'no-store' });
      const data = await res.json();
      if (data.success && Array.isArray(data.thicknesses)) {
        fromApi = data.thicknesses;
      }
    } catch (error) {
      console.warn('loadThicknessOptions failed, using fallback', error);
    }
    thicknessOptions = normalizeThicknessOptions([...fromApi, ...FALLBACK_THICKNESS_OPTIONS]);
    refreshThicknessSelects();
  }

  // Baiduブラウザ対応の時間入力フォーマット関数（上料時間用）
  function formatTimeInput(input) {
    let value = input.value.replace(/\D/g, '');
    if (value.length >= 2) {
      value = value.substring(0, 2) + ':' + value.substring(2, 4);
    }
    input.value = value;
  }
  window.formatTimeInput = formatTimeInput;

  function clampTimePart(value, max) {
    const n = parseInt(String(value).replace(/\D/g, ''), 10);
    if (!Number.isFinite(n)) return '';
    return pad(Math.min(max, Math.max(0, n)));
  }

  function combineTimeParts(hour, minute) {
    const h = String(hour ?? '').trim();
    const m = String(minute ?? '').trim();
    if (!h && !m) return '';
    return `${h ? clampTimePart(h, 23) : '00'}:${m ? clampTimePart(m, 59) : '00'}`;
  }

  function timeSplitHtml(className) {
    return `<div class="time-split ${className}">
      <input type="text" class="time-hour" inputmode="numeric" maxlength="2" placeholder="時" aria-label="時" />
      <span class="time-sep">:</span>
      <input type="text" class="time-minute" inputmode="numeric" maxlength="2" placeholder="分" aria-label="分" />
    </div>`;
  }

  function getSplitTimeValue(row, className) {
    const wrap = row.querySelector(`.time-split.${className}`);
    if (!wrap) return '';
    const hour = wrap.querySelector('.time-hour')?.value ?? '';
    const minute = wrap.querySelector('.time-minute')?.value ?? '';
    return combineTimeParts(hour, minute);
  }

  function getSplitTimeParts(row, className) {
    const wrap = row.querySelector(`.time-split.${className}`);
    if (!wrap) return { hour: '', minute: '' };
    return {
      hour: String(wrap.querySelector('.time-hour')?.value ?? '').trim(),
      minute: String(wrap.querySelector('.time-minute')?.value ?? '').trim(),
    };
  }

  function isSplitTimeComplete(row, className) {
    const { hour, minute } = getSplitTimeParts(row, className);
    return hour !== '' && minute !== '';
  }

  function splitTimeToMinutes(row, className) {
    const { hour, minute } = getSplitTimeParts(row, className);
    if (!hour || !minute) return null;
    const h = parseInt(hour, 10);
    const m = parseInt(minute, 10);
    if (!Number.isFinite(h) || !Number.isFinite(m)) return null;
    if (h < 0 || h > 23 || m < 0 || m > 59) return null;
    return h * 60 + m;
  }

  function applySplitTimePaste(wrap, text) {
    const match = String(text).trim().match(/^(\d{1,2})(?::(\d{1,2}))?$/);
    if (!match || !wrap) return false;
    const hourInput = wrap.querySelector('.time-hour');
    const minuteInput = wrap.querySelector('.time-minute');
    if (hourInput) hourInput.value = clampTimePart(match[1], 23);
    if (minuteInput) minuteInput.value = clampTimePart(match[2] || '0', 59);
    return true;
  }

  function bindTimeSplitInputs(root) {
    root.addEventListener('input', (e) => {
      const input = e.target;
      if (!input.matches('.time-hour, .time-minute')) return;
      input.value = input.value.replace(/\D/g, '').substring(0, 2);
      if (input.classList.contains('time-hour') && input.value.length >= 2) {
        const minute = input.closest('.time-split')?.querySelector('.time-minute');
        minute?.focus();
        minute?.select();
      }
    });

    root.addEventListener('blur', (e) => {
      const input = e.target;
      if (!input.matches('.time-hour, .time-minute')) return;
      if (input.value === '') return;
      input.value = input.classList.contains('time-hour')
        ? clampTimePart(input.value, 23)
        : clampTimePart(input.value, 59);
    }, true);

    root.addEventListener('paste', (e) => {
      const input = e.target;
      if (!input.matches('.time-hour, .time-minute')) return;
      const wrap = input.closest('.time-split');
      const text = e.clipboardData?.getData('text') || '';
      if (!applySplitTimePaste(wrap, text)) return;
      e.preventDefault();
      wrap?.dispatchEvent(new Event('input', { bubbles: true }));
    });
  }

  // 3段階チェックボックスのクラス
  class ThreeStateCheckbox {
    constructor(element) {
      this.element = element;
      this.state = 0; // 0: 空欄, 1: ✓, 2: ✘
      this.element.addEventListener('click', () => this.toggle());
      this.updateDisplay();
    }

    toggle() {
      this.state = (this.state + 1) % 3; // 0→1→2→0のループ
      this.updateDisplay();
    }

    updateDisplay() {
      const text = this.element.querySelector('.checkbox-text');
      
      switch(this.state) {
        case 0: // 空欄
          text.textContent = '';
          text.className = 'checkbox-text empty';
          break;
        case 1: // ✓ (緑色)
          text.textContent = '✓';
          text.className = 'checkbox-text checked';
          break;
        case 2: // ✘ (赤色)
          text.textContent = '✘';
          text.className = 'checkbox-text failed';
          break;
      }
    }

    getValue() {
      switch(this.state) {
        case 0: return '';
        case 1: return '✓';
        case 2: return '✘';
      }
    }

    setValue(value) {
      switch(value) {
        case '': this.state = 0; break;
        case '✓': this.state = 1; break;
        case '✘': this.state = 2; break;
        default: this.state = 0; break;
      }
      this.updateDisplay();
    }
  }

  const HK_TIMEZONE = 'Asia/Hong_Kong';

  function getHongKongDateParts(date = new Date()) {
    const parts = new Intl.DateTimeFormat('en-CA', {
      timeZone: HK_TIMEZONE,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    }).formatToParts(date);
    const pick = (type) => parts.find((p) => p.type === type)?.value || '';
    return { y: pick('year'), m: pick('month'), d: pick('day') };
  }

  function setToday(){
    const { y, m, d } = getHongKongDateParts();
    [
      ['year', 'month', 'day'],
      ['year2', 'month2', 'day2'],
    ].forEach(([yId, mId, dId]) => {
      const yEl = document.getElementById(yId);
      const mEl = document.getElementById(mId);
      const dEl = document.getElementById(dId);
      if (yEl) yEl.value = y;
      if (mEl) mEl.value = m;
      if (dEl) dEl.value = d;
    });
    renderAllProductionRecords();
  }

  function createRow(){
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${timeSplitHtml('time-load')}</td>
      <td>${timeSplitHtml('time-start')}</td>
      <td><input type="text" value="" placeholder="" style="text-transform: uppercase;" autocomplete="off" /></td>
      <td>${thicknessSelectHtml()}</td>
      <td><input type="text" placeholder="闊度" /></td>
      <td><input type="text" placeholder="高度" /></td>
      <td class="name"><input type="text" value="" placeholder="" /></td>
      <td><input type="text" placeholder="長度" /></td>
      <td class="chk"><div class="three-state-checkbox" data-field="length_tolerance"><div class="checkbox-box"><span class="checkbox-text"></span></div></div></td>
      <td class="chk"><div class="three-state-checkbox" data-field="section_size"><div class="checkbox-box"><span class="checkbox-text"></span></div></div></td>
      <td class="chk"><div class="three-state-checkbox" data-field="left_right_bend"><div class="checkbox-box"><span class="checkbox-text"></span></div></div></td>
      <td class="chk"><div class="three-state-checkbox" data-field="up_down_bend"><div class="checkbox-box"><span class="checkbox-text"></span></div></div></td>
      <td class="chk"><div class="three-state-checkbox" data-field="twist"><div class="checkbox-box"><span class="checkbox-text"></span></div></div></td>
      <td>
        <select>
          <option value=""></option>
          <option value="達">達</option>
          <option value="群">群</option>
          <option value="嫻">嫻</option>
        </select>
      </td>
      <td>${timeSplitHtml('time-finish')}</td>
      <td>
        <select onchange="window.formatSpeedDisplay(this, { userInitiated: true })">
          <option value=""></option>
          <option value="轉機">轉機</option>
          ${Array.from({length: 25}, (_, i) => i * 5).map(speed => 
            `<option value="${speed}">${speed}</option>`
          ).join('')}
        </select>
      </td>
      <td><input type="text" placeholder="其他" /></td>
    `;
    
    // 3段階チェックボックスを初期化
    const checkboxes = tr.querySelectorAll('.three-state-checkbox');
    checkboxes.forEach(checkbox => {
      checkbox.threeStateInstance = new ThreeStateCheckbox(checkbox);
    });
    
    return tr;
  }

  function copyMainRowProductFields(sourceRow, targetRow) {
    const productInput = targetRow.querySelector('td:nth-child(3) input');
    if (productInput) productInput.value = sourceRow.querySelector('td:nth-child(3) input')?.value ?? '';

    setThicknessSelectValue(
      getThicknessSelect(targetRow, 4),
      getThicknessValue(sourceRow, 4),
    );

    const widthInput = targetRow.querySelector('td:nth-child(5) input');
    if (widthInput) widthInput.value = sourceRow.querySelector('td:nth-child(5) input')?.value ?? '';

    const heightInput = targetRow.querySelector('td:nth-child(6) input');
    if (heightInput) heightInput.value = sourceRow.querySelector('td:nth-child(6) input')?.value ?? '';

    const nameInput = targetRow.querySelector('td:nth-child(7) input');
    const sourceNameInput = sourceRow.querySelector('td:nth-child(7) input');
    if (nameInput && sourceNameInput) {
      nameInput.value = sourceNameInput.value;
      refreshProductNotFoundUI(productInput, nameInput);
    }

    const lengthInput = targetRow.querySelector('td:nth-child(8) input');
    if (lengthInput) lengthInput.value = sourceRow.querySelector('td:nth-child(8) input')?.value ?? '';
  }

  function insertMainRowAfter(sourceRow, newRow) {
    let anchor = sourceRow;
    const hintRow = anchor.nextElementSibling;
    if (hintRow?.classList.contains('product-hint-row')) {
      anchor = hintRow;
    }
    anchor.insertAdjacentElement('afterend', newRow);
  }

  function appendCopiedMainRow(sourceRow) {
    const body = sourceRow.closest('#tableBody, #tableBody2');
    if (!body) return null;
    const tr = createRow();
    copyMainRowProductFields(sourceRow, tr);
    insertMainRowAfter(sourceRow, tr);
    return tr;
  }

  function addRow(n = 1, targetTableBody = null) {
    const body = targetTableBody || tableBody;
    if (!body) return;

    for (let i = 0; i < n; i++) {
      const mainRows = [...body.querySelectorAll('tr')].filter(isMainDataRow);
      const lastRow = mainRows[mainRows.length - 1];
      const shouldCopy = lastRow?.querySelector('td:nth-child(16) select')?.value === '轉機';

      const tr = createRow();
      if (shouldCopy && lastRow) copyMainRowProductFields(lastRow, tr);
      body.appendChild(tr);
    }

    if (body === tableBody) persistLocal();
  }

  function addMaterialRow(n=1, targetTableBody=null){
    const tableBody = targetTableBody || materialTableBody;
    console.log('addMaterialRow called, n=', n);
    console.log('tableBody:', tableBody);
    for(let i=0;i<n;i++) {
      const row = createMaterialRow();
      console.log('Created row:', row);
      tableBody.appendChild(row);
    }
    persistLocal();
  }

  function orderNoCellHtml() {
    return `<td class="order-no-cell">
      <div class="order-no-editor">
        <select class="order-no-prefix" aria-label="單號類型">
          <option value=""></option>
          <option value="F">F</option>
          <option value="C">C</option>
          <option value="K">K</option>
        </select>
        <input type="text" class="order-no-number" placeholder="單號" inputmode="numeric" autocomplete="off" />
      </div>
      <button type="button" class="order-no-formatted" hidden></button>
    </td>`;
  }

  function parseOrderNo(value) {
    const raw = String(value ?? '').trim().toUpperCase();
    const fs = raw.match(/^FS#(\d+)$/);
    if (fs) return { prefix: 'F', number: fs[1], full: `FS#${fs[1]}` };
    const ck = raw.match(/^CK#(\d+)$/);
    if (ck) return { prefix: 'C', number: ck[1], full: `CK#${ck[1]}` };
    const ks = raw.match(/^KS#(\d+)$/);
    if (ks) return { prefix: 'K', number: ks[1], full: `KS#${ks[1]}` };
    const digits = raw.replace(/\D/g, '');
    return { prefix: '', number: digits, full: digits ? digits : '' };
  }

  function formatOrderNo(prefix, number) {
    const num = String(number ?? '').replace(/\D/g, '');
    if (!num) return '';
    if (prefix === 'F') return `FS#${num}`;
    if (prefix === 'C') return `CK#${num}`;
    if (prefix === 'K') return `KS#${num}`;
    return num;
  }

  function getMaterialOrderNoCell(tr) {
    return tr?.querySelector('td.order-no-cell') || null;
  }

  function showMaterialOrderNoEditor(cell) {
    if (!cell) return;
    const editor = cell.querySelector('.order-no-editor');
    const formatted = cell.querySelector('.order-no-formatted');
    if (editor) editor.hidden = false;
    if (formatted) formatted.hidden = true;
  }

  function finalizeMaterialOrderNo(cell) {
    if (!cell) return;
    const prefix = cell.querySelector('.order-no-prefix')?.value || '';
    const number = (cell.querySelector('.order-no-number')?.value || '').replace(/\D/g, '');
    const full = formatOrderNo(prefix, number);
    if (!full || !prefix) {
      showMaterialOrderNoEditor(cell);
      delete cell.dataset.orderNoFull;
      return;
    }
    cell.dataset.orderNoFull = full;
    const formatted = cell.querySelector('.order-no-formatted');
    const editor = cell.querySelector('.order-no-editor');
    if (formatted) {
      formatted.textContent = full;
      formatted.hidden = false;
    }
    if (editor) editor.hidden = true;
  }

  function getMaterialOrderNoValue(tr) {
    const cell = getMaterialOrderNoCell(tr);
    if (!cell) {
      return tr?.querySelector('td:nth-child(1) input')?.value?.trim() || '';
    }
    const formatted = cell.querySelector('.order-no-formatted');
    if (formatted && !formatted.hidden) {
      return formatted.textContent.trim();
    }
    const prefix = cell.querySelector('.order-no-prefix')?.value || '';
    const number = (cell.querySelector('.order-no-number')?.value || '').replace(/\D/g, '');
    return formatOrderNo(prefix, number);
  }

  function setMaterialOrderNoValue(tr, value) {
    const cell = getMaterialOrderNoCell(tr);
    if (!cell) {
      const legacy = tr.querySelector('td:nth-child(1) input');
      if (legacy) legacy.value = value || '';
      return;
    }
    const parsed = parseOrderNo(value);
    const prefixSelect = cell.querySelector('.order-no-prefix');
    const numberInput = cell.querySelector('.order-no-number');
    const formatted = cell.querySelector('.order-no-formatted');
    if (prefixSelect) prefixSelect.value = parsed.prefix;
    if (numberInput) numberInput.value = parsed.number;
    if (parsed.prefix && parsed.number) {
      cell.dataset.orderNoFull = parsed.full;
      if (formatted) {
        formatted.textContent = parsed.full;
        formatted.hidden = false;
      }
      cell.querySelector('.order-no-editor').hidden = true;
    } else {
      delete cell.dataset.orderNoFull;
      if (formatted) formatted.hidden = true;
      showMaterialOrderNoEditor(cell);
    }
  }

  function resetMaterialOrderNoCell(row) {
    const cell = getMaterialOrderNoCell(row);
    if (!cell) return;
    const prefixSelect = cell.querySelector('.order-no-prefix');
    const numberInput = cell.querySelector('.order-no-number');
    const formatted = cell.querySelector('.order-no-formatted');
    if (prefixSelect) prefixSelect.value = '';
    if (numberInput) numberInput.value = '';
    if (formatted) {
      formatted.textContent = '';
      formatted.hidden = true;
    }
    delete cell.dataset.orderNoFull;
    showMaterialOrderNoEditor(cell);
  }

  function bindMaterialOrderNoInputs(root) {
    root.addEventListener('input', (e) => {
      const input = e.target;
      if (!input.matches('.order-no-number')) return;
      input.value = input.value.replace(/\D/g, '');
    });
    root.addEventListener('blur', (e) => {
      const input = e.target;
      if (!input.matches('.order-no-number')) return;
      finalizeMaterialOrderNo(input.closest('td.order-no-cell'));
    }, true);
    root.addEventListener('change', (e) => {
      const select = e.target;
      if (!select.matches('.order-no-prefix')) return;
      const cell = select.closest('td.order-no-cell');
      const number = (cell?.querySelector('.order-no-number')?.value || '').replace(/\D/g, '');
      if (number) finalizeMaterialOrderNo(cell);
    });
    root.addEventListener('click', (e) => {
      const formatted = e.target.closest('.order-no-formatted');
      if (!formatted || formatted.hidden) return;
      showMaterialOrderNoEditor(formatted.closest('td.order-no-cell'));
      formatted.closest('td.order-no-cell')?.querySelector('.order-no-number')?.focus();
    });
  }

  function createMaterialRow(){
    const tr = document.createElement('tr');
    tr.innerHTML = `
      ${orderNoCellHtml()}
      <td>${thicknessSelectHtml()}</td>
      <td><input type="text" placeholder="闊度" /></td>
      <td><input type="text" placeholder="重量" /></td>
      <td>${thicknessSelectHtml()}</td>
      <td><input type="text" placeholder="闊度" /></td>
      <td><input type="text" placeholder="高度" /></td>
      <td><input type="text" oninput="window.searchProductName(this)" onkeyup="handleProductCodeInput(this, event)" style="text-transform: uppercase;" autocomplete="off" /></td>
      <td><input type="text" /></td>
      <td><input type="text" placeholder="長度" /></td>
      <td><input type="text" placeholder="數量" /></td>
      <td class="chk"><input type="checkbox" onchange="window.toggleCompletion(this)" /></td>
      <td class="chk"><input type="checkbox" /></td>
      <td class="chk"><input type="checkbox" checked class="checkbox-red" onchange="window.toggleIncomplete(this)" /></td>
      <td class="w-160 btns">
        <button class="btn btn-primary btn-row-send" type="button">送出</button>
      </td>
    `;
    
    return tr;
  }

  // グローバル変数：stockデータキャッシュ
  let stockDataCache = null;

  // 製品コード入力処理関数（バックスペース対応）
  window.handleProductCodeInput = function(inputElement, event) {
    // 大文字に変換
    inputElement.value = inputElement.value.toUpperCase();
    
    // バックスペースやDeleteキーの場合はフォーマット処理をスキップ
    if (event.key === 'Backspace' || event.key === 'Delete') {
      return;
    }
    
    // フォーマット処理を実行
    window.formatProductCode(inputElement);
  };

  // コード番号自動フォーマット関数
  window.formatProductCode = function(inputElement) {
    let value = inputElement.value.toUpperCase();
    
    // 空の場合は処理しない
    if (value.length === 0) {
      return;
    }
    
    // 既にハイフンが含まれている場合は処理しない
    if (value.includes('-')) {
      return;
    }
    
    // 2文字のプレフィックス（AC, AP, BD, FC）をチェック
    const prefixes = ['AC', 'AP', 'BD', 'FC'];
    const hasValidPrefix = prefixes.some(prefix => value.startsWith(prefix));
    
    if (hasValidPrefix && value.length === 2) {
      // 2文字入力されたら自動でハイフンを追加
      inputElement.value = value + '-';
    }
  };

  const transferRowModal = document.getElementById('transferRowModal');
  let transferRowConfirmResolver = null;

  function closeTransferRowConfirmModal() {
    if (transferRowModal) transferRowModal.hidden = true;
  }

  function showTransferRowConfirmModal(speedLabel) {
    const msgEl = document.getElementById('transferRowModalMessage');
    if (msgEl) {
      msgEl.textContent = `你揀咗「${speedLabel}」。係咪要自動加一行（上段＋用料記録），並複製呢行嘅產品編號、材料厚度、闊度、高度、產品名稱同長度？`;
    }
    if (transferRowModal) transferRowModal.hidden = false;
    return new Promise((resolve) => {
      transferRowConfirmResolver = resolve;
    });
  }

  function resolveTransferRowConfirm(confirmed) {
    closeTransferRowConfirmModal();
    if (transferRowConfirmResolver) {
      transferRowConfirmResolver(confirmed);
      transferRowConfirmResolver = null;
    }
  }

  function bindTransferRowConfirmEvents() {
    document.getElementById('transferRowConfirmBtn')?.addEventListener('click', () => resolveTransferRowConfirm(true));
    document.getElementById('transferRowCancelBtn')?.addEventListener('click', () => resolveTransferRowConfirm(false));
  }

  async function handleSpeedCopyRowConfirm(selectElement, row, speedLabel, datasetKey) {
    const confirmed = await showTransferRowConfirmModal(speedLabel);
    selectElement.dataset[datasetKey] = '1';
    if (confirmed) {
      appendCopiedMainRow(row);
      appendMaterialRowFromMainRow(row);
    }
    persistLocal();
  }

  async function handleTransferSpeedSelection(selectElement, row) {
    await handleSpeedCopyRowConfirm(selectElement, row, '轉機', 'transferRowAdded');
  }

  async function handleNumericSpeedSelection(selectElement, row, value) {
    await handleSpeedCopyRowConfirm(selectElement, row, `速${value}`, 'speedRowAdded');
  }

  // 速度表示フォーマット関数
  window.formatSpeedDisplay = async function(selectElement, options = {}) {
    const { userInitiated = false } = options;
    const value = selectElement.value;
    if (value && value !== '') {
      const selectedOption = selectElement.options[selectElement.selectedIndex];
      if (value === '轉機') {
        selectedOption.textContent = '轉機';
      } else {
        selectedOption.textContent = `速${value}`;
      }
      selectElement.style.background = '#f8f9fa';
      selectElement.style.fontWeight = '600';
      selectElement.style.color = 'var(--primary)';
    } else {
      const defaultOption = selectElement.options[0];
      defaultOption.textContent = '';
      selectElement.style.background = '#fff';
      selectElement.style.fontWeight = 'normal';
      selectElement.style.color = 'var(--text)';
    }

    if (userInitiated && value && value !== '') {
      const row = selectElement.closest('tr');
      if (row && isMainDataRow(row)) {
        if (value === '轉機' && !selectElement.dataset.transferRowAdded) {
          await handleTransferSpeedSelection(selectElement, row);
          return;
        }
        if (value !== '轉機' && !selectElement.dataset.speedRowAdded) {
          await handleNumericSpeedSelection(selectElement, row, value);
          return;
        }

        persistLocal();
      }
    }
  };

  // すべての入力フィールドをクリアする関数
  function clearProductNumberFields() {
    // 通常テーブルのすべての入力フィールドをクリア
    const normalTableRows = document.querySelectorAll('#tableBody tr');
    normalTableRows.forEach(row => {
      // すべてのinput要素をクリア
      const inputs = row.querySelectorAll('input[type="text"], input[type="time"]');
      inputs.forEach(input => {
        input.value = '';
      });
      
      // すべてのcheckboxをクリア
      const checkboxes = row.querySelectorAll('input[type="checkbox"]');
      checkboxes.forEach(checkbox => {
        checkbox.checked = false;
      });
      
      // すべてのselectをクリア
      const selects = row.querySelectorAll('select');
      selects.forEach(select => {
        select.selectedIndex = 0;
        // 速度選択の場合は表示もリセット
        if (select.onchange && select.onchange.toString().includes('formatSpeedDisplay')) {
          window.formatSpeedDisplay(select);
        }
      });
    });
    
    // 用料記録テーブルのすべての入力フィールドをクリア（未完成チェックは保持）
    const materialTableRows = document.querySelectorAll('#materialTableBody tr');
    materialTableRows.forEach(row => {
      row.querySelectorAll('input[type="text"], input[type="time"]').forEach((input) => {
        if (input.classList.contains('order-no-number')) return;
        input.value = '';
      });
      resetMaterialOrderNoCell(row);
      
      // チェックボックスをクリア（未完成チェックは除く）
      const checkboxes = row.querySelectorAll('input[type="checkbox"]');
      checkboxes.forEach((checkbox, index) => {
        // 未完成チェック（3番目のチェックボックス）はデフォルトでチェック
        if (index === 2) {
          checkbox.checked = true;
        } else {
          checkbox.checked = false;
        }
      });
      
      // すべてのselectをクリア
      const selects = row.querySelectorAll('select');
      selects.forEach(select => {
        if (select.classList.contains('order-no-prefix')) return;
        select.selectedIndex = 0;
      });
    });
    
    // 第2ページの用料記録テーブルも同様に処理
    const materialTableRows2 = document.querySelectorAll('#materialTableBody2 tr');
    materialTableRows2.forEach(row => {
      row.querySelectorAll('input[type="text"], input[type="time"]').forEach((input) => {
        if (input.classList.contains('order-no-number')) return;
        input.value = '';
      });
      resetMaterialOrderNoCell(row);
      
      // チェックボックスをクリア（未完成チェックは除く）
      const checkboxes = row.querySelectorAll('input[type="checkbox"]');
      checkboxes.forEach((checkbox, index) => {
        // 未完成チェック（3番目のチェックボックス）はデフォルトでチェック
        if (index === 2) {
          checkbox.checked = true;
        } else {
          checkbox.checked = false;
        }
      });
      
      // すべてのselectをクリア
      const selects = row.querySelectorAll('select');
      selects.forEach(select => {
        if (select.classList.contains('order-no-prefix')) return;
        select.selectedIndex = 0;
      });
    });
  }

  // 既存の監視システムのAPIを使用してstockデータを取得する関数
  async function fetchStockData() {
    if (stockDataCache) {
      return stockDataCache;
    }
    
    try {
      // API_BASEを使用してAPIを呼び出し
      const response = await fetch(`${API_BASE}/api/pq_form/stock`);
      const result = await response.json();
      
      if (result.success && result.data) {
        stockDataCache = result.data;
        console.log('Stock data loaded:', Object.keys(stockDataCache).length, 'items');
        return stockDataCache;
      } else {
        console.error('Failed to fetch stock data from API:', result.error);
        return {};
      }
    } catch (error) {
      console.error('Error fetching stock data from API:', error);
      return {};
    }
  }

  const resolveTimers = new WeakMap();

  function getPageRoot(el) {
    return el.closest('#autoPage') || getMoldingPageRoot();
  }

  function getSelectedType(pageRoot) {
    const checked = pageRoot?.querySelector('input[name="type"]:checked');
    return checked ? checked.value : '';
  }

  const NOT_FOUND_CODE = '暫時未搵到產品編碼';
  const NOT_FOUND_NAME = '暫時未搵到產品名稱';

  function normalizeProductCodeForSubmit(value) {
    const text = String(value ?? '').trim();
    if (!text || text === NOT_FOUND_CODE) return '';
    return text.toUpperCase();
  }

  function hideProductMatchPicker(row) {
    row.querySelector('.product-match-picker')?.remove();
  }

  function isMainDataRow(tr) {
    return tr && !tr.classList.contains('product-hint-row');
  }

  function hideProductResolveHint(row) {
    if (!row) return;
    delete row.dataset.hasHint;
    const next = row.nextElementSibling;
    if (next?.classList.contains('product-hint-row')) next.remove();

    const tableWrap = row.closest('.table-wrap');
    const panel = tableWrap && hintPanelByWrap.get(tableWrap);
    if (!panel) return;

    const tbody = row.parentElement;
    const anyHint = tbody && [...tbody.querySelectorAll('tr')].some(
      (tr) => isMainDataRow(tr) && tr.dataset.hasHint === '1'
    );
    if (!anyHint) {
      panel.hidden = true;
      panel.textContent = '';
      panel.classList.remove('product-hint-outside--success', 'product-hint-outside--error');
    }
  }

  const hintPanelByWrap = new WeakMap();

  function getHintPanel(tableWrap) {
    let panel = hintPanelByWrap.get(tableWrap);
    if (!panel) {
      panel = document.createElement('div');
      panel.className = 'product-hint-outside';
      panel.hidden = true;
      tableWrap.after(panel);
      hintPanelByWrap.set(tableWrap, panel);
    }
    return panel;
  }

  function showOutsideMessage(row, message, kind = 'error') {
    const tableWrap = row?.closest('.table-wrap');
    if (!tableWrap) return;

    const panel = getHintPanel(tableWrap);
    if (!message) {
      panel.hidden = true;
      panel.textContent = '';
      panel.classList.remove('product-hint-outside--success', 'product-hint-outside--error');
      return;
    }

    if (row) {
      const tbody = row.parentElement;
      tbody?.querySelectorAll('tr').forEach((tr) => {
        if (isMainDataRow(tr)) delete tr.dataset.hasHint;
      });
      if (kind === 'error') row.dataset.hasHint = '1';
    }

    const next = row?.nextElementSibling;
    if (next?.classList.contains('product-hint-row')) next.remove();

    panel.textContent = message;
    panel.hidden = false;
    panel.classList.remove('product-hint-outside--success', 'product-hint-outside--error');
    panel.classList.add(kind === 'success' ? 'product-hint-outside--success' : 'product-hint-outside--error');
  }

  function showProductResolveHint(row, message) {
    showOutsideMessage(row, message, 'error');
  }

  function getRowSpecValues(row) {
    return {
      thickness: getThicknessValue(row, 4),
      width: row.querySelector('td:nth-child(5) input')?.value?.trim() || '',
      height: row.querySelector('td:nth-child(6) input')?.value?.trim() || '',
      length: row.querySelector('td:nth-child(8) input')?.value?.trim() || '',
    };
  }

  function hasAllSpecValues(spec) {
    return !!(spec.thickness && spec.width && spec.height && spec.length);
  }

  function isProductNotFound(code, name) {
    return code === NOT_FOUND_CODE || name === NOT_FOUND_NAME;
  }

  function buildProvisionalProductName(type, spec, otherText = '') {
    const typeLabel = type === '其他' ? (otherText || '其他') : type;
    const t = formatThicknessValue(spec.thickness);
    const w = String(spec.width ?? '').trim();
    const h = String(spec.height ?? '').trim();
    const l = String(spec.length ?? '').trim();
    return `${t}x${w}x${h} ${typeLabel} ${l}mm`;
  }

  function refreshProductNotFoundUI(codeInput, nameInput) {
    if (!codeInput || !nameInput) return;
    codeInput.classList.toggle('product-not-found', codeInput.value === NOT_FOUND_CODE);
    nameInput.classList.toggle('product-not-found', nameInput.value === NOT_FOUND_NAME);
  }

  function setProductNotFoundUI(codeInput, nameInput, active) {
    if (!codeInput || !nameInput) return;
    if (active) {
      codeInput.value = NOT_FOUND_CODE;
      nameInput.value = NOT_FOUND_NAME;
    } else {
      if (codeInput.value === NOT_FOUND_CODE) codeInput.value = '';
      if (nameInput.value === NOT_FOUND_NAME) nameInput.value = '';
    }
    refreshProductNotFoundUI(codeInput, nameInput);
  }

  function syncProductNotFoundUI(sourceCodeInput, sourceNameInput, targetCodeInput, targetNameInput) {
    refreshProductNotFoundUI(sourceCodeInput, sourceNameInput);
    refreshProductNotFoundUI(targetCodeInput, targetNameInput);
  }

  function getMaterialRowSpecValues(row) {
    const thicknessRaw = getThicknessValue(row, 5) || getThicknessValue(row, 2);
    const widthRaw = row.querySelector('td:nth-child(6) input')?.value?.trim()
      || row.querySelector('td:nth-child(3) input')?.value?.trim()
      || '';
    return {
      thickness: thicknessRaw,
      width: widthRaw,
      height: row.querySelector('td:nth-child(7) input')?.value?.trim() || '',
      length: row.querySelector('td:nth-child(10) input')?.value?.trim() || '',
    };
  }

  function isMaterialRow(row) {
    return !!(row?.closest('#materialTableBody') || row.closest('#materialTableBody2'));
  }

  function getProductResolveContext(row) {
    if (isMaterialRow(row)) {
      return {
        spec: getMaterialRowSpecValues(row),
        codeInput: row.querySelector('td:nth-child(8) input'),
        nameInput: row.querySelector('td:nth-child(9) input'),
        pickerCol: 8,
      };
    }
    return {
      spec: getRowSpecValues(row),
      codeInput: row.querySelector('td:nth-child(3) input'),
      nameInput: row.querySelector('td:nth-child(7) input'),
      pickerCol: 3,
    };
  }

  function clearProductNotFoundFields(row) {
    const { codeInput, nameInput } = getProductResolveContext(row);
    setProductNotFoundUI(codeInput, nameInput, false);
  }

  function showProductNotFound(codeInput, nameInput) {
    setProductNotFoundUI(codeInput, nameInput, true);
  }

  function showProductCodeNotFoundWithProvisionalName(codeInput, nameInput, provisionalName) {
    if (!codeInput || !nameInput) return;
    codeInput.value = NOT_FOUND_CODE;
    nameInput.value = provisionalName;
    refreshProductNotFoundUI(codeInput, nameInput);
  }

  function showProductMatchPicker(row, matches, onPick, pickerCol = 3) {
    hideProductMatchPicker(row);
    const cell = row.querySelector(`td:nth-child(${pickerCol})`);
    if (!cell || matches.length === 0) return;
    const wrap = document.createElement('div');
    wrap.className = 'product-match-picker';
    const select = document.createElement('select');
    select.className = 'product-match-select';
    select.innerHTML = '<option value="">請選擇產品…</option>' +
      matches.map((m, i) => {
        const label = `${m.code} — ${m.name}`.replace(/"/g, '&quot;');
        return `<option value="${i}">${label}</option>`;
      }).join('');
    select.addEventListener('change', () => {
      if (select.value === '') return;
      onPick(matches[Number(select.value)]);
      hideProductMatchPicker(row);
      hideProductResolveHint(row);
    });
    wrap.appendChild(select);
    cell.appendChild(wrap);
  }

  function hasPendingProductMatchSelection(row) {
    const picker = row.querySelector('.product-match-picker select.product-match-select');
    return !!(picker && picker.value === '');
  }

  function getMainMaterialRowPair(tr, inAutoPage) {
    const isMaterialTable = tr.closest('#materialTableBody') !== null || tr.closest('#materialTableBody2') !== null;
    const mainBody = inAutoPage ? tableBody2 : tableBody;
    const materialBody = getMaterialBodyForPage(inAutoPage);
    const mainRows = [...mainBody.querySelectorAll('tr')].filter(isMainDataRow);
    const materialRows = [...materialBody.querySelectorAll('tr')];

    if (isMaterialTable) {
      const materialTr = tr;
      const linked = materialTr.dataset.mainRowIndex;
      if (linked !== undefined && linked !== '') {
        const mainTr = mainRows[Number(linked)] || null;
        return { mainTr, materialTr, rowIndex: Number(linked) };
      }
      const rowIndex = materialRows.indexOf(materialTr);
      const mainTr = mainRows[rowIndex] || null;
      return { mainTr, materialTr, rowIndex };
    }

    const mainTr = tr;
    const rowIndex = mainRows.indexOf(mainTr);
    const materialTr = findMaterialRowForMain(mainTr, inAutoPage)
      || materialRows[rowIndex]
      || null;
    return { mainTr, materialTr, rowIndex };
  }

  function validateMaterialRowBeforeSend(materialTr, hintRow) {
    if (!materialTr) {
      showOutsideMessage(hintRow, '搵唔到對應嘅用料記錄', 'error');
      return false;
    }
    const data = serializeMaterialRow(materialTr);
    const missing = [];
    if (!String(data.orderNo).trim()) missing.push('單號');
    if (!String(data.weight).trim()) missing.push('卷材重量');
    if (!String(data.length).trim()) missing.push('長度');
    if (!String(data.qty).trim()) missing.push('數量');

    const messages = [];
    if (missing.length) {
      messages.push(`請填寫用料記錄：${missing.join('、')}`);
    }
    if (!data.incomplete && (!data.complete || !data.oldCoil)) {
      messages.push('請勾選用料記錄嘅「完成」同「舊卷材」');
    }
    if (messages.length) {
      showOutsideMessage(hintRow, messages.join('；'), 'error');
      return false;
    }
    hideProductResolveHint(hintRow);
    return true;
  }

  function getMoldingPageRoot() {
    return document.getElementById('moldingPage');
  }

  function getPageRootForAuto(inAutoPage) {
    return inAutoPage ? document.getElementById('autoPage') : getMoldingPageRoot();
  }

  function hasMachineSelected(inAutoPage) {
    const root = getPageRootForAuto(inAutoPage);
    if (!root) return false;
    return [...root.querySelectorAll('input[name="machine"]')].some((el) => el.checked);
  }

  function showMachineSelectHint(inAutoPage, message) {
    const hint = getPageRootForAuto(inAutoPage)?.querySelector('.machine-select-hint');
    if (!hint) return;
    if (!message) {
      hint.hidden = true;
      hint.textContent = '';
      return;
    }
    hint.textContent = message;
    hint.hidden = false;
  }

  function validateMachineBeforeSend(inAutoPage) {
    if (hasMachineSelected(inAutoPage)) {
      showMachineSelectHint(inAutoPage, '');
      return true;
    }
    showMachineSelectHint(inAutoPage, '請選擇生產機械名稱');
    return false;
  }

  function validateRowBeforeSend(mainTr) {
    if (hasPendingProductMatchSelection(mainTr)) {
      showOutsideMessage(mainTr, '請選擇產品編號', 'error');
      return false;
    }
    const missing = [];
    if (!isSplitTimeComplete(mainTr, 'time-load')) missing.push('上料時間');
    if (!isSplitTimeComplete(mainTr, 'time-start')) missing.push('開始時間');
    if (!isSplitTimeComplete(mainTr, 'time-finish')) missing.push('完成時間');
    if (missing.length) {
      showOutsideMessage(mainTr, `請填寫${missing.join('、')}`, 'error');
      return false;
    }
    const startMinutes = splitTimeToMinutes(mainTr, 'time-start');
    const finishMinutes = splitTimeToMinutes(mainTr, 'time-finish');
    if (startMinutes !== null && finishMinutes !== null && finishMinutes < startMinutes) {
      showOutsideMessage(mainTr, '完成時間唔可以早過開始時間', 'error');
      return false;
    }
    hideProductResolveHint(mainTr);
    return true;
  }

  function isSpecInputCell(row, el) {
    const td = el.closest('td');
    if (!td || !row.contains(td)) return false;
    const idx = [...row.children].indexOf(td) + 1;
    if (idx === 4 && el.matches('select.thickness-select')) return true;
    return [5, 6, 8].includes(idx) && el.matches('input');
  }

  function isMaterialSpecInputCell(row, el) {
    if (!isMaterialRow(row)) return false;
    const td = el.closest('td');
    if (!td || !row.contains(td)) return false;
    const idx = [...row.children].indexOf(td) + 1;
    if ((idx === 2 || idx === 5) && el.matches('select.thickness-select')) return true;
    return [3, 6, 7, 10].includes(idx) && el.matches('input');
  }

  async function tryResolveProductForRow(row) {
    const pageRoot = getPageRoot(row);
    const type = getSelectedType(pageRoot);
    const { spec, codeInput, nameInput, pickerCol } = getProductResolveContext(row);
    const { thickness, width, height, length } = spec;

    if (!hasAllSpecValues(spec)) {
      hideProductResolveHint(row);
      hideProductMatchPicker(row);
      clearProductNotFoundFields(row);
      return;
    }

    if (!type) {
      showProductResolveHint(row, '請先選擇產品種類');
      return;
    }
    if (type === '其他') {
      const other = pageRoot.querySelector('#typeOther')?.value?.trim();
      if (!other) {
        showProductResolveHint(row, '請輸入其他產品種類');
        return;
      }
    }

    if (!codeInput || !nameInput) return;

    hideProductResolveHint(row);

    const lookupThickness = thicknessForProductLookup(thickness);
    const recordedThickness = getRecordedThickness(row);
    const params = new URLSearchParams({ type, t: lookupThickness, w: width, h: height, l: length });
    if (type === '其他') {
      params.set('other', pageRoot.querySelector('#typeOther')?.value?.trim() || '');
    }

    const displayProductName = (name) => applyRecordedThicknessToProductName(name, recordedThickness);
    const displaySpec = { ...spec, thickness: recordedThickness || thickness };

    try {
      const res = await fetch(`${API_BASE}/api/pq_form/plist/search?${params.toString()}`, { cache: 'no-store' });
      const data = await res.json();
      if (!data.success) {
        hideProductMatchPicker(row);
        showProductNotFound(codeInput, nameInput);
        persistLocal();
        adjustNameColumnWidth();
        return;
      }

      hideProductMatchPicker(row);
      if (data.matches.length === 1) {
        codeInput.value = data.matches[0].code;
        nameInput.value = displayProductName(data.matches[0].name);
        codeInput.classList.remove('product-not-found');
        nameInput.classList.remove('product-not-found');
        hideProductResolveHint(row);
        persistLocal();
        adjustNameColumnWidth();
      } else if (data.matches.length > 1) {
        const uniqueNames = [...new Set(data.matches.map((m) => displayProductName(m.name)))];
        codeInput.value = '';
        codeInput.classList.remove('product-not-found');
        nameInput.classList.remove('product-not-found');
        if (uniqueNames.length === 1) {
          nameInput.value = uniqueNames[0];
        } else {
          nameInput.value = '';
        }
        showProductMatchPicker(row, data.matches, (match) => {
          codeInput.value = match.code;
          nameInput.value = displayProductName(match.name);
          codeInput.classList.remove('product-not-found');
          nameInput.classList.remove('product-not-found');
          hideProductResolveHint(row);
          persistLocal();
          adjustNameColumnWidth();
        }, pickerCol);
      } else {
        const other = type === '其他'
          ? pageRoot.querySelector('#typeOther')?.value?.trim() || ''
          : '';
        showProductCodeNotFoundWithProvisionalName(
          codeInput,
          nameInput,
          buildProvisionalProductName(type, displaySpec, other),
        );
        if (data.hint) showProductResolveHint(row, data.hint);
        persistLocal();
        adjustNameColumnWidth();
      }
    } catch (error) {
      console.error('plist search failed', error);
      hideProductMatchPicker(row);
      showProductNotFound(codeInput, nameInput);
      persistLocal();
      adjustNameColumnWidth();
    }
  }

  function scheduleResolveProduct(row) {
    clearTimeout(resolveTimers.get(row));
    resolveTimers.set(row, setTimeout(() => tryResolveProductForRow(row), 300));
  }

  function bindSingleTypeSelection(pageRoot) {
    if (!pageRoot) return;
    pageRoot.querySelectorAll('input[name="type"]').forEach((el) => {
      el.addEventListener('change', () => {
        if (!el.checked) return;
        pageRoot.querySelectorAll('input[name="type"]').forEach((other) => {
          if (other !== el) other.checked = false;
        });
        const tableId = pageRoot.id === 'autoPage' ? '#tableBody2' : '#tableBody';
        const materialTableId = pageRoot.id === 'autoPage' ? '#materialTableBody2' : '#materialTableBody';
        pageRoot.querySelectorAll(`${tableId} tr`).forEach((row) => {
          if (!isMainDataRow(row)) return;
          scheduleResolveProduct(row);
        });
        pageRoot.querySelectorAll(`${materialTableId} tr`).forEach((row) => {
          scheduleResolveProduct(row);
        });
      });
    });

    const typeOther = pageRoot.querySelector('#typeOther');
    if (typeOther) {
      typeOther.addEventListener('input', () => {
        const otherChecked = pageRoot.querySelector('input[name="type"][value="其他"]')?.checked;
        if (!otherChecked) return;
        pageRoot.querySelectorAll('#tableBody tr, #materialTableBody tr').forEach((row) => {
          if (row.closest('#tableBody') && !isMainDataRow(row)) return;
          scheduleResolveProduct(row);
        });
      });
    }
  }

  function getSelectedMachineValue(pageRoot) {
    return pageRoot?.querySelector('input[name="machine"]:checked')?.value || null;
  }

  function setMachineSelection(pageRoot, value) {
    if (!pageRoot) return;
    pageRoot.querySelectorAll('input[name="machine"]').forEach((input) => {
      input.checked = value ? input.value === value : false;
    });
  }

  const lastMachineByPage = new WeakMap();
  const machineResetModal = document.getElementById('machineResetModal');
  let machineResetConfirmResolver = null;

  function closeMachineResetConfirmModal() {
    if (machineResetModal) machineResetModal.hidden = true;
  }

  function showMachineResetConfirmModal() {
    if (machineResetModal) machineResetModal.hidden = false;
    return new Promise((resolve) => {
      machineResetConfirmResolver = resolve;
    });
  }

  function resolveMachineResetConfirm(confirmed) {
    closeMachineResetConfirmModal();
    if (machineResetConfirmResolver) {
      machineResetConfirmResolver(confirmed);
      machineResetConfirmResolver = null;
    }
  }

  function bindMachineResetConfirmEvents() {
    document.getElementById('machineResetConfirmBtn')?.addEventListener('click', () => resolveMachineResetConfirm(true));
    document.getElementById('machineResetCancelBtn')?.addEventListener('click', () => resolveMachineResetConfirm(false));
  }

  function normalizeMachineSelection(pageRoot) {
    if (!pageRoot) return null;
    const checked = [...pageRoot.querySelectorAll('input[name="machine"]:checked')];
    if (checked.length <= 1) return checked[0]?.value || null;
    const keepValue = checked[checked.length - 1].value;
    setMachineSelection(pageRoot, keepValue);
    return keepValue;
  }

  function bindSingleMachineSelection(pageRoot) {
    if (!pageRoot || pageRoot.dataset.machineSingleBound === '1') return;
    pageRoot.dataset.machineSingleBound = '1';
    const inAutoPage = pageRoot.id === 'autoPage';
    lastMachineByPage.set(pageRoot, normalizeMachineSelection(pageRoot));

    pageRoot.querySelectorAll('input[name="machine"]').forEach((el) => {
      el.addEventListener('change', async () => {
        const previous = lastMachineByPage.get(pageRoot) || null;

        if (!el.checked) {
          if (previous === el.value) el.checked = true;
          return;
        }

        if (previous && previous !== el.value) {
          setMachineSelection(pageRoot, previous);
          const confirmed = await showMachineResetConfirmModal();
          if (!confirmed) return;
          clearAllForPage(inAutoPage);
          setMachineSelection(pageRoot, el.value);
          lastMachineByPage.set(pageRoot, el.value);
          showMachineSelectHint(inAutoPage, '');
          persistLocal();
          return;
        }

        pageRoot.querySelectorAll('input[name="machine"]').forEach((other) => {
          if (other !== el) other.checked = false;
        });
        lastMachineByPage.set(pageRoot, el.value);
        showMachineSelectHint(inAutoPage, '');
      });
    });
  }

  function bindMachineSelectHints(pageRoot) {
    bindSingleMachineSelection(pageRoot);
  }

  function bindPlistResolveInputs(root) {
    root.addEventListener('input', (e) => {
      const input = e.target;
      if (!input.matches('input')) return;
      const row = input.closest('#tableBody tr, #tableBody2 tr, #materialTableBody tr, #materialTableBody2 tr');
      if (!row) return;
      if (isMainDataRow(row) && isSpecInputCell(row, input)) {
        scheduleResolveProduct(row);
        return;
      }
      if (isMaterialSpecInputCell(row, input)) {
        scheduleResolveProduct(row);
      }
    });
    root.addEventListener('change', (e) => {
      const select = e.target;
      if (!select.matches('select.thickness-select')) return;
      const row = select.closest('#tableBody tr, #tableBody2 tr, #materialTableBody tr, #materialTableBody2 tr');
      if (!row) return;
      if (isMainDataRow(row)) {
        scheduleResolveProduct(row);
        return;
      }
      if (isMaterialSpecInputCell(row, select)) {
        scheduleResolveProduct(row);
      }
    });
  }

  function getTopProductSyncFields(row) {
    return {
      thickness: getThicknessValue(row, 4),
      width: row.querySelector('td:nth-child(5) input')?.value ?? '',
      height: row.querySelector('td:nth-child(6) input')?.value ?? '',
      code: row.querySelector('td:nth-child(3) input')?.value ?? '',
      name: row.querySelector('td:nth-child(7) input')?.value ?? '',
      length: row.querySelector('td:nth-child(8) input')?.value ?? '',
    };
  }

  function getMainRowIndex(topRow) {
    if (!topRow?.parentElement) return -1;
    return [...topRow.parentElement.querySelectorAll('tr')].filter(isMainDataRow).indexOf(topRow);
  }

  function isMaterialRowTemplateEmpty(tr) {
    const data = serializeMaterialRow(tr);
    return !String(data.orderNo).trim()
      && !String(data.weight).trim()
      && !String(data.productNo).trim()
      && !String(data.name).trim()
      && !String(data.length).trim()
      && !String(data.qty).trim()
      && !String(data.width1).trim()
      && !String(data.height).trim()
      && !String(data.thickness1).trim();
  }

  function removeTemplateEmptyMaterialRows(materialBody) {
    if (!materialBody) return;
    [...materialBody.querySelectorAll('tr')].forEach((tr) => {
      if (isMaterialRowTemplateEmpty(tr)) tr.remove();
    });
  }

  function fillMaterialRowFromMain(materialRow, topRow) {
    const { thickness, width, height, code, name, length } = getTopProductSyncFields(topRow);
    const thicknessSelect1 = getThicknessSelect(materialRow, 2);
    const thicknessSelect2 = getThicknessSelect(materialRow, 5);
    const widthInput2 = materialRow.querySelector('td:nth-child(6) input');
    const heightInput = materialRow.querySelector('td:nth-child(7) input');
    const codeInput = materialRow.querySelector('td:nth-child(8) input');
    const nameInput = materialRow.querySelector('td:nth-child(9) input');
    const lengthInput = materialRow.querySelector('td:nth-child(10) input');

    setThicknessSelectValue(thicknessSelect1, thickness);
    setThicknessSelectValue(thicknessSelect2, thickness);
    if (widthInput2) widthInput2.value = width;
    if (heightInput) heightInput.value = height;
    if (codeInput) codeInput.value = code;
    if (nameInput) nameInput.value = name;
    if (lengthInput) lengthInput.value = length;

    const topCodeInput = topRow.querySelector('td:nth-child(3) input');
    const topNameInput = topRow.querySelector('td:nth-child(7) input');
    syncProductNotFoundUI(topCodeInput, topNameInput, codeInput, nameInput);
  }

  function findMaterialRowForMain(mainTr, inAutoPage) {
    const materialBody = getMaterialBodyForPage(inAutoPage);
    if (!materialBody || !mainTr) return null;
    const mainRowIndex = getMainRowIndex(mainTr);
    if (mainRowIndex < 0) return null;

    const materialRows = [...materialBody.querySelectorAll('tr')];
    const linked = materialRows.find((tr) => tr.dataset.mainRowIndex === String(mainRowIndex));
    if (linked) return linked;

    const atIndex = materialRows[mainRowIndex];
    if (atIndex && !isMaterialRowTemplateEmpty(atIndex)) return atIndex;

    return null;
  }

  function appendMaterialRowFromMainRow(topRow) {
    const pageRoot = getPageRoot(topRow);
    const isAuto = pageRoot?.id === 'autoPage';
    const materialBody = isAuto ? materialTableBody2 : materialTableBody;
    if (!materialBody) return null;

    const mainRowIndex = getMainRowIndex(topRow);
    if (mainRowIndex < 0) return null;

    const materialRows = [...materialBody.querySelectorAll('tr')];
    let materialRow = materialRows.find((tr) => tr.dataset.mainRowIndex === String(mainRowIndex));
    if (!materialRow) {
      const atIndex = materialRows[mainRowIndex];
      if (atIndex && !isMaterialRowTemplateEmpty(atIndex)) {
        materialRow = atIndex;
        materialRow.dataset.mainRowIndex = String(mainRowIndex);
      }
    }

    if (materialRow) {
      fillMaterialRowFromMain(materialRow, topRow);
      removeTemplateEmptyMaterialRows(materialBody);
      return materialRow;
    }

    removeTemplateEmptyMaterialRows(materialBody);
    materialRow = createMaterialRow();
    materialRow.dataset.mainRowIndex = String(mainRowIndex);
    materialBody.appendChild(materialRow);
    fillMaterialRowFromMain(materialRow, topRow);
    return materialRow;
  }

  // 產品編號から產品名稱を検索する関数（下のテーブル用）
  window.searchProductName = async function(inputElement) {
    // 入力値を大文字に変換
    const inputValue = inputElement.value.trim().toUpperCase();
    inputElement.value = inputValue; // 表示も大文字に更新
    
    if (!inputValue) {
      return;
    }
    
    try {
      const stockData = await fetchStockData();
      // 大文字小文字を区別せずに検索
      let productInfo = null;
      
      // 直接検索
      if (stockData[inputValue]) {
        productInfo = stockData[inputValue];
      } else {
        // 大文字小文字を無視して検索
        for (const key in stockData) {
          if (key.toUpperCase() === inputValue) {
            productInfo = stockData[key];
            break;
          }
        }
      }
      
      if (productInfo) {
        const row = inputElement.closest('tr');
        
        // 產品名稱を設定（9番目の列）
        const productNameInput = row.querySelector('td:nth-child(9) input[type="text"]');
        if (productNameInput) {
          productNameInput.value = productInfo.name || '';
          console.log('Product name set:', productInfo.name);
        }
        
        // 材料厚度2を設定（5番目の列）
        const thickness2Select = getThicknessSelect(row, 5);
        if (thickness2Select) {
          setThicknessSelectValue(thickness2Select, productInfo.thickness2 || '');
          console.log('Thickness2 set:', productInfo.thickness2);
        }
        
        // 闊度2を設定（6番目の列）
        const width2Input = row.querySelector('td:nth-child(6) input[type="text"]');
        if (width2Input) {
          width2Input.value = productInfo.width2 || '';
          console.log('Width2 set:', productInfo.width2);
        }
        
        // 高度を設定（7番目の列）
        const heightInput = row.querySelector('td:nth-child(7) input[type="text"]');
        if (heightInput) {
          heightInput.value = productInfo.height || '';
          console.log('Height set:', productInfo.height);
        }
        
      } else {
        console.log('Product not found for:', productNumber);
      }
    } catch (error) {
      console.error('Error searching product name:', error);
    }
  };

  // 產品編號から產品名稱を検索する関数（上のテーブル用）
  window.searchProductNameTop = async function(inputElement) {
    // 入力値を大文字に変換
    const inputValue = inputElement.value.trim().toUpperCase();
    inputElement.value = inputValue; // 表示も大文字に更新
    
    if (!inputValue) {
      return;
    }
    
    try {
      const stockData = await fetchStockData();
      console.log('Searching for:', inputValue);
      console.log('Total keys count:', Object.keys(stockData).length);
      console.log('Available keys (first 20):', Object.keys(stockData).slice(0, 20));
      console.log('AC keys:', Object.keys(stockData).filter(key => key.startsWith('AC')).slice(0, 10));
      console.log('Searching for exact match:', inputValue);
      console.log('Direct lookup result:', stockData[inputValue]);
      
      // 大文字小文字を区別せずに検索
      let productInfo = null;
      
      // 直接検索
      if (stockData[inputValue]) {
        productInfo = stockData[inputValue];
        console.log('Found by direct match:', inputValue);
      } else {
        // 大文字小文字を無視して検索
        for (const key in stockData) {
          if (key.toUpperCase() === inputValue) {
            productInfo = stockData[key];
            console.log('Found by case-insensitive match:', key);
            break;
          }
        }
        
        // 部分一致検索（AC-15 → AC-015）
        if (!productInfo) {
          for (const key in stockData) {
            if (key.toUpperCase().includes(inputValue) || inputValue.includes(key.toUpperCase())) {
              productInfo = stockData[key];
              console.log('Found by partial match:', key, 'for input:', inputValue);
              break;
            }
          }
        }
      }
      
      if (productInfo) {
        const row = inputElement.closest('tr');
        
        // 產品名稱を設定（7番目の列）
        const productNameInput = row.querySelector('td:nth-child(7) input[type="text"]');
        if (productNameInput) {
          productNameInput.value = productInfo.name || '';
          console.log('Product name set (top table):', productInfo.name);
        }
        
        // 材料厚度を設定（4番目の列）→ K列
        const thicknessSelect = getThicknessSelect(row, 4);
        if (thicknessSelect) {
          setThicknessSelectValue(thicknessSelect, productInfo.thickness2 || '');
          console.log('Thickness set (top table):', productInfo.thickness2);
        } else {
          console.log('Thickness select not found (top table)');
        }
        
        // 闊度を設定（5番目の列）→ H列（闊度1）
        const widthInput = row.querySelector('td:nth-child(5) input[type="text"]');
        if (widthInput) {
          widthInput.value = productInfo.width2 || '';
          console.log('Width set (top table):', productInfo.width2);
        } else {
          console.log('Width input not found (top table)');
        }
        
        // 高度を設定（6番目の列）→ I列
        const heightInput = row.querySelector('td:nth-child(6) input[type="text"]');
        if (heightInput) {
          heightInput.value = productInfo.height || '';
          console.log('Height set (top table):', productInfo.height);
        } else {
          console.log('Height input not found (top table)');
        }
        
        // 長度を設定（8番目の列）→ F列
        const lengthInput = row.querySelector('td:nth-child(8) input[type="text"]');
        if (lengthInput) {
          lengthInput.value = productInfo.length || '';
          console.log('Length set (top table):', productInfo.length);
        } else {
          console.log('Length input not found (top table)');
        }
        
      } else {
        console.log('Product not found for (top table):', inputValue);
      }
    } catch (error) {
      console.error('Error searching product name (top table):', error);
    }
  };


  window.toggleCompletion = function(checkbox) {
    const row = checkbox.closest('tr');
    // 未完成は最後のチェックボックス（3番目）
    const incompleteCheckbox = row.querySelector('td:nth-child(14) input[type="checkbox"]');
    console.log('toggleCompletion: incompleteCheckbox found:', incompleteCheckbox);
    
    if (checkbox.checked) {
      // 未完成を外す
      if (incompleteCheckbox) {
        incompleteCheckbox.checked = false;
        incompleteCheckbox.classList.remove('checkbox-red', 'checkbox-blue');
        console.log('未完成 unchecked');
      }
      // 完成を青色にする
      checkbox.classList.add('checkbox-blue');
      checkbox.classList.remove('checkbox-red');
      console.log('完成 checked (blue)');
    } else {
      // 完成を外したら未完成を自動チェック
      checkbox.classList.remove('checkbox-blue', 'checkbox-red');
      if (incompleteCheckbox) {
        incompleteCheckbox.checked = true;
        incompleteCheckbox.classList.add('checkbox-red');
        incompleteCheckbox.classList.remove('checkbox-blue');
        console.log('完成 unchecked, 未完成 auto-checked (red)');
      }
    }
  }

  window.toggleIncomplete = function(checkbox) {
    const row = checkbox.closest('tr');
    const completeCheckbox = row.querySelector('td:nth-child(12) input[type="checkbox"]');
    console.log('toggleIncomplete: completeCheckbox found:', completeCheckbox);
    
    if (checkbox.checked) {
      // 完成を外す
      if (completeCheckbox) {
        completeCheckbox.checked = false;
        completeCheckbox.classList.remove('checkbox-blue', 'checkbox-red');
        console.log('完成 unchecked');
      }
      // 未完成を赤色にする
      checkbox.classList.add('checkbox-red');
      checkbox.classList.remove('checkbox-blue');
      console.log('未完成 checked (red)');
    } else {
      checkbox.classList.remove('checkbox-red', 'checkbox-blue');
      console.log('未完成 unchecked');
    }
  }

  function removeRow(targetTableBody=null){
    const body = targetTableBody || tableBody;
    if (!body) return;
    let last = body.lastElementChild;
    if (last?.classList.contains('product-hint-row')) {
      body.removeChild(last);
      last = body.lastElementChild;
    }
    if (last && isMainDataRow(last)) body.removeChild(last);
    const trailingHint = body.lastElementChild;
    if (trailingHint?.classList.contains('product-hint-row')) {
      body.removeChild(trailingHint);
    }
    if (body === tableBody) persistLocal();
  }

  function removeMaterialRow(targetTableBody=null){
    const body = targetTableBody || materialTableBody;
    const last = body?.lastElementChild;
    if (last) body.removeChild(last);
    if (body === materialTableBody) persistLocal();
  }

  function clearAll(){
    tableBody.innerHTML='';
    materialTableBody.innerHTML='';
    persistLocal();
  }

  function clearMaterialAll(){
    materialTableBody.innerHTML='';
    persistLocal();
  }

  function clearAllForPage(inAutoPage) {
    const mainBody = inAutoPage ? tableBody2 : tableBody;
    const materialBody = inAutoPage ? materialTableBody2 : materialTableBody;
    if (mainBody) mainBody.innerHTML = '';
    if (materialBody) materialBody.innerHTML = '';
    if (!inAutoPage) persistLocal();
  }

  function serializeMainRow(tr) {
    const threeStateFields = ['length_tolerance', 'section_size', 'left_right_bend', 'up_down_bend', 'twist'];
    const checks = {};
    threeStateFields.forEach((field) => {
      const checkbox = tr.querySelector(`.three-state-checkbox[data-field="${field}"]`);
      checks[field] = checkbox?.threeStateInstance?.getValue?.() || '';
    });
    return {
      load: getSplitTimeValue(tr, 'time-load'),
      start: getSplitTimeValue(tr, 'time-start'),
      productNo: normalizeProductCodeForSubmit(tr.querySelector('td:nth-child(3) input')?.value),
      thickness: getThicknessValue(tr, 4),
      width: tr.querySelector('td:nth-child(5) input')?.value || '',
      height: tr.querySelector('td:nth-child(6) input')?.value || '',
      name: tr.querySelector('td:nth-child(7) input')?.value || '',
      length: tr.querySelector('td:nth-child(8) input')?.value || '',
      operator: tr.querySelector('td:nth-child(14) select')?.value || '',
      finish: getSplitTimeValue(tr, 'time-finish'),
      speed: tr.querySelector('td:nth-child(16) select')?.value || '',
      other: tr.querySelector('td:nth-child(17) input')?.value || '',
      ...checks,
    };
  }

  function applySplitTimeToRow(row, className, value) {
    if (!value) return;
    const wrap = row.querySelector(`.time-split.${className}`);
    if (!wrap) return;
    applySplitTimePaste(wrap, value);
  }

  function deserializeMainRow(tr, rowData) {
    const data = migrateLegacyMainRow(rowData);
    if (!data) return;

    applySplitTimeToRow(tr, 'time-load', data.load || '');

    applySplitTimeToRow(tr, 'time-start', data.start || '');

    const productInput = tr.querySelector('td:nth-child(3) input');
    if (productInput) productInput.value = data.productNo || '';

    setThicknessSelectValue(getThicknessSelect(tr, 4), data.thickness || '');

    const widthInput = tr.querySelector('td:nth-child(5) input');
    if (widthInput) widthInput.value = data.width || '';

    const heightInput = tr.querySelector('td:nth-child(6) input');
    if (heightInput) heightInput.value = data.height || '';

    const nameInput = tr.querySelector('td:nth-child(7) input');
    if (nameInput) nameInput.value = data.name || '';
    refreshProductNotFoundUI(productInput, nameInput);

    const lengthInput = tr.querySelector('td:nth-child(8) input');
    if (lengthInput) lengthInput.value = data.length || '';

    const operatorSelect = tr.querySelector('td:nth-child(14) select');
    if (operatorSelect) operatorSelect.value = data.operator || '';

    applySplitTimeToRow(tr, 'time-finish', data.finish || '');

    const speedSelect = tr.querySelector('td:nth-child(16) select');
    if (speedSelect) {
      speedSelect.value = data.speed || '';
      if (data.speed) {
        window.formatSpeedDisplay(speedSelect);
        if (data.speed === '轉機') speedSelect.dataset.transferRowAdded = '1';
        else speedSelect.dataset.speedRowAdded = '1';
      }
    }

    const otherInput = tr.querySelector('td:nth-child(17) input');
    if (otherInput) otherInput.value = data.other || '';

    ['length_tolerance', 'section_size', 'left_right_bend', 'up_down_bend', 'twist'].forEach((field) => {
      const checkbox = tr.querySelector(`.three-state-checkbox[data-field="${field}"]`);
      if (checkbox?.threeStateInstance && data[field]) {
        checkbox.threeStateInstance.setValue(data[field]);
      }
    });
  }

  function migrateLegacyMainRow(rowData) {
    if (rowData && typeof rowData === 'object' && !Array.isArray(rowData)) {
      return rowData;
    }
    if (!Array.isArray(rowData)) return null;

    // v2: 12 inputs + 5 three-state values
    if (rowData.length >= 12) {
      return {
        load: rowData[0] || '',
        start: rowData[1] || '',
        productNo: rowData[2] || '',
        thickness: rowData[3] || '',
        width: rowData[4] || '',
        height: rowData[5] || '',
        name: rowData[6] || '',
        length: rowData[7] || '',
        operator: rowData[8] || '',
        finish: rowData[9] || '',
        speed: rowData[10] || '',
        other: rowData[11] || '',
        length_tolerance: rowData[12] || '',
        section_size: rowData[13] || '',
        left_right_bend: rowData[14] || '',
        up_down_bend: rowData[15] || '',
        twist: rowData[16] || '',
      };
    }
    return null;
  }

  function serializeMaterialRow(tr) {
    const checkboxes = tr.querySelectorAll('td.chk input[type="checkbox"]');
    return {
      orderNo: getMaterialOrderNoValue(tr),
      thickness1: getThicknessValue(tr, 2),
      width1: tr.querySelector('td:nth-child(3) input')?.value || '',
      weight: tr.querySelector('td:nth-child(4) input')?.value || '',
      thickness2: getThicknessValue(tr, 5),
      width2: tr.querySelector('td:nth-child(6) input')?.value || '',
      height: tr.querySelector('td:nth-child(7) input')?.value || '',
      productNo: normalizeProductCodeForSubmit(tr.querySelector('td:nth-child(8) input')?.value),
      name: tr.querySelector('td:nth-child(9) input')?.value || '',
      length: tr.querySelector('td:nth-child(10) input')?.value || '',
      qty: tr.querySelector('td:nth-child(11) input')?.value || '',
      complete: !!checkboxes[0]?.checked,
      oldCoil: !!checkboxes[1]?.checked,
      incomplete: checkboxes[2]?.checked ?? true,
      mainRowIndex: tr.dataset.mainRowIndex || '',
    };
  }

  function migrateLegacyMaterialRow(rowData) {
    if (rowData && typeof rowData === 'object' && !Array.isArray(rowData)) {
      return rowData;
    }
    if (!Array.isArray(rowData) || rowData.length < 11) return null;
    return {
      orderNo: rowData[0] || '',
      thickness1: rowData[1] || '',
      width1: rowData[2] || '',
      weight: rowData[3] || '',
      thickness2: rowData[4] || '',
      width2: rowData[5] || '',
      height: rowData[6] || '',
      productNo: rowData[7] || '',
      name: rowData[8] || '',
      length: rowData[9] || '',
      qty: rowData[10] || '',
      complete: rowData[11] === 'true' || rowData[11] === true,
      oldCoil: rowData[12] === 'true' || rowData[12] === true,
      incomplete: rowData[13] !== 'false' && rowData[13] !== false,
    };
  }

  function deserializeMaterialRow(tr, rowData) {
    const data = migrateLegacyMaterialRow(rowData);
    if (!data) return;

    setMaterialOrderNoValue(tr, data.orderNo || '');
    setThicknessSelectValue(getThicknessSelect(tr, 2), data.thickness1 || '');
    tr.querySelector('td:nth-child(3) input').value = data.width1 || '';
    tr.querySelector('td:nth-child(4) input').value = data.weight || '';
    setThicknessSelectValue(getThicknessSelect(tr, 5), data.thickness2 || '');
    tr.querySelector('td:nth-child(6) input').value = data.width2 || '';
    tr.querySelector('td:nth-child(7) input').value = data.height || '';
    tr.querySelector('td:nth-child(8) input').value = data.productNo || '';
    const materialNameInput = tr.querySelector('td:nth-child(9) input');
    materialNameInput.value = data.name || '';
    refreshProductNotFoundUI(tr.querySelector('td:nth-child(8) input'), materialNameInput);
    tr.querySelector('td:nth-child(10) input').value = data.length || '';
    tr.querySelector('td:nth-child(11) input').value = data.qty || '';

    const checkboxes = tr.querySelectorAll('td.chk input[type="checkbox"]');
    if (checkboxes[0]) checkboxes[0].checked = !!data.complete;
    if (checkboxes[1]) checkboxes[1].checked = !!data.oldCoil;
    if (checkboxes[2]) {
      checkboxes[2].checked = data.incomplete !== false;
      checkboxes[2].classList.toggle('checkbox-red', checkboxes[2].checked);
    }
    if (data.mainRowIndex !== undefined && data.mainRowIndex !== '') {
      tr.dataset.mainRowIndex = String(data.mainRowIndex);
    }
  }

  const PRODUCTION_RECORDS_KEY = 'pq-form-production-records-v1';
  const PRODUCTION_RECORD_KNOWN_TYPES = [
    '企筒', '地槽', '鐵角', '批灰角', 'W角', '闊槽', 'C槽', '其他', 'CT企筒打孔',
  ];
  let productionRecordsCache = [];
  let productionRecordFilterMonth = '';
  let productionRecordFilterType = '';

  function getFormDateString(inAutoPage) {
    const y = document.getElementById(inAutoPage ? 'year2' : 'year')?.value || '';
    const m = document.getElementById(inAutoPage ? 'month2' : 'month')?.value || '';
    const d = document.getElementById(inAutoPage ? 'day2' : 'day')?.value || '';
    if (!y && !m && !d) return '';
    return `${y}/${m}/${d}`;
  }

  function getMaterialBodyForPage(inAutoPage) {
    return inAutoPage ? materialTableBody2 : materialTableBody;
  }

  function getProductionPageType(inAutoPage) {
    return inAutoPage ? 'auto' : 'molding';
  }

  function buildProductionRecord(mainTr, sheetRow, inAutoPage) {
    const rowIndex = getMainRowIndex(mainTr);
    const materialTr = findMaterialRowForMain(mainTr, inAutoPage);
    return {
      id: crypto.randomUUID(),
      recordDate: getFormDateString(inAutoPage),
      pageType: getProductionPageType(inAutoPage),
      productTypes: collectChecks(inAutoPage ? '#autoPage input[name="type"]' : '#moldingPage input[name="type"]'),
      machines: collectChecks(inAutoPage ? '#autoPage input[name="machine"]' : '#moldingPage input[name="machine"]'),
      main: serializeMainRow(mainTr),
      material: materialTr ? serializeMaterialRow(materialTr) : {},
      sheetRow: sheetRow || null,
      sheetName: 'pq-form',
      correctionNote: '',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
  }

  function buildProductionRecordPreview(mainTr, inAutoPage) {
    const materialTr = findMaterialRowForMain(mainTr, inAutoPage);
    return {
      recordDate: getFormDateString(inAutoPage),
      pageType: getProductionPageType(inAutoPage),
      main: serializeMainRow(mainTr),
      material: materialTr ? serializeMaterialRow(materialTr) : {},
    };
  }

  function getProductionRecordDuplicateKey(record) {
    const main = record?.main || {};
    const mat = record?.material || {};
    return [
      record?.recordDate || '',
      String(main.productNo || '').trim().toUpperCase(),
      formatThicknessValue(main.thickness),
      String(main.width || '').trim(),
      String(main.height || '').trim(),
      String(main.length || '').trim(),
      String(mat.productNo || '').trim().toUpperCase(),
      formatThicknessValue(mat.thickness1),
      String(mat.width1 || '').trim(),
      formatThicknessValue(mat.thickness2),
      String(mat.width2 || '').trim(),
      String(mat.height || '').trim(),
      String(mat.length || '').trim(),
      String(mat.qty || '').trim(),
    ].join('\u0001');
  }

  function findDuplicateProductionRecord(mainTr, inAutoPage) {
    const preview = buildProductionRecordPreview(mainTr, inAutoPage);
    if (!preview.recordDate || !preview.main?.productNo) return null;
    const key = getProductionRecordDuplicateKey(preview);
    const matches = productionRecordsCache.filter((record) =>
      isActiveProductionRecord(record)
      && record.recordDate === preview.recordDate
      && getProductionRecordDuplicateKey(record) === key);
    if (!matches.length) return null;
    return matches.sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0))[0];
  }

  function normalizeOrderNo(value) {
    return String(value ?? '').trim();
  }

  function findDuplicateByOrderNo(mainTr, inAutoPage) {
    const preview = buildProductionRecordPreview(mainTr, inAutoPage);
    const orderNo = normalizeOrderNo(preview.material?.orderNo);
    if (!orderNo || !preview.recordDate) return null;
    const matches = productionRecordsCache.filter((record) =>
      isActiveProductionRecord(record)
      && record.recordDate === preview.recordDate
      && normalizeOrderNo(record.material?.orderNo) === orderNo);
    if (!matches.length) return null;
    return matches.sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0))[0];
  }

  function saveProductionRecordsLocal() {
    try {
      localStorage.setItem(PRODUCTION_RECORDS_KEY, JSON.stringify(productionRecordsCache));
    } catch (e) {
      console.warn('saveProductionRecordsLocal failed', e);
    }
  }

  function loadProductionRecordsLocal() {
    try {
      const raw = localStorage.getItem(PRODUCTION_RECORDS_KEY);
      productionRecordsCache = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(productionRecordsCache)) productionRecordsCache = [];
      const deletedIds = loadDeletedProductionRecordIds();
      productionRecordsCache = productionRecordsCache.filter((r) =>
        !r.deletedAt && !deletedIds.has(r.id));
    } catch (e) {
      productionRecordsCache = [];
    }
  }

  function escapeHtml(text) {
    return String(text ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function formatRecordSubmittedAtDisplay(record) {
    const iso = record?.updatedAt || record?.createdAt;
    if (!iso) return '—';
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return '—';
    const parts = new Intl.DateTimeFormat('en-GB', {
      timeZone: HK_TIMEZONE,
      hour: '2-digit',
      minute: '2-digit',
      hour12: false,
    }).formatToParts(d);
    const pick = (type) => parts.find((p) => p.type === type)?.value || '';
    return `${pick('hour')}:${pick('minute')}`;
  }

  function isActiveProductionRecord(record) {
    return !record?.deletedAt;
  }

  function formatMonthFilterLabel(monthKey) {
    const m = String(monthKey).match(/^(\d{4})\/(\d{2})$/);
    if (!m) return monthKey;
    return `${m[1]}年${m[2]}月`;
  }

  function parseRecordMonthKey(recordDate) {
    const text = String(recordDate ?? '').trim();
    const m = text.match(/^(\d{4})[\/\-](\d{1,2})/);
    if (!m) return '';
    return `${m[1]}/${String(parseInt(m[2], 10)).padStart(2, '0')}`;
  }

  function getRecordProductTypeLabel(record) {
    const types = record?.productTypes || {};
    const selected = Object.keys(types).filter((key) => types[key] && key !== '其他入力');
    if (!selected.length) return '（未分類）';
    const key = selected[0];
    if (key === '其他') {
      const other = String(types['其他入力'] || '').trim();
      return other ? `其他：${other}` : '其他';
    }
    return key;
  }

  function getSortedActiveProductionRecords() {
    return productionRecordsCache
      .filter((r) => isActiveProductionRecord(r))
      .sort((a, b) => {
        const dateCmp = String(b.recordDate || '').localeCompare(String(a.recordDate || ''));
        if (dateCmp !== 0) return dateCmp;
        return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
      });
  }

  function getFilteredProductionRecords() {
    return getSortedActiveProductionRecords().filter((record) => {
      if (productionRecordFilterMonth && parseRecordMonthKey(record.recordDate) !== productionRecordFilterMonth) {
        return false;
      }
      if (productionRecordFilterType && getRecordProductTypeLabel(record) !== productionRecordFilterType) {
        return false;
      }
      return true;
    });
  }

  function syncProductionRecordFilterSelects() {
    document.querySelectorAll('.production-record-month-filter').forEach((select) => {
      select.value = productionRecordFilterMonth;
    });
    document.querySelectorAll('.production-record-type-filter').forEach((select) => {
      select.value = productionRecordFilterType;
    });
  }

  function refreshProductionRecordFilterOptions() {
    const records = getSortedActiveProductionRecords();
    const monthOptions = [...new Set(records.map((r) => parseRecordMonthKey(r.recordDate)).filter(Boolean))]
      .sort((a, b) => b.localeCompare(a));
    const typeOptions = [...new Set([
      ...PRODUCTION_RECORD_KNOWN_TYPES,
      ...records.map(getRecordProductTypeLabel),
    ])].filter(Boolean)
      .sort((a, b) => a.localeCompare(b, 'zh-Hant'));

    if (productionRecordFilterMonth && !monthOptions.includes(productionRecordFilterMonth)) {
      productionRecordFilterMonth = '';
    }
    if (productionRecordFilterType && !typeOptions.includes(productionRecordFilterType)) {
      productionRecordFilterType = '';
    }

    const monthHtml = ['<option value="">全部月份</option>']
      .concat(monthOptions.map((month) => `<option value="${escapeHtml(month)}">${escapeHtml(formatMonthFilterLabel(month))}</option>`))
      .join('');
    const typeHtml = ['<option value="">全部種類</option>']
      .concat(typeOptions.map((type) => `<option value="${escapeHtml(type)}">${escapeHtml(type)}</option>`))
      .join('');

    document.querySelectorAll('.production-record-month-filter').forEach((select) => {
      select.innerHTML = monthHtml;
      select.value = productionRecordFilterMonth;
    });
    document.querySelectorAll('.production-record-type-filter').forEach((select) => {
      select.innerHTML = typeHtml;
      select.value = productionRecordFilterType;
    });
  }

  function buildProductionRecordsTableBodyHtml(records, emptyMessage = '尚無生產紀錄') {
    if (records.length === 0) {
      return `<tr class="production-record-empty"><td colspan="15">${escapeHtml(emptyMessage)}</td></tr>`;
    }

    return records.map((record) => {
      const m = record.main || {};
      const mat = record.material || {};
      const corrected = !!record.correctedAt || !!record.correctionNote;
      return `<tr data-record-id="${escapeHtml(record.id)}" class="${corrected ? 'is-corrected' : ''}">
        <td>${escapeHtml(record.recordDate)}</td>
        <td>${escapeHtml(formatRecordSubmittedAtDisplay(record))}</td>
        <td>${escapeHtml(normalizeProductCodeForSubmit(m.productNo))}</td>
        <td>${escapeHtml(m.name)}</td>
        <td>${escapeHtml(m.thickness)}</td>
        <td>${escapeHtml(m.width)}</td>
        <td>${escapeHtml(m.height)}</td>
        <td>${escapeHtml(m.length)}</td>
        <td>${escapeHtml(m.operator)}</td>
        <td>${escapeHtml(m.load)}</td>
        <td>${escapeHtml(m.start)}</td>
        <td>${escapeHtml(m.finish)}</td>
        <td>${escapeHtml(mat.orderNo)}</td>
        <td>${escapeHtml(mat.qty)}</td>
        <td class="production-record-actions">
          <button type="button" class="btn btn-secondary btn-edit-production-record">修正</button>
          <button type="button" class="btn btn-secondary btn-delete-production-record">刪除</button>
        </td>
      </tr>`;
    }).join('');
  }

  function renderProductionRecordsTable() {
    const allRecords = getSortedActiveProductionRecords();
    const records = getFilteredProductionRecords();
    const countText = (productionRecordFilterMonth || productionRecordFilterType)
      ? `${records.length} 筆（共 ${allRecords.length} 筆）`
      : `${records.length} 筆`;
    const emptyMessage = (productionRecordFilterMonth || productionRecordFilterType) && records.length === 0
      ? '沒有符合篩選條件的生產紀錄'
      : '尚無生產紀錄';
    const bodyHtml = buildProductionRecordsTableBodyHtml(records, emptyMessage);

    [document.getElementById('productionRecordBody'), document.getElementById('productionRecordBody2')].forEach((tbody) => {
      if (tbody) tbody.innerHTML = bodyHtml;
    });
    [document.getElementById('productionRecordCount'), document.getElementById('productionRecordCount2')].forEach((el) => {
      if (el) el.textContent = countText;
    });
  }

  function renderAllProductionRecords() {
    refreshProductionRecordFilterOptions();
    renderProductionRecordsTable();
  }

  function upsertProductionRecord(record) {
    const idx = productionRecordsCache.findIndex((r) => r.id === record.id);
    if (idx >= 0) productionRecordsCache[idx] = record;
    else productionRecordsCache.unshift(record);
    saveProductionRecordsLocal();
    renderAllProductionRecords();
  }

  const PRODUCTION_DELETED_IDS_KEY = 'pq-form-production-deleted-ids-v1';

  function loadDeletedProductionRecordIds() {
    try {
      const raw = localStorage.getItem(PRODUCTION_DELETED_IDS_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      return new Set(Array.isArray(arr) ? arr.filter(Boolean) : []);
    } catch (_) {
      return new Set();
    }
  }

  function persistDeletedProductionRecordId(recordId) {
    if (!recordId) return;
    const ids = loadDeletedProductionRecordIds();
    ids.add(recordId);
    localStorage.setItem(PRODUCTION_DELETED_IDS_KEY, JSON.stringify([...ids]));
  }

  function removeProductionRecordFromCache(recordId) {
    if (!recordId) return;
    productionRecordsCache = productionRecordsCache.filter((r) => r.id !== recordId);
    persistDeletedProductionRecordId(recordId);
    saveProductionRecordsLocal();
    renderAllProductionRecords();
  }

  async function fetchProductionRecordsFromServer() {
    try {
      const res = await fetch(`${API_BASE}/api/pq_form/production_records`, { cache: 'no-store' });
      const data = await res.json();
      if (!data.success) return false;
      const serverRecords = data.records || [];
      const serverIds = new Set(serverRecords.map((r) => r.id));
      const deletedIds = loadDeletedProductionRecordIds();
      const localActiveUnsynced = productionRecordsCache.filter((r) =>
        !r.deletedAt && !serverIds.has(r.id) && !deletedIds.has(r.id));
      const merged = [...serverRecords, ...localActiveUnsynced];
      const seen = new Set();
      productionRecordsCache = merged.filter((r) => {
        if (!r.id || seen.has(r.id)) return false;
        seen.add(r.id);
        return true;
      });
      saveProductionRecordsLocal();
      renderAllProductionRecords();
      return true;
    } catch (e) {
      console.warn('fetchProductionRecordsFromServer failed', e);
      return false;
    }
  }

  async function saveProductionRecordToServer(record) {
    try {
      const res = await fetch(`${API_BASE}/api/pq_form/production_records`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        cache: 'no-store',
        body: JSON.stringify(record),
      });
      const data = await res.json();
      if (data.success && data.record) return data.record;
    } catch (e) {
      console.warn('saveProductionRecordToServer failed', e);
    }
    return null;
  }

  async function refreshProductionRecords() {
    loadProductionRecordsLocal();
    const ok = await fetchProductionRecordsFromServer();
    if (!ok) renderAllProductionRecords();
  }

  const productionRecordModal = document.getElementById('productionRecordModal');
  const productionRecordEditForm = document.getElementById('productionRecordEditForm');
  const productionRecordEditError = document.getElementById('productionRecordEditError');

  function closeProductionRecordModal() {
    if (productionRecordModal) productionRecordModal.hidden = true;
    if (productionRecordEditError) productionRecordEditError.hidden = true;
  }

  function openProductionRecordModal(recordId) {
    const record = productionRecordsCache.find((r) => r.id === recordId);
    if (!record || !productionRecordModal) return;
    const m = record.main || {};
    const mat = record.material || {};
    document.getElementById('editRecordId').value = record.id;
    document.getElementById('editMainLoad').value = m.load || '';
    document.getElementById('editMainStart').value = m.start || '';
    document.getElementById('editMainFinish').value = m.finish || '';
    document.getElementById('editMainProductNo').value = m.productNo || '';
    document.getElementById('editMainThickness').value = m.thickness || '';
    document.getElementById('editMainWidth').value = m.width || '';
    document.getElementById('editMainHeight').value = m.height || '';
    document.getElementById('editMainLength').value = m.length || '';
    document.getElementById('editMainName').value = m.name || '';
    document.getElementById('editMainOperator').value = m.operator || '';
    document.getElementById('editMainSpeed').value = m.speed || '';
    document.getElementById('editMainOther').value = m.other || '';
    document.getElementById('editMatOrderNo').value = mat.orderNo || '';
    document.getElementById('editMatThickness1').value = mat.thickness1 || '';
    document.getElementById('editMatWidth1').value = mat.width1 || '';
    document.getElementById('editMatWeight').value = mat.weight || '';
    document.getElementById('editMatThickness2').value = mat.thickness2 || '';
    document.getElementById('editMatWidth2').value = mat.width2 || '';
    document.getElementById('editMatHeight').value = mat.height || '';
    document.getElementById('editMatProductNo').value = mat.productNo || '';
    document.getElementById('editMatName').value = mat.name || '';
    document.getElementById('editMatLength').value = mat.length || '';
    document.getElementById('editMatQty').value = mat.qty || '';
    document.getElementById('editCorrectionNote').value = record.correctionNote || '';
    productionRecordEditError.hidden = true;
    productionRecordModal.hidden = false;
  }

  async function handleProductionRecordEditSubmit(e) {
    e.preventDefault();
    const id = document.getElementById('editRecordId').value;
    const record = productionRecordsCache.find((r) => r.id === id);
    if (!record) return;

    const correctionNote = document.getElementById('editCorrectionNote').value.trim();
    if (!correctionNote) {
      productionRecordEditError.textContent = '請填寫修正理由';
      productionRecordEditError.hidden = false;
      return;
    }

    const main = {
      ...record.main,
      load: document.getElementById('editMainLoad').value,
      start: document.getElementById('editMainStart').value,
      finish: document.getElementById('editMainFinish').value,
      productNo: document.getElementById('editMainProductNo').value,
      thickness: document.getElementById('editMainThickness').value,
      width: document.getElementById('editMainWidth').value,
      height: document.getElementById('editMainHeight').value,
      length: document.getElementById('editMainLength').value,
      name: document.getElementById('editMainName').value,
      operator: document.getElementById('editMainOperator').value,
      speed: document.getElementById('editMainSpeed').value,
      other: document.getElementById('editMainOther').value,
    };
    const material = {
      ...record.material,
      orderNo: document.getElementById('editMatOrderNo').value,
      thickness1: document.getElementById('editMatThickness1').value,
      width1: document.getElementById('editMatWidth1').value,
      weight: document.getElementById('editMatWeight').value,
      thickness2: document.getElementById('editMatThickness2').value,
      width2: document.getElementById('editMatWidth2').value,
      height: document.getElementById('editMatHeight').value,
      productNo: document.getElementById('editMatProductNo').value,
      name: document.getElementById('editMatName').value,
      length: document.getElementById('editMatLength').value,
      qty: document.getElementById('editMatQty').value,
    };

    try {
      const res = await fetch(`${API_BASE}/api/pq_form/production_records/${encodeURIComponent(id)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        cache: 'no-store',
        body: JSON.stringify({ main, material, correction_note: correctionNote }),
      });
      const data = await res.json();
      if (res.status === 503) {
        upsertProductionRecord({
          ...record,
          main,
          material,
          correctionNote,
          correctedAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        });
        closeProductionRecordModal();
        return;
      }
      if (!data.success) {
        productionRecordEditError.textContent = data.error || '保存失敗';
        productionRecordEditError.hidden = false;
        return;
      }
      upsertProductionRecord(data.record);
      closeProductionRecordModal();
    } catch (err) {
      productionRecordEditError.textContent = '保存失敗';
      productionRecordEditError.hidden = false;
    }
  }

  function bindProductionRecordEvents() {
    document.getElementById('refreshProductionRecordsBtn')?.addEventListener('click', refreshProductionRecords);
    document.getElementById('refreshProductionRecordsBtn2')?.addEventListener('click', refreshProductionRecords);
    document.addEventListener('change', (e) => {
      if (e.target.matches('.production-record-month-filter')) {
        productionRecordFilterMonth = e.target.value;
        syncProductionRecordFilterSelects();
        renderProductionRecordsTable();
        return;
      }
      if (e.target.matches('.production-record-type-filter')) {
        productionRecordFilterType = e.target.value;
        syncProductionRecordFilterSelects();
        renderProductionRecordsTable();
      }
    });
    document.getElementById('productionRecordModalClose')?.addEventListener('click', closeProductionRecordModal);
    document.getElementById('productionRecordModalCancel')?.addEventListener('click', closeProductionRecordModal);
    productionRecordEditForm?.addEventListener('submit', handleProductionRecordEditSubmit);
    document.addEventListener('click', (e) => {
      const editBtn = e.target.closest('.btn-edit-production-record');
      if (editBtn) {
        const id = editBtn.closest('tr')?.dataset.recordId;
        if (id) openProductionRecordModal(id);
        return;
      }
      const deleteBtn = e.target.closest('.btn-delete-production-record');
      if (deleteBtn) {
        const id = deleteBtn.closest('tr')?.dataset.recordId;
        if (id) confirmAndDeleteProductionRecord(id);
      }
    });
    ['year', 'month', 'day', 'year2', 'month2', 'day2'].forEach((id) => {
      document.getElementById(id)?.addEventListener('change', renderAllProductionRecords);
      document.getElementById(id)?.addEventListener('input', renderAllProductionRecords);
    });
  }

  const deleteRecordModal = document.getElementById('deleteRecordModal');
  const deleteRecordModalSummary = document.getElementById('deleteRecordModalSummary');
  let deleteConfirmResolver = null;

  function closeDeleteConfirmModal() {
    if (deleteRecordModal) deleteRecordModal.hidden = true;
  }

  function showDeleteConfirmModal(record) {
    const m = record?.main || {};
    const mat = record?.material || {};
    if (deleteRecordModalSummary) {
      deleteRecordModalSummary.innerHTML = `
        <strong>將刪除嘅紀錄</strong>
        期日：${escapeHtml(record.recordDate)}<br>
        產品：${escapeHtml(m.productNo)} ${escapeHtml(m.name)}<br>
        規格：${escapeHtml(formatThicknessValue(m.thickness))} × ${escapeHtml(m.width)} × ${escapeHtml(m.height)} × ${escapeHtml(m.length)}<br>
        單號：${escapeHtml(mat.orderNo || '（空）')}<br>
        用料數量：${escapeHtml(mat.qty || '（空）')}
      `;
    }
    if (deleteRecordModal) deleteRecordModal.hidden = false;
    return new Promise((resolve) => {
      deleteConfirmResolver = resolve;
    });
  }

  function resolveDeleteConfirm(confirmed) {
    closeDeleteConfirmModal();
    if (deleteConfirmResolver) {
      deleteConfirmResolver(confirmed);
      deleteConfirmResolver = null;
    }
  }

  function bindDeleteConfirmEvents() {
    document.getElementById('deleteRecordConfirmBtn')?.addEventListener('click', () => resolveDeleteConfirm(true));
    document.getElementById('deleteRecordCancelBtn')?.addEventListener('click', () => resolveDeleteConfirm(false));
  }

  function showProductionRecordMessage(message, kind = 'error') {
    document.querySelectorAll('.production-record-wrap').forEach((wrap) => {
      const panel = getHintPanel(wrap);
      if (!message) {
        panel.hidden = true;
        panel.textContent = '';
        panel.classList.remove('product-hint-outside--success', 'product-hint-outside--error');
        return;
      }
      panel.textContent = message;
      panel.hidden = false;
      panel.classList.remove('product-hint-outside--success', 'product-hint-outside--error');
      panel.classList.add(kind === 'success' ? 'product-hint-outside--success' : 'product-hint-outside--error');
    });
  }

  async function softDeleteProductionRecord(recordId) {
    const record = productionRecordsCache.find((r) => r.id === recordId);
    if (!record || record.deletedAt) return { ok: false, error: '紀錄不存在' };

    try {
      const res = await fetch(`${API_BASE}/api/pq_form/production_records/${encodeURIComponent(recordId)}`, {
        method: 'DELETE',
        cache: 'no-store',
      });
      let data = {};
      try {
        data = await res.json();
      } catch (_) {}

      if (res.status === 503) {
        removeProductionRecordFromCache(recordId);
        showProductionRecordMessage('已刪除（本機紀錄）', 'success');
        return { ok: true };
      }
      if (res.ok && data.success) {
        removeProductionRecordFromCache(recordId);
        showProductionRecordMessage('已刪除', 'success');
        return { ok: true };
      }
      const error = data.error || '刪除失敗，請稍後再試';
      const missingOnServer = res.status === 410
        || /single JSON object|not found|already deleted|No rows/i.test(error);
      if (missingOnServer) {
        removeProductionRecordFromCache(recordId);
        showProductionRecordMessage('已刪除', 'success');
        return { ok: true };
      }
      console.warn('softDeleteProductionRecord failed', error);
      return { ok: false, error };
    } catch (e) {
      console.warn('softDeleteProductionRecord failed', e);
      removeProductionRecordFromCache(recordId);
      showProductionRecordMessage('已刪除（本機紀錄）', 'success');
      return { ok: true };
    }
  }

  async function confirmAndDeleteProductionRecord(recordId) {
    const record = productionRecordsCache.find((r) => r.id === recordId);
    if (!record || record.deletedAt) return;
    showProductionRecordMessage('', 'success');
    const confirmed = await showDeleteConfirmModal(record);
    if (!confirmed) return;
    const result = await softDeleteProductionRecord(recordId);
    if (!result.ok) {
      showProductionRecordMessage(result.error || '刪除失敗，請稍後再試', 'error');
    }
  }

  const duplicateRecordModal = document.getElementById('duplicateRecordModal');
  const duplicateRecordModalSummary = document.getElementById('duplicateRecordModalSummary');
  let duplicateConfirmResolver = null;

  function closeDuplicateConfirmModal() {
    if (duplicateRecordModal) duplicateRecordModal.hidden = true;
  }

  function showDuplicateConfirmModal(existingRecord, reason = 'spec') {
    const m = existingRecord?.main || {};
    const mat = existingRecord?.material || {};
    const msgEl = document.getElementById('duplicateRecordModalMessage');
    if (msgEl) {
      if (reason === 'orderNo') {
        const orderNo = mat.orderNo || '（空）';
        msgEl.textContent = `同一單號（${orderNo}）嘅生產紀錄已經存在。你想覆蓋舊紀錄定係另起新紀錄送出？`;
      } else {
        msgEl.textContent = '同日期、同產品、同規格、同用料同數量嘅紀錄已經存在。你想點做？';
      }
    }
    if (duplicateRecordModalSummary) {
      duplicateRecordModalSummary.innerHTML = `
        <strong>已有紀錄</strong>
        期日：${escapeHtml(existingRecord.recordDate)}<br>
        產品：${escapeHtml(m.productNo)} ${escapeHtml(m.name)}<br>
        規格：${escapeHtml(formatThicknessValue(m.thickness))} × ${escapeHtml(m.width)} × ${escapeHtml(m.height)} × ${escapeHtml(m.length)}<br>
        單號：${escapeHtml(mat.orderNo || '（空）')}<br>
        用料數量：${escapeHtml(mat.qty || '（空）')}<br>
        上料／開始／完成：${escapeHtml(m.load || '—')}／${escapeHtml(m.start || '—')}／${escapeHtml(m.finish || '—')}
      `;
    }
    if (duplicateRecordModal) duplicateRecordModal.hidden = false;
    return new Promise((resolve) => {
      duplicateConfirmResolver = resolve;
    });
  }

  function resolveDuplicateConfirm(choice) {
    closeDuplicateConfirmModal();
    if (duplicateConfirmResolver) {
      duplicateConfirmResolver(choice);
      duplicateConfirmResolver = null;
    }
  }

  function bindDuplicateConfirmEvents() {
    document.getElementById('duplicateRecordOverwriteBtn')?.addEventListener('click', () => resolveDuplicateConfirm('overwrite'));
    document.getElementById('duplicateRecordNewBtn')?.addEventListener('click', () => resolveDuplicateConfirm('new'));
    document.getElementById('duplicateRecordCancelBtn')?.addEventListener('click', () => resolveDuplicateConfirm('cancel'));
  }

  function buildMappedRowFromTr(tr, isMaterialTable) {
    const cells = [...tr.querySelectorAll('input,select')];
    let tLoad, tStart, productNo, thickness, roundness, height, name, lengthVal;

    if (isMaterialTable) {
      tLoad = cells[0]?.value || '';
      tStart = '';
      productNo = normalizeProductCodeForSubmit(cells[7]?.value);
      thickness = cells[1]?.value || '';
      roundness = cells[2]?.value || '';
      height = cells[6]?.value || '';
      name = cells[8]?.value || '';
      lengthVal = cells[9]?.value || '';
    } else {
      tLoad = getSplitTimeValue(tr, 'time-load');
      tStart = getSplitTimeValue(tr, 'time-start');
      productNo = normalizeProductCodeForSubmit(tr.querySelector('td:nth-child(3) input')?.value);
      thickness = getThicknessValue(tr, 4);
      roundness = tr.querySelector('td:nth-child(5) input')?.value || '';
      height = tr.querySelector('td:nth-child(6) input')?.value || '';
      name = tr.querySelector('td:nth-child(7) input')?.value || '';
      lengthVal = tr.querySelector('td:nth-child(8) input')?.value || '';
    }

    const getThreeStateValue = (fieldName) => {
      const checkbox = tr.querySelector(`.three-state-checkbox[data-field="${fieldName}"]`);
      if (checkbox && checkbox.threeStateInstance) {
        return checkbox.threeStateInstance.getValue();
      }
      return '';
    };

    let chkLenTol, chkCutDim, chkLeftRight, chkUpDown, chkTwist, operator, tFinish, note, other;

    if (isMaterialTable) {
      chkLenTol = toTF((cells[11]?.type === 'checkbox') ? !!cells[11].checked : false);
      chkCutDim = toTF((cells[12]?.type === 'checkbox') ? !!cells[12].checked : false);
      chkLeftRight = toTF((cells[13]?.type === 'checkbox') ? !!cells[13].checked : false);
      chkUpDown = 'FALSE';
      chkTwist = 'FALSE';
      operator = '';
      tFinish = '';
      note = '';
      other = '';
    } else {
      chkLenTol = getThreeStateValue('length_tolerance');
      chkCutDim = getThreeStateValue('section_size');
      chkLeftRight = getThreeStateValue('left_right_bend');
      chkUpDown = getThreeStateValue('up_down_bend');
      chkTwist = getThreeStateValue('twist');
      operator = tr.querySelector('td:nth-child(14) select')?.value || '';
      tFinish = getSplitTimeValue(tr, 'time-finish');
      note = tr.querySelector('td:nth-child(16) select')?.value || '';
      other = tr.querySelector('td:nth-child(17) input')?.value || '';
    }

    return [
      tLoad, tStart, productNo, thickness, roundness, height,
      name,
      '', '', '',
      lengthVal,
      chkLenTol,
      chkCutDim,
      chkLeftRight,
      chkUpDown,
      chkTwist,
      operator,
      tFinish,
      note,
      other,
    ];
  }

  async function overwriteExistingProductionRecord(tr, inAutoPage, existingRecord, sheetRow) {
    const main = serializeMainRow(tr);
    const materialTr = findMaterialRowForMain(tr, inAutoPage);
    const material = materialTr ? serializeMaterialRow(materialTr) : (existingRecord.material || {});
    const fallbackRecord = {
      ...existingRecord,
      main,
      material,
      sheetRow: sheetRow || existingRecord.sheetRow || null,
      productTypes: collectChecks(inAutoPage ? '#autoPage input[name="type"]' : '#moldingPage input[name="type"]'),
      machines: collectChecks(inAutoPage ? '#autoPage input[name="machine"]' : '#moldingPage input[name="machine"]'),
      correctionNote: '重複送出覆蓋',
      correctedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    try {
      const res = await fetch(`${API_BASE}/api/pq_form/production_records/${encodeURIComponent(existingRecord.id)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        cache: 'no-store',
        body: JSON.stringify({
          main,
          material,
          correction_note: '重複送出覆蓋',
        }),
      });
      const data = await res.json();
      if (res.status === 503) {
        upsertProductionRecord(fallbackRecord);
        return;
      }
      if (data.success && data.record) {
        upsertProductionRecord(data.record);
        return;
      }
    } catch (e) {
      console.warn('overwriteExistingProductionRecord failed', e);
    }
    upsertProductionRecord(fallbackRecord);
  }

  async function performRowSend(mainTr, inAutoPage, options = {}) {
    const { overwriteRecord = null } = options;

    const headerPayload = {
      date: {
        y: document.getElementById(inAutoPage ? 'year2' : 'year').value,
        m: document.getElementById(inAutoPage ? 'month2' : 'month').value,
        d: document.getElementById(inAutoPage ? 'day2' : 'day').value,
      },
      types: collectChecks(inAutoPage ? '#autoPage input[name="type"]' : '#moldingPage input[name="type"]'),
      machines: collectChecks(inAutoPage ? '#autoPage input[name="machine"]' : '#moldingPage input[name="machine"]'),
    };
    fetch(`${API_BASE}/api/pq_form/update_header`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      cache: 'no-store',
      body: JSON.stringify(headerPayload),
    }).catch(() => {});

    const mapped = buildMappedRowFromTr(mainTr, false);
    const payload = { rows: [mapped] };
    if (overwriteRecord?.sheetRow) payload.targetRow = overwriteRecord.sheetRow;

    try {
      const res = await fetch(`${API_BASE}/api/pq_form/submit`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        cache: 'no-store',
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      console.log('submit result', data);
      if (!data.success) {
        showOutsideMessage(mainTr, '送出失敗: ' + (data.error || ''), 'error');
        return;
      }

      if (overwriteRecord) {
        showOutsideMessage(mainTr, '已覆蓋舊紀錄（伺服器寫入）', 'success');
        await overwriteExistingProductionRecord(mainTr, inAutoPage, overwriteRecord, data.row);
        return;
      }

      showOutsideMessage(mainTr, '已送出（伺服器寫入）', 'success');
      const record = buildProductionRecord(mainTr, data.row, inAutoPage);
      const saved = await saveProductionRecordToServer(record);
      upsertProductionRecord(saved || record);
    } catch (err) {
      console.error(err);
      showOutsideMessage(mainTr, '送出失敗', 'error');
    }
  }

  function persistLocal(){
    try{
      const data = {
        y: document.getElementById('year').value,
        m: document.getElementById('month').value,
        d: document.getElementById('day').value,
        rows: [...tableBody.querySelectorAll('tr')].filter(isMainDataRow).map((tr) => serializeMainRow(tr)),
        materialRows: [...materialTableBody.querySelectorAll('tr')]
          .filter((tr) => !isMaterialRowTemplateEmpty(tr))
          .map((tr) => serializeMaterialRow(tr)),
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    }catch(e){/* noop */}
  }

  // 產品名稱の列幅（上下テーブル共通・固定250px）
  function adjustNameColumnWidth(){
    const width = 250;
    document.querySelectorAll('thead th:nth-child(7)').forEach(th => {
      if (th.textContent.trim() === '產品名稱') th.style.width = width + 'px';
    });
    document.querySelectorAll('#tableBody td.name, #tableBody2 td.name').forEach(td => {
      td.style.width = width + 'px';
    });
    document.querySelectorAll('#materialTableBody td:nth-child(9), #materialTableBody2 td:nth-child(9)').forEach(td => {
      td.style.width = width + 'px';
    });
    document.querySelectorAll('#materialTableBody, #materialTableBody2').forEach(tbody => {
      const th = tbody.closest('table')?.querySelector('thead th:nth-child(9)');
      if (th) th.style.width = width + 'px';
    });
  }

  function initAutoPageRows() {
    if (tableBody2 && ![...tableBody2.querySelectorAll('tr')].some(isMainDataRow)) {
      addRow(1, tableBody2);
    }
  }

  function restoreLocal(){
    try{
      const raw = localStorage.getItem(STORAGE_KEY) || localStorage.getItem('pq-form-ui-v2');
      if(!raw){ 
        addRow(1);
        return; 
      }
      const data = JSON.parse(raw);
      tableBody.innerHTML='';
      (data.rows||[]).forEach((r)=>{
        const tr = createRow();
        deserializeMainRow(tr, r);
        tableBody.appendChild(tr);
      });

      materialTableBody.innerHTML='';
      (data.materialRows||[]).forEach((r)=>{
        const tr = createMaterialRow();
        deserializeMaterialRow(tr, r);
        if (!isMaterialRowTemplateEmpty(tr)) {
          materialTableBody.appendChild(tr);
        }
      });
      
      if(!data.rows || data.rows.length === 0) {
        addRow(1);
      }
      
    }catch(e){ 
      addRow(1);
    }
    // 初期ロード後に列幅調整
    adjustNameColumnWidth();
  }

  // Events
  addRowBtn.addEventListener('click', ()=> addRow(1));
  removeRowBtn.addEventListener('click', () => removeRow());
  clearBtn.addEventListener('click', () => clearAllForPage(false));
  
  // 用料記録ボタンのイベントリスナー
  addMaterialRowBtn.addEventListener('click', ()=>addMaterialRow(1));
  removeMaterialRowBtn.addEventListener('click', () => removeMaterialRow());
  clearMaterialBtn.addEventListener('click', clearMaterialAll);

  // 用料記録行の送出ボタン
  document.addEventListener('click', (e)=>{
    const target = e.target;
    if (!target.classList.contains('btn-row-send')) return;
      const tr = target.closest('tr');
      if (!tr?.closest('#materialTableBody, #materialTableBody2')) return;
      const inAutoPage = tr.closest('#autoPage') !== null;
      const { mainTr, materialTr } = getMainMaterialRowPair(tr, inAutoPage);

      if (!mainTr) {
        showOutsideMessage(tr, '搵唔到對應嘅上段記錄', 'error');
        return;
      }

      (async () => {
        if (!validateMachineBeforeSend(inAutoPage)) return;
        if (!validateRowBeforeSend(mainTr)) return;
        if (!validateMaterialRowBeforeSend(materialTr, materialTr || mainTr)) return;

        await fetchProductionRecordsFromServer();

        const orderDuplicate = findDuplicateByOrderNo(mainTr, inAutoPage);
        if (orderDuplicate) {
          const orderChoice = await showDuplicateConfirmModal(orderDuplicate, 'orderNo');
          if (orderChoice === 'cancel') return;
          if (orderChoice === 'overwrite') {
            await performRowSend(mainTr, inAutoPage, { overwriteRecord: orderDuplicate });
            return;
          }
        }

        const specDuplicate = findDuplicateProductionRecord(mainTr, inAutoPage);
        if (specDuplicate) {
          const specChoice = await showDuplicateConfirmModal(specDuplicate, 'spec');
          if (specChoice === 'cancel') return;
          if (specChoice === 'overwrite') {
            await performRowSend(mainTr, inAutoPage, { overwriteRecord: specDuplicate });
            return;
          }
        }

        await performRowSend(mainTr, inAutoPage);
      })();
  });

  function collectChecks(selector){
    const obj = {};
    document.querySelectorAll(selector).forEach(el=>{
      obj[el.value] = !!el.checked;
    });
    if(selector.includes('type')){
      const other = document.getElementById('typeOther');
      obj['其他入力'] = other ? other.value : '';
    }
    return obj;
  }

  document.addEventListener('input', (e)=>{
    if(e.target.matches('input,select')) persistLocal();
    if(e.target.closest('td.name')) adjustNameColumnWidth();
  });

  bindSingleTypeSelection(getMoldingPageRoot());
  bindSingleTypeSelection(document.getElementById('autoPage'));
  bindMachineSelectHints(getMoldingPageRoot());
  bindMachineSelectHints(document.getElementById('autoPage'));
  bindPlistResolveInputs(document);
  bindTimeSplitInputs(document);
  bindMaterialOrderNoInputs(document);

  // Init
  async function initApp() {
    await loadThicknessOptions();
    restoreLocal();
    initAutoPageRows();
    refreshThicknessSelects();
    setToday();
    loadProductionRecordsLocal();
    renderAllProductionRecords();
    bindProductionRecordEvents();
    bindDuplicateConfirmEvents();
    bindDeleteConfirmEvents();
    bindTransferRowConfirmEvents();
    bindMachineResetConfirmEvents();
    [getMoldingPageRoot(), document.getElementById('autoPage')].forEach((pageRoot) => {
      if (!pageRoot) return;
      lastMachineByPage.set(pageRoot, normalizeMachineSelection(pageRoot));
    });
    refreshProductionRecords();
  }
  initApp();

  window.addEventListener('pageshow', () => {
    setToday();
  });

  // 初期計測（フォント読み込み後）
  window.addEventListener('load', ()=>{
    adjustNameColumnWidth();
  });
  window.addEventListener('resize', adjustNameColumnWidth);

  // ページ切り替え機能
  const moldingPage = getMoldingPageRoot();
  const autoPage = document.getElementById('autoPage');
  const moldingBtn = document.getElementById('moldingBtn');
  const autoBtn = document.getElementById('autoBtn');
  const moldingBtn2 = document.getElementById('moldingBtn2');
  const autoBtn2 = document.getElementById('autoBtn2');

  function switchToMolding() {
    moldingPage.style.display = 'block';
    autoPage.style.display = 'none';
    moldingBtn.classList.add('active');
    autoBtn.classList.remove('active');
    moldingBtn2.classList.add('active');
    autoBtn2.classList.remove('active');
  }

  function switchToAuto() {
    moldingPage.style.display = 'none';
    autoPage.style.display = 'block';
    moldingBtn.classList.remove('active');
    autoBtn.classList.add('active');
    moldingBtn2.classList.remove('active');
    autoBtn2.classList.add('active');
  }

  // イベントリスナーを追加
  moldingBtn.addEventListener('click', switchToMolding);
  autoBtn.addEventListener('click', switchToAuto);
  moldingBtn2.addEventListener('click', switchToMolding);
  autoBtn2.addEventListener('click', switchToAuto);

  // 第2ページ用の機能を初期化
  if (autoPage) {
    // 第2ページの初期行数は、restoreLocal()で既に設定済み
    // 第2ページのイベントリスナーを追加
    if (addRowBtn2) addRowBtn2.addEventListener('click', () => addRow(1, tableBody2));
    if (removeRowBtn2) removeRowBtn2.addEventListener('click', () => removeRow(tableBody2));
    if (clearBtn2) clearBtn2.addEventListener('click', () => clearAllForPage(true));
    if (addMaterialRowBtn2) addMaterialRowBtn2.addEventListener('click', () => addMaterialRow(1, materialTableBody2));
    if (removeMaterialRowBtn2) removeMaterialRowBtn2.addEventListener('click', () => removeMaterialRow(materialTableBody2));
    if (clearMaterialBtn2) clearMaterialBtn2.addEventListener('click', () => {
      materialTableBody2.innerHTML='';
    });

    console.log('Auto page initialized');
  }
})();


