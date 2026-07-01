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
  const saveLocalBtn = document.getElementById('saveLocalBtn');
  const loadByDateBtn = document.getElementById('loadByDateBtn');
  const addMaterialRowBtn = document.getElementById('addMaterialRowBtn');
  const removeMaterialRowBtn = document.getElementById('removeMaterialRowBtn');
  const clearMaterialBtn = document.getElementById('clearMaterialBtn');
  // 第2ページ用の要素
  const addRowBtn2 = document.getElementById('addRowBtn2');
  const removeRowBtn2 = document.getElementById('removeRowBtn2');
  const clearBtn2 = document.getElementById('clearBtn2');
  const saveLocalBtn2 = document.getElementById('saveLocalBtn2');
  const loadByDateBtn2 = document.getElementById('loadByDateBtn2');
  const addMaterialRowBtn2 = document.getElementById('addMaterialRowBtn2');
  const removeMaterialRowBtn2 = document.getElementById('removeMaterialRowBtn2');
  const clearMaterialBtn2 = document.getElementById('clearMaterialBtn2');

  function pad(n){return n.toString().padStart(2,'0');}

  function toTF(v){ return v ? 'TRUE' : 'FALSE'; }

  const FALLBACK_THICKNESS_OPTIONS = ['0.3', '0.4', '0.5', '0.6', '0.8', '1', '1.2', '1.5', '3'];
  let thicknessOptions = [...FALLBACK_THICKNESS_OPTIONS];

  function buildThicknessSelectOptions(selected = '') {
    const options = ['<option value=""></option>'];
    thicknessOptions.forEach((t) => {
      options.push(`<option value="${t}"${selected === t ? ' selected' : ''}>${t}</option>`);
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
    const v = String(value ?? '').trim();
    if (v && !thicknessOptions.includes(v) && !select.querySelector(`option[value="${v}"]`)) {
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = v;
      select.appendChild(opt);
    }
    select.value = v;
  }

  function refreshThicknessSelects() {
    document.querySelectorAll('select.thickness-select').forEach((sel) => {
      const current = sel.value;
      sel.innerHTML = buildThicknessSelectOptions();
      setThicknessSelectValue(sel, current);
    });
  }

  async function loadThicknessOptions() {
    try {
      const res = await fetch(`${API_BASE}/api/pq_form/plist/thicknesses`, { cache: 'no-store' });
      const data = await res.json();
      if (data.success && Array.isArray(data.thicknesses) && data.thicknesses.length) {
        thicknessOptions = data.thicknesses;
        refreshThicknessSelects();
      }
    } catch (error) {
      console.warn('loadThicknessOptions failed, using fallback', error);
    }
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

  function setToday(){
    const now = new Date();
    document.getElementById('year').value = now.getFullYear();
    document.getElementById('month').value = pad(now.getMonth()+1);
    document.getElementById('day').value = pad(now.getDate());
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
          <option value="燕">燕</option>
          <option value="年">年</option>
        </select>
      </td>
      <td>${timeSplitHtml('time-finish')}</td>
      <td>
        <select onchange="window.formatSpeedDisplay(this)">
          <option value=""></option>
          <option value="轉機">轉機</option>
          ${Array.from({length: 25}, (_, i) => i * 5).map(speed => 
            `<option value="${speed}">${speed}</option>`
          ).join('')}
        </select>
      </td>
      <td><input type="text" placeholder="其他" /></td>
      <td class="w-160 btns">
        <button class="btn btn-secondary btn-row-save" type="button">暫存</button>
        <button class="btn btn-primary btn-row-send" type="button">送出</button>
      </td>
    `;
    
    // 3段階チェックボックスを初期化
    const checkboxes = tr.querySelectorAll('.three-state-checkbox');
    checkboxes.forEach(checkbox => {
      checkbox.threeStateInstance = new ThreeStateCheckbox(checkbox);
    });
    
    return tr;
  }

  function addRow(n=1){
    for(let i=0;i<n;i++) tableBody.appendChild(createRow());
    persistLocal();
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

  function createMaterialRow(){
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="text" /></td>
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
        <button class="btn btn-secondary btn-row-save" type="button">暫存</button>
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

  // 速度表示フォーマット関数
  window.formatSpeedDisplay = function(selectElement) {
    const value = selectElement.value;
    if (value && value !== '') {
      const selectedOption = selectElement.options[selectElement.selectedIndex];
      if (value === '轉機') {
        // 轉機の場合はそのまま表示
        selectedOption.textContent = '轉機';
      } else {
        // 数字の場合は「速XX」に変更
        selectedOption.textContent = `速${value}`;
      }
      selectElement.style.background = '#f8f9fa';
      selectElement.style.fontWeight = '600';
      selectElement.style.color = 'var(--primary)';
    } else {
      // デフォルトオプションを元に戻す
      const defaultOption = selectElement.options[0];
      defaultOption.textContent = '';
      selectElement.style.background = '#fff';
      selectElement.style.fontWeight = 'normal';
      selectElement.style.color = 'var(--text)';
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
      // すべてのinput要素をクリア
      const inputs = row.querySelectorAll('input[type="text"], input[type="time"]');
      inputs.forEach(input => {
        input.value = '';
      });
      
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
        select.selectedIndex = 0;
      });
    });
    
    // 第2ページの用料記録テーブルも同様に処理
    const materialTableRows2 = document.querySelectorAll('#materialTableBody2 tr');
    materialTableRows2.forEach(row => {
      // すべてのinput要素をクリア
      const inputs = row.querySelectorAll('input[type="text"], input[type="time"]');
      inputs.forEach(input => {
        input.value = '';
      });
      
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
    return el.closest('#autoPage') || document.querySelector('.container:first-of-type');
  }

  function getSelectedType(pageRoot) {
    const checked = pageRoot?.querySelector('input[name="type"]:checked');
    return checked ? checked.value : '';
  }

  const NOT_FOUND_CODE = '暫時未搵到產品編碼';
  const NOT_FOUND_NAME = '暫時未搵到產品名稱';

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

  function clearProductNotFoundFields(row) {
    const codeInput = row.querySelector('td:nth-child(3) input');
    const nameInput = row.querySelector('td:nth-child(7) input');
    if (codeInput?.value === NOT_FOUND_CODE) codeInput.value = '';
    if (nameInput?.value === NOT_FOUND_NAME) nameInput.value = '';
    codeInput?.classList.remove('product-not-found');
    nameInput?.classList.remove('product-not-found');
  }

  function showProductNotFound(codeInput, nameInput) {
    codeInput.value = NOT_FOUND_CODE;
    nameInput.value = NOT_FOUND_NAME;
    codeInput.classList.add('product-not-found');
    nameInput.classList.add('product-not-found');
  }

  function showProductMatchPicker(row, matches, onPick) {
    hideProductMatchPicker(row);
    const cell = row.querySelector('td:nth-child(3)');
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
    });
    wrap.appendChild(select);
    cell.appendChild(wrap);
  }

  function isSpecInputCell(row, el) {
    const td = el.closest('td');
    if (!td || !row.contains(td)) return false;
    const idx = [...row.children].indexOf(td) + 1;
    if (idx === 4 && el.matches('select.thickness-select')) return true;
    return [5, 6, 8].includes(idx) && el.matches('input');
  }

  async function tryResolveProductForRow(row) {
    if (row.closest('#materialTableBody') || row.closest('#materialTableBody2')) return;

    const pageRoot = getPageRoot(row);
    const type = getSelectedType(pageRoot);
    const spec = getRowSpecValues(row);
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

    const codeInput = row.querySelector('td:nth-child(3) input');
    const nameInput = row.querySelector('td:nth-child(7) input');
    if (!codeInput || !nameInput) return;

    hideProductResolveHint(row);

    const params = new URLSearchParams({ type, t: thickness, w: width, h: height, l: length });
    if (type === '其他') {
      params.set('other', pageRoot.querySelector('#typeOther')?.value?.trim() || '');
    }

    try {
      const res = await fetch(`${API_BASE}/api/pq_form/plist/search?${params.toString()}`, { cache: 'no-store' });
      const data = await res.json();
      if (!data.success) {
        hideProductMatchPicker(row);
        showProductNotFound(codeInput, nameInput);
        persistLocal();
        adjustNameColumnWidth();
        syncTopRowToMaterial(row);
        return;
      }

      hideProductMatchPicker(row);
      if (data.matches.length === 1) {
        codeInput.value = data.matches[0].code;
        nameInput.value = data.matches[0].name;
        codeInput.classList.remove('product-not-found');
        nameInput.classList.remove('product-not-found');
        hideProductResolveHint(row);
        persistLocal();
        adjustNameColumnWidth();
        syncTopRowToMaterial(row);
      } else if (data.matches.length > 1) {
        const uniqueNames = [...new Set(data.matches.map((m) => m.name))];
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
          nameInput.value = match.name;
          codeInput.classList.remove('product-not-found');
          nameInput.classList.remove('product-not-found');
          hideProductResolveHint(row);
          persistLocal();
          adjustNameColumnWidth();
          syncTopRowToMaterial(row);
        });
      } else {
        showProductNotFound(codeInput, nameInput);
        if (data.hint) showProductResolveHint(row, data.hint);
        persistLocal();
        adjustNameColumnWidth();
        syncTopRowToMaterial(row);
      }
    } catch (error) {
      console.error('plist search failed', error);
      hideProductMatchPicker(row);
      showProductNotFound(codeInput, nameInput);
      persistLocal();
      adjustNameColumnWidth();
      syncTopRowToMaterial(row);
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
        pageRoot.querySelectorAll(`${tableId} tr`).forEach((row) => {
          if (!isMainDataRow(row)) return;
          scheduleResolveProduct(row);
        });
      });
    });

    const typeOther = pageRoot.querySelector('#typeOther');
    if (typeOther) {
      typeOther.addEventListener('input', () => {
        const otherChecked = pageRoot.querySelector('input[name="type"][value="其他"]')?.checked;
        if (!otherChecked) return;
        pageRoot.querySelectorAll('#tableBody tr').forEach((row) => {
          if (!isMainDataRow(row)) return;
          scheduleResolveProduct(row);
        });
      });
    }
  }

  function bindPlistResolveInputs(root) {
    root.addEventListener('input', (e) => {
      const input = e.target;
      if (!input.matches('input')) return;
      const row = input.closest('#tableBody tr, #tableBody2 tr');
      if (!row || !isMainDataRow(row) || !isSpecInputCell(row, input)) return;
      scheduleResolveProduct(row);
      syncTopRowToMaterial(row);
    });
    root.addEventListener('change', (e) => {
      const select = e.target;
      if (!select.matches('select.thickness-select')) return;
      const row = select.closest('#tableBody tr, #tableBody2 tr');
      if (!row) return;
      scheduleResolveProduct(row);
      syncTopRowToMaterial(row);
    });
  }

  function getTopProductSyncFields(row) {
    return {
      thickness: getThicknessValue(row, 4),
      width: row.querySelector('td:nth-child(5) input')?.value ?? '',
      height: row.querySelector('td:nth-child(6) input')?.value ?? '',
      code: row.querySelector('td:nth-child(3) input')?.value ?? '',
      name: row.querySelector('td:nth-child(7) input')?.value ?? '',
    };
  }

  function ensureMaterialRows(materialBody, minCount) {
    if (!materialBody) return;
    while (materialBody.querySelectorAll('tr').length < minCount) {
      materialBody.appendChild(createMaterialRow());
    }
  }

  function syncTopRowToMaterial(topRow) {
    const pageRoot = getPageRoot(topRow);
    const isAuto = pageRoot?.id === 'autoPage';
    const materialBody = isAuto ? materialTableBody2 : materialTableBody;
    if (!materialBody || !topRow.parentElement) return;

    const rowIndex = [...topRow.parentElement.children].indexOf(topRow);
    if (rowIndex < 0) return;

    ensureMaterialRows(materialBody, rowIndex + 1);
    const materialRow = materialBody.querySelectorAll('tr')[rowIndex];
    if (!materialRow) return;

    const { thickness, width, height, code, name } = getTopProductSyncFields(topRow);
    const thicknessSelect1 = getThicknessSelect(materialRow, 2);
    const thicknessSelect2 = getThicknessSelect(materialRow, 5);
    const widthInput2 = materialRow.querySelector('td:nth-child(6) input');
    const heightInput = materialRow.querySelector('td:nth-child(7) input');
    const codeInput = materialRow.querySelector('td:nth-child(8) input');
    const nameInput = materialRow.querySelector('td:nth-child(9) input');

    setThicknessSelectValue(thicknessSelect1, thickness);
    setThicknessSelectValue(thicknessSelect2, thickness);
    if (widthInput2) widthInput2.value = width;
    if (heightInput) heightInput.value = height;
    if (codeInput) codeInput.value = code;
    if (nameInput) nameInput.value = name;
  }

  function syncAllTopToMaterial(pageRoot) {
    const isAuto = pageRoot?.id === 'autoPage';
    const topBody = isAuto ? tableBody2 : tableBody;
    const materialBody = isAuto ? materialTableBody2 : materialTableBody;
    if (!topBody || !materialBody) return;

    const topRows = [...topBody.querySelectorAll('tr')].filter(isMainDataRow);
    ensureMaterialRows(materialBody, topRows.length);
    topRows.forEach((row) => syncTopRowToMaterial(row));
  }

  function isTopProductSyncInput(row, el) {
    const td = el.closest('td');
    if (!td || !row.contains(td)) return false;
    const idx = [...row.children].indexOf(td) + 1;
    if (idx === 4 && el.matches('select.thickness-select')) return true;
    return [3, 5, 6, 7].includes(idx) && el.matches('input');
  }

  function bindTopToMaterialSync(root) {
    const handler = (e) => {
      const el = e.target;
      const row = el.closest('#tableBody tr, #tableBody2 tr');
      if (!row) return;
      if (el.matches('select.thickness-select') && el.closest('td:nth-child(4)')) {
        syncTopRowToMaterial(row);
        return;
      }
      if (el.matches('input') && isTopProductSyncInput(row, el)) {
        syncTopRowToMaterial(row);
      }
    };
    root.addEventListener('input', handler);
    root.addEventListener('change', handler);
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

  function removeRow(){
    let last = tableBody.lastElementChild;
    if (last?.classList.contains('product-hint-row')) {
      tableBody.removeChild(last);
      last = tableBody.lastElementChild;
    }
    if (last && isMainDataRow(last)) tableBody.removeChild(last);
    const trailingHint = tableBody.lastElementChild;
    if (trailingHint?.classList.contains('product-hint-row')) {
      tableBody.removeChild(trailingHint);
    }
    persistLocal();
  }

  function removeMaterialRow(){
    const last = materialTableBody.lastElementChild;
    if(last) materialTableBody.removeChild(last);
    persistLocal();
  }

  function clearAll(){
    tableBody.innerHTML='';
    persistLocal();
  }

  function clearMaterialAll(){
    materialTableBody.innerHTML='';
    persistLocal();
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
      productNo: tr.querySelector('td:nth-child(3) input')?.value || '',
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

    const lengthInput = tr.querySelector('td:nth-child(8) input');
    if (lengthInput) lengthInput.value = data.length || '';

    const operatorSelect = tr.querySelector('td:nth-child(14) select');
    if (operatorSelect) operatorSelect.value = data.operator || '';

    applySplitTimeToRow(tr, 'time-finish', data.finish || '');

    const speedSelect = tr.querySelector('td:nth-child(16) select');
    if (speedSelect) {
      speedSelect.value = data.speed || '';
      if (data.speed) window.formatSpeedDisplay(speedSelect);
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
      orderNo: tr.querySelector('td:nth-child(1) input')?.value || '',
      thickness1: getThicknessValue(tr, 2),
      width1: tr.querySelector('td:nth-child(3) input')?.value || '',
      weight: tr.querySelector('td:nth-child(4) input')?.value || '',
      thickness2: getThicknessValue(tr, 5),
      width2: tr.querySelector('td:nth-child(6) input')?.value || '',
      height: tr.querySelector('td:nth-child(7) input')?.value || '',
      productNo: tr.querySelector('td:nth-child(8) input')?.value || '',
      name: tr.querySelector('td:nth-child(9) input')?.value || '',
      length: tr.querySelector('td:nth-child(10) input')?.value || '',
      qty: tr.querySelector('td:nth-child(11) input')?.value || '',
      complete: !!checkboxes[0]?.checked,
      oldCoil: !!checkboxes[1]?.checked,
      incomplete: checkboxes[2]?.checked ?? true,
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

    tr.querySelector('td:nth-child(1) input').value = data.orderNo || '';
    setThicknessSelectValue(getThicknessSelect(tr, 2), data.thickness1 || '');
    tr.querySelector('td:nth-child(3) input').value = data.width1 || '';
    tr.querySelector('td:nth-child(4) input').value = data.weight || '';
    setThicknessSelectValue(getThicknessSelect(tr, 5), data.thickness2 || '');
    tr.querySelector('td:nth-child(6) input').value = data.width2 || '';
    tr.querySelector('td:nth-child(7) input').value = data.height || '';
    tr.querySelector('td:nth-child(8) input').value = data.productNo || '';
    tr.querySelector('td:nth-child(9) input').value = data.name || '';
    tr.querySelector('td:nth-child(10) input').value = data.length || '';
    tr.querySelector('td:nth-child(11) input').value = data.qty || '';

    const checkboxes = tr.querySelectorAll('td.chk input[type="checkbox"]');
    if (checkboxes[0]) checkboxes[0].checked = !!data.complete;
    if (checkboxes[1]) checkboxes[1].checked = !!data.oldCoil;
    if (checkboxes[2]) {
      checkboxes[2].checked = data.incomplete !== false;
      checkboxes[2].classList.toggle('checkbox-red', checkboxes[2].checked);
    }
  }

  function persistLocal(){
    try{
      const data = {
        y: document.getElementById('year').value,
        m: document.getElementById('month').value,
        d: document.getElementById('day').value,
        rows: [...tableBody.querySelectorAll('tr')].filter(isMainDataRow).map((tr) => serializeMainRow(tr)),
        materialRows: [...materialTableBody.querySelectorAll('tr')].map((tr) => serializeMaterialRow(tr)),
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

  function restoreLocal(){
    try{
      const raw = localStorage.getItem(STORAGE_KEY) || localStorage.getItem('pq-form-ui-v2');
      if(!raw){ 
        addRow(4); 
        addMaterialRow(4); // 用料記録の4行も追加
        return; 
      }
      const data = JSON.parse(raw);
      document.getElementById('year').value = data.y || '';
      document.getElementById('month').value = data.m || '';
      document.getElementById('day').value = data.d || '';
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
        materialTableBody.appendChild(tr);
      });
      
      if(!data.rows || data.rows.length === 0) {
        addRow(4);
      }
      if(!data.materialRows || data.materialRows.length === 0) {
        addMaterialRow(4);
      }
      
    }catch(e){ 
      addRow(4); 
      addMaterialRow(4); // エラー時も用料記録の4行を追加
    }
    // 初期ロード後に列幅調整
    adjustNameColumnWidth();
  }

  // Events
  addRowBtn.addEventListener('click', ()=> addRow(1));
  removeRowBtn.addEventListener('click', removeRow);
  clearBtn.addEventListener('click', clearAll);
  saveLocalBtn.addEventListener('click', ()=>{
    persistLocal();
    try{ alert('已暫存於此裝置（瀏覽器 LocalStorage）'); }catch(e){}
  });
  
  // 用料記録ボタンのイベントリスナー
  addMaterialRowBtn.addEventListener('click', ()=>addMaterialRow(1));
  removeMaterialRowBtn.addEventListener('click', removeMaterialRow);
  clearMaterialBtn.addEventListener('click', clearMaterialAll);

  // 行内ボタン: 暫存/送出（ひとまずローカル保存とコンソール出力の雛形）
  document.addEventListener('click', (e)=>{
    const target = e.target;
    if(target.classList.contains('btn-row-save')){
      persistLocal();
      try{ alert('此行已暫存（目前僅保存在本機）'); }catch(_){}
    }
    if(target.classList.contains('btn-row-send')){
      const tr = target.closest('tr');
      const inAutoPage = tr.closest('#autoPage') !== null;
      const headerPayload = {
        date: {
          y: document.getElementById(inAutoPage ? 'year2' : 'year').value,
          m: document.getElementById(inAutoPage ? 'month2' : 'month').value,
          d: document.getElementById(inAutoPage ? 'day2' : 'day').value
        },
        types: collectChecks(inAutoPage ? '#autoPage input[name="type"]' : '.container:first-of-type input[name="type"]'),
        machines: collectChecks(inAutoPage ? '#autoPage input[name="machine"]' : '.container:first-of-type input[name="machine"]')
      };
      fetch(`${API_BASE}/api/pq_form/update_header`,{
        method:'POST', headers:{'Content-Type':'application/json'}, cache:'no-store', body: JSON.stringify(headerPayload)
      }).catch(()=>{});

      const cells = [...tr.querySelectorAll('input,select')];
      
      // テーブル判定（用料記録テーブルかどうか）
      const isMaterialTable = tr.closest('#materialTableBody') !== null || tr.closest('#materialTableBody2') !== null;
      
      // 取得（UIの並び順）
      let tLoad, tStart, productNo, thickness, roundness, height, name, lengthVal;
      
      if (isMaterialTable) {
        // 用料記録テーブルの場合
        tLoad = cells[0]?.value || '';  // 單號
        tStart = '';  // 開始時間はない
        productNo = (cells[7]?.value || '').toUpperCase(); // 產品編號（8番目の列）を大文字に統一
        thickness = cells[1]?.value || '';  // 材料厚度
        roundness = cells[2]?.value || '';  // 闊度
        height = cells[6]?.value || '';     // 高度
        name = cells[8]?.value || '';       // 產品名稱
        lengthVal = cells[9]?.value || '';  // 長度
      } else {
        // 通常のテーブルの場合
        tLoad = getSplitTimeValue(tr, 'time-load');
        tStart = getSplitTimeValue(tr, 'time-start');
        productNo = (tr.querySelector('td:nth-child(3) input')?.value || '').toUpperCase();
        thickness = getThicknessValue(tr, 4);
        roundness = tr.querySelector('td:nth-child(5) input')?.value || '';
        height = tr.querySelector('td:nth-child(6) input')?.value || '';
        name = tr.querySelector('td:nth-child(7) input')?.value || '';
        lengthVal = tr.querySelector('td:nth-child(8) input')?.value || '';
      }
      // 3段階チェックボックスの値を取得する関数
      const getThreeStateValue = (fieldName) => {
        const checkbox = tr.querySelector(`.three-state-checkbox[data-field="${fieldName}"]`);
        if (checkbox && checkbox.threeStateInstance) {
          return checkbox.threeStateInstance.getValue();
        }
        return '';
      };

      let chkLenTol, chkCutDim, chkLeftRight, chkUpDown, chkTwist, operator, tFinish, note, other;
      
      if (isMaterialTable) {
        // 用料記録テーブルの場合
        chkLenTol    = toTF((cells[11]?.type === 'checkbox') ? !!cells[11].checked : false);  // 完成
        chkCutDim    = toTF((cells[12]?.type === 'checkbox') ? !!cells[12].checked : false);  // 舊卷材
        chkLeftRight = toTF((cells[13]?.type === 'checkbox') ? !!cells[13].checked : false);  // 未完成
        chkUpDown    = 'FALSE';  // 用料記録にはない
        chkTwist     = 'FALSE';  // 用料記録にはない
        operator = '';  // 用料記録にはない
        tFinish  = '';  // 用料記録にはない
        note     = '';  // 用料記録にはない
        other    = '';  // 用料記録にはない
      } else {
        // 通常のテーブルの場合 - 3段階チェックボックスから値を取得
        chkLenTol    = getThreeStateValue('length_tolerance');
        chkCutDim    = getThreeStateValue('section_size');
        chkLeftRight = getThreeStateValue('left_right_bend');
        chkUpDown    = getThreeStateValue('up_down_bend');
        chkTwist     = getThreeStateValue('twist');
        operator = tr.querySelector('td:nth-child(14) select')?.value || '';
        tFinish = getSplitTimeValue(tr, 'time-finish');
        note = tr.querySelector('td:nth-child(16) select')?.value || '';
        other = tr.querySelector('td:nth-child(17) input')?.value || '';
      }

      // シート列（A〜）への固定マッピング
      // A 上料時間, B 開始時間, C 產品編號, D 材料厚度, E 闊度, F 高度,
      // G 產品名稱, H 空, I 空, J 空, K 長度,
      // L 長度公差, M 切面尺寸, N 左右弯曲, O 上下弯曲, P 扭曲,
      // Q 轉機員/檢查員, R 完成時間, S 備註, T 其他
      const mapped = [
        tLoad, tStart, productNo, thickness, roundness, height,
        name,          // G: 產品名稱
        '', '', '',    // H,I,J: 空
        lengthVal,     // K: 長度
        chkLenTol,     // L: 長度公差
        chkCutDim,     // M: 切面尺寸
        chkLeftRight,  // N: 左右弯曲
        chkUpDown,     // O: 上下弯曲
        chkTwist,      // P: 扭曲
        operator,      // Q: 轉機員/檢查員
        tFinish,       // R: 完成時間
        note,          // S: 備註
        other          // T: 其他
      ];

      const payload = { rows: [mapped] };
      fetch(`${API_BASE}/api/pq_form/submit`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        cache:'no-store',
        body: JSON.stringify(payload)
      }).then(r=>r.json()).then(res=>{
        console.log('submit result', res);
        if (res.success) {
          showOutsideMessage(tr, '已送出（伺服器寫入）', 'success');
        } else {
          showOutsideMessage(tr, '送出失敗: ' + (res.error || ''), 'error');
        }
      }).catch(err=>{
        console.error(err);
        showOutsideMessage(tr, '送出失敗', 'error');
      });
    }
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

  // 由日期載入（雛形）
  loadByDateBtn.addEventListener('click', ()=>{
    const y = document.getElementById('year').value;
    const m = document.getElementById('month').value;
    const d = document.getElementById('day').value;
    const dateStr = `${y}/${m}/${d}`;
    fetch(`${API_BASE}/api/pq_form/fetch?date=${encodeURIComponent(dateStr)}`)
      .then(r=>r.json())
      .then(res=>{
        console.log('fetch result', res);
        alert(res.success? `讀取 ${res.rows.length} 行` : '讀取失敗: '+(res.error||''));
      }).catch(err=>{
        console.error(err);
        alert('讀取失敗');
      });
  });
  document.addEventListener('input', (e)=>{
    if(e.target.matches('input,select')) persistLocal();
    if(e.target.closest('td.name')) adjustNameColumnWidth();
  });

  bindSingleTypeSelection(document.querySelector('.container:first-of-type'));
  bindSingleTypeSelection(document.getElementById('autoPage'));
  bindPlistResolveInputs(document);
  bindTopToMaterialSync(document);
  bindTimeSplitInputs(document);

  // Init
  async function initApp() {
    await loadThicknessOptions();
    setToday();
    restoreLocal();
    syncAllTopToMaterial(document.querySelector('.container:first-of-type'));
    syncAllTopToMaterial(document.getElementById('autoPage'));
  }
  initApp();

  // 初期計測（フォント読み込み後）
  window.addEventListener('load', ()=>{
    adjustNameColumnWidth();
    // コード番号フィールドをクリア
    clearProductNumberFields();
  });
  window.addEventListener('resize', adjustNameColumnWidth);
  
  // ページリフレッシュ時にコード番号フィールドをクリア
  window.addEventListener('beforeunload', () => {
    clearProductNumberFields();
  });

  // ページ切り替え機能
  const moldingPage = document.querySelector('.container:first-of-type');
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
    if (addRowBtn2) addRowBtn2.addEventListener('click', () => addRow(1));
    if (removeRowBtn2) removeRowBtn2.addEventListener('click', () => {
      let last = tableBody2.lastElementChild;
      if (last?.classList.contains('product-hint-row')) {
        tableBody2.removeChild(last);
        last = tableBody2.lastElementChild;
      }
      if (last && isMainDataRow(last)) tableBody2.removeChild(last);
      const trailingHint = tableBody2.lastElementChild;
      if (trailingHint?.classList.contains('product-hint-row')) {
        tableBody2.removeChild(trailingHint);
      }
    });
    if (clearBtn2) clearBtn2.addEventListener('click', () => {
      tableBody2.innerHTML='';
    });
    if (saveLocalBtn2) saveLocalBtn2.addEventListener('click', () => saveLocal());
    if (loadByDateBtn2) loadByDateBtn2.addEventListener('click', () => {
      const y = document.getElementById('year2').value;
      const m = document.getElementById('month2').value;
      const d = document.getElementById('day2').value;
      const dateStr = `${y}/${m}/${d}`;
      fetch(`${API_BASE}/api/pq_form/fetch?date=${encodeURIComponent(dateStr)}`)
        .then(r=>r.json())
        .then(res=>{
          console.log('fetch result', res);
          alert(res.success? `讀取 ${res.rows.length} 行` : '讀取失敗: '+(res.error||''));
        }).catch(err=>{
          console.error(err);
          alert('讀取失敗');
        });
    });
    if (addMaterialRowBtn2) addMaterialRowBtn2.addEventListener('click', () => addMaterialRow(1, materialTableBody2));
    if (removeMaterialRowBtn2) removeMaterialRowBtn2.addEventListener('click', () => removeMaterialRow(materialTableBody2));
    if (clearMaterialBtn2) clearMaterialBtn2.addEventListener('click', () => {
      materialTableBody2.innerHTML='';
    });
    
    // 第2ページの日付を今日に設定
    const now = new Date();
    const year2 = document.getElementById('year2');
    const month2 = document.getElementById('month2');
    const day2 = document.getElementById('day2');
    if (year2) year2.value = now.getFullYear();
    if (month2) month2.value = pad(now.getMonth()+1);
    if (day2) day2.value = pad(now.getDate());
    
    console.log('Auto page initialized');
  }
})();


