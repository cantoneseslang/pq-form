(() => {
  const isLocal = ['localhost', '127.0.0.1'].includes(location.hostname);
  const isStaticDev = location.port === '5508';
  const API_BASE = (isLocal && isStaticDev) ? 'http://localhost:5013' : '';
  const STORAGE_KEY = 'material-qc-ui-v2';

  const form = document.getElementById('materialQcForm');
  const receiptDate = document.getElementById('receiptDate');
  const lotPrefix = document.getElementById('lotPrefix');
  const lotSuffix = document.getElementById('lotSuffix');
  const materialThickness = document.getElementById('materialThickness');
  const materialWidth = document.getElementById('materialWidth');
  const supplier = document.getElementById('supplier');
  const inspector = document.getElementById('inspector');
  const substrate = document.getElementById('substrate');
  const thicknessUm = document.getElementById('thicknessUm');
  const standardUm = document.getElementById('standardUm');
  const judgmentBox = document.getElementById('judgmentBox');
  const judgmentValue = document.getElementById('judgmentValue');
  const cameraInput = document.getElementById('cameraInput');
  const captureBtn = document.getElementById('captureBtn');
  const retakeBtn = document.getElementById('retakeBtn');
  const previewWrap = document.getElementById('previewWrap');
  const previewImg = document.getElementById('previewImg');
  const scanStatus = document.getElementById('scanStatus');
  const submitBtn = document.getElementById('submitBtn');
  const clearBtn = document.getElementById('clearBtn');
  const formMessage = document.getElementById('formMessage');

  let imageBase64 = '';
  let imageMimeType = 'image/jpeg';
  let persistTimer = null;
  let scanning = false;

  function todayIsoDate() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function lotPrefixFromDate(isoDate) {
    const text = String(isoDate || '').trim();
    const m = text.match(/^(\d{4})-(\d{2})/);
    if (!m) return 'KS#----';
    const yy = m[1].slice(-2);
    const mm = m[2];
    return `KS#${yy}${mm}`;
  }

  function buildLotNo() {
    const suffix = lotSuffix.value.replace(/\D/g, '').slice(0, 3);
    if (suffix.length !== 3) return '';
    return `${lotPrefixFromDate(receiptDate.value)}${suffix}`;
  }

  function formatThicknessOneDecimal(value) {
    const num = parseFloat(value);
    if (!Number.isFinite(num)) return '';
    return num.toFixed(1);
  }

  function buildMaterialSpec() {
    const parts = [];
    if (materialThickness.value) parts.push(`${materialThickness.value}mm`);
    if (materialWidth.value) parts.push(`${materialWidth.value}mm`);
    return parts.join(' × ');
  }

  function populateStandardOptions() {
    for (let i = 10; i <= 120; i += 1) {
      const val = (i / 10).toFixed(1);
      const opt = document.createElement('option');
      opt.value = val;
      opt.textContent = `${val} μm`;
      standardUm.appendChild(opt);
    }
  }

  function updateLotPrefix() {
    lotPrefix.textContent = lotPrefixFromDate(receiptDate.value);
  }

  function updateJudgment() {
    const measured = parseFloat(thicknessUm.value);
    const standard = parseFloat(standardUm.value);
    if (!Number.isFinite(measured) || !Number.isFinite(standard)) {
      judgmentBox.hidden = true;
      return;
    }

    judgmentBox.hidden = false;
    const pass = measured >= standard;
    judgmentValue.textContent = pass ? '合格' : '不合格';
    judgmentValue.className = `mqc-judgment__value mqc-judgment__value--${pass ? 'pass' : 'fail'}`;
  }

  function showMessage(text, type = 'ok') {
    formMessage.hidden = false;
    formMessage.textContent = text;
    formMessage.className = `mqc-message mqc-message--${type}`;
  }

  function hideMessage() {
    formMessage.hidden = true;
    formMessage.textContent = '';
  }

  function setScanStatus(text, type = 'loading') {
    scanStatus.hidden = false;
    scanStatus.textContent = text;
    scanStatus.className = `mqc-scan-status mqc-scan-status--${type}`;
  }

  function hideScanStatus() {
    scanStatus.hidden = true;
    scanStatus.textContent = '';
  }

  function getFormState() {
    return {
      receiptDate: receiptDate.value,
      lotSuffix: lotSuffix.value,
      materialThickness: materialThickness.value,
      materialWidth: materialWidth.value,
      supplier: supplier.value,
      inspector: inspector.value,
      substrate: substrate.value,
      thicknessUm: thicknessUm.value,
      standardUm: standardUm.value,
    };
  }

  function applyFormState(state) {
    if (!state || typeof state !== 'object') return;
    receiptDate.value = state.receiptDate || todayIsoDate();
    lotSuffix.value = state.lotSuffix || '';
    materialThickness.value = state.materialThickness || '';
    materialWidth.value = state.materialWidth || '';
    supplier.value = state.supplier || '';
    inspector.value = state.inspector || '';
    substrate.value = state.substrate || '';
    thicknessUm.value = state.thicknessUm || '';
    standardUm.value = state.standardUm || '';
    updateLotPrefix();
    updateJudgment();
  }

  function schedulePersist() {
    clearTimeout(persistTimer);
    persistTimer = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(getFormState()));
      } catch { /* ignore quota */ }
    }, 300);
  }

  function loadPersistedState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) applyFormState(JSON.parse(raw));
      else receiptDate.value = todayIsoDate();
    } catch {
      receiptDate.value = todayIsoDate();
    }
    updateLotPrefix();
  }

  function compressImage(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const img = new Image();
        img.onload = () => {
          const maxWidth = 1200;
          const scale = img.width > maxWidth ? maxWidth / img.width : 1;
          const width = Math.round(img.width * scale);
          const height = Math.round(img.height * scale);
          const canvas = document.createElement('canvas');
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, width, height);
          const mimeType = 'image/jpeg';
          const dataUrl = canvas.toDataURL(mimeType, 0.75);
          resolve({
            dataUrl,
            base64: dataUrl.split(',')[1] || '',
            mimeType,
          });
        };
        img.onerror = () => reject(new Error('讀取相片失敗'));
        img.src = reader.result;
      };
      reader.onerror = () => reject(new Error('讀取檔案失敗'));
      reader.readAsDataURL(file);
    });
  }

  async function runScan() {
    if (!imageBase64 || scanning) return;
    scanning = true;
    captureBtn.disabled = true;
    submitBtn.disabled = true;
    setScanStatus('讀取中…', 'loading');

    try {
      const res = await fetch(`${API_BASE}/api/pq_form/material_qc/scan`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ imageBase64, mimeType: imageMimeType }),
      });
      const data = await res.json();

      if (!res.ok) {
        throw new Error(data?.error || `Scan failed (${res.status})`);
      }

      if (data.substrate) {
        const normalized = data.substrate === '鉄' ? '鐵' : data.substrate === '非鉄' ? '非鐵' : data.substrate;
        substrate.value = normalized;
        if (![...substrate.options].some((o) => o.value === normalized)) {
          const opt = document.createElement('option');
          opt.value = normalized;
          opt.textContent = normalized;
          substrate.appendChild(opt);
          substrate.value = normalized;
        }
      }

      if (data.thicknessUm != null) {
        thicknessUm.value = formatThicknessOneDecimal(data.thicknessUm);
        updateJudgment();
      }

      schedulePersist();

      if (data.success) {
        setScanStatus(`讀取完成：${substrate.value || data.substrate} / ${thicknessUm.value} μm`, 'ok');
      } else {
        setScanStatus(data.error || '無法自動讀取，請手動輸入。', 'warn');
      }
    } catch (e) {
      setScanStatus(e?.message || '讀取錯誤，請手動輸入。', 'error');
    } finally {
      scanning = false;
      captureBtn.disabled = false;
      submitBtn.disabled = false;
    }
  }

  async function handleFileSelected(file) {
    if (!file) return;
    hideMessage();

    try {
      const compressed = await compressImage(file);
      imageBase64 = compressed.base64;
      imageMimeType = compressed.mimeType;
      previewImg.src = compressed.dataUrl;
      previewWrap.hidden = false;
      retakeBtn.hidden = false;
      await runScan();
    } catch (e) {
      showMessage(e?.message || '相片處理失敗', 'error');
    }
  }

  function resetPhoto() {
    imageBase64 = '';
    imageMimeType = 'image/jpeg';
    previewImg.removeAttribute('src');
    previewWrap.hidden = true;
    retakeBtn.hidden = true;
    hideScanStatus();
    cameraInput.value = '';
  }

  function resetForm(keepInspector = true) {
    const savedInspector = keepInspector ? inspector.value : '';
    form.reset();
    receiptDate.value = todayIsoDate();
    inspector.value = savedInspector;
    updateLotPrefix();
    judgmentBox.hidden = true;
    resetPhoto();
    hideMessage();
    schedulePersist();
  }

  async function handleSubmit(event) {
    event.preventDefault();
    hideMessage();

    const lot = buildLotNo();
    const insp = inspector.value.trim();
    const sub = substrate.value.trim();
    const thick = parseFloat(formatThicknessOneDecimal(thicknessUm.value));
    const standard = parseFloat(standardUm.value);
    const judgment = Number.isFinite(thick) && Number.isFinite(standard)
      ? (thick >= standard ? '合格' : '不合格')
      : '';

    if (!lot) {
      showMessage('請輸入 3 位數字單號', 'error');
      lotSuffix.focus();
      return;
    }
    if (!insp) {
      showMessage('請輸入檢查員', 'error');
      inspector.focus();
      return;
    }
    if (!sub) {
      showMessage('請選擇基材', 'error');
      substrate.focus();
      return;
    }
    if (!Number.isFinite(thick)) {
      showMessage('請輸入塗裝厚度 (μm)', 'error');
      thicknessUm.focus();
      return;
    }
    if (!imageBase64) {
      showMessage('請先拍攝測厚計', 'error');
      captureBtn.focus();
      return;
    }

    submitBtn.disabled = true;
    submitBtn.textContent = '保存中…';

    try {
      const payload = {
        receiptDate: receiptDate.value,
        lotNo: lot,
        materialThickness: materialThickness.value,
        materialWidth: materialWidth.value.replace(/\D/g, '').slice(0, 3),
        materialSpec: buildMaterialSpec(),
        supplier: supplier.value,
        inspector: insp,
        substrate: sub,
        thicknessUm: thick,
        standardUm: Number.isFinite(standard) ? standard : null,
        judgment,
        imageBase64,
        mimeType: imageMimeType,
      };

      const res = await fetch(`${API_BASE}/api/pq_form/material_qc/submit`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const data = await res.json();

      if (!res.ok || !data.success) {
        throw new Error(data?.error || `Submit failed (${res.status})`);
      }

      let msg = `已保存（第 ${data.row || '—'} 行）`;
      if (data.photoUrl) msg += '，相片已上傳。';
      showMessage(msg, 'ok');

      lotSuffix.value = '';
      materialThickness.value = '';
      materialWidth.value = '';
      supplier.value = '';
      substrate.value = '';
      thicknessUm.value = '';
      standardUm.value = '';
      judgmentBox.hidden = true;
      resetPhoto();
      schedulePersist();
    } catch (e) {
      showMessage(e?.message || '保存失敗', 'error');
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = '保存記錄';
    }
  }

  lotSuffix.addEventListener('input', () => {
    lotSuffix.value = lotSuffix.value.replace(/\D/g, '').slice(0, 3);
    schedulePersist();
  });

  materialWidth.addEventListener('input', () => {
    materialWidth.value = materialWidth.value.replace(/\D/g, '').slice(0, 3);
    schedulePersist();
  });

  thicknessUm.addEventListener('input', () => {
    updateJudgment();
    schedulePersist();
  });

  receiptDate.addEventListener('change', () => {
    updateLotPrefix();
    schedulePersist();
  });

  captureBtn.addEventListener('click', () => cameraInput.click());
  retakeBtn.addEventListener('click', () => {
    resetPhoto();
    cameraInput.click();
  });
  cameraInput.addEventListener('change', () => {
    const file = cameraInput.files?.[0];
    if (file) handleFileSelected(file);
  });

  form.addEventListener('submit', handleSubmit);
  clearBtn.addEventListener('click', () => resetForm(false));

  [materialThickness, supplier, inspector, substrate, standardUm].forEach((el) => {
    el.addEventListener('input', schedulePersist);
    el.addEventListener('change', () => {
      if (el === standardUm) updateJudgment();
      schedulePersist();
    });
  });

  populateStandardOptions();
  loadPersistedState();
})();
