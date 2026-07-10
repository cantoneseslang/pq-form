(() => {
  const isLocal = ['localhost', '127.0.0.1'].includes(location.hostname);
  const isStaticDev = location.port === '5508';
  const API_BASE = (isLocal && isStaticDev) ? 'http://localhost:5013' : '';
  const STORAGE_KEY = 'material-qc-ui-v1';

  const form = document.getElementById('materialQcForm');
  const receiptDate = document.getElementById('receiptDate');
  const lotNo = document.getElementById('lotNo');
  const materialSpec = document.getElementById('materialSpec');
  const supplier = document.getElementById('supplier');
  const inspector = document.getElementById('inspector');
  const substrate = document.getElementById('substrate');
  const thicknessUm = document.getElementById('thicknessUm');
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
      lotNo: lotNo.value,
      materialSpec: materialSpec.value,
      supplier: supplier.value,
      inspector: inspector.value,
      substrate: substrate.value,
      thicknessUm: thicknessUm.value,
    };
  }

  function applyFormState(state) {
    if (!state || typeof state !== 'object') return;
    receiptDate.value = state.receiptDate || todayIsoDate();
    lotNo.value = state.lotNo || '';
    materialSpec.value = state.materialSpec || '';
    supplier.value = state.supplier || '';
    inspector.value = state.inspector || '';
    substrate.value = state.substrate || '';
    thicknessUm.value = state.thicknessUm || '';
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
  }

  function compressImage(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const img = new Image();
        img.onload = () => {
          const maxWidth = 1600;
          const scale = img.width > maxWidth ? maxWidth / img.width : 1;
          const width = Math.round(img.width * scale);
          const height = Math.round(img.height * scale);
          const canvas = document.createElement('canvas');
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, width, height);
          const mimeType = 'image/jpeg';
          const dataUrl = canvas.toDataURL(mimeType, 0.85);
          resolve({
            dataUrl,
            base64: dataUrl.split(',')[1] || '',
            mimeType,
          });
        };
        img.onerror = () => reject(new Error('画像の読み込みに失敗しました'));
        img.src = reader.result;
      };
      reader.onerror = () => reject(new Error('ファイルの読み込みに失敗しました'));
      reader.readAsDataURL(file);
    });
  }

  async function runScan() {
    if (!imageBase64 || scanning) return;
    scanning = true;
    captureBtn.disabled = true;
    submitBtn.disabled = true;
    setScanStatus('読取中…', 'loading');

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
        substrate.value = data.substrate;
        if (![...substrate.options].some((o) => o.value === data.substrate)) {
          const opt = document.createElement('option');
          opt.value = data.substrate;
          opt.textContent = data.substrate;
          substrate.appendChild(opt);
          substrate.value = data.substrate;
        }
      }

      if (data.thicknessUm != null) {
        thicknessUm.value = String(data.thicknessUm);
      }

      schedulePersist();

      if (data.success) {
        setScanStatus(`読取完了: ${data.substrate} / ${data.thicknessUm} ${data.unit || 'μm'}`, 'ok');
      } else {
        setScanStatus(data.error || '自動読取できませんでした。手入力してください。', 'warn');
      }
    } catch (e) {
      setScanStatus(e?.message || '読取エラー。手入力してください。', 'error');
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
      showMessage(e?.message || '画像処理に失敗しました', 'error');
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
    resetPhoto();
    hideMessage();
    schedulePersist();
  }

  async function handleSubmit(event) {
    event.preventDefault();
    hideMessage();

    const lot = lotNo.value.trim();
    const insp = inspector.value.trim();
    const sub = substrate.value.trim();
    const thick = parseFloat(thicknessUm.value);

    if (!lot) {
      showMessage('ロット番号を入力してください', 'error');
      lotNo.focus();
      return;
    }
    if (!insp) {
      showMessage('検査員を入力してください', 'error');
      inspector.focus();
      return;
    }
    if (!sub) {
      showMessage('基材を選択してください', 'error');
      substrate.focus();
      return;
    }
    if (!Number.isFinite(thick)) {
      showMessage('塗装厚 (μm) を入力してください', 'error');
      thicknessUm.focus();
      return;
    }

    submitBtn.disabled = true;
    submitBtn.textContent = '送信中…';

    try {
      const payload = {
        receiptDate: receiptDate.value,
        lotNo: lot,
        materialSpec: materialSpec.value.trim(),
        supplier: supplier.value.trim(),
        inspector: insp,
        substrate: sub,
        thicknessUm: thick,
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

      let msg = `記録しました（行 ${data.row || '—'}）`;
      if (data.photoUrl) msg += '。写真を保存しました。';
      else if (data.photoError) msg += `。写真保存失敗: ${data.photoError}`;
      showMessage(msg, data.photoError ? 'error' : 'ok');

      lotNo.value = '';
      materialSpec.value = '';
      supplier.value = '';
      substrate.value = '';
      thicknessUm.value = '';
      resetPhoto();
      schedulePersist();
    } catch (e) {
      showMessage(e?.message || '送信に失敗しました', 'error');
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = '記録する';
    }
  }

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

  [receiptDate, lotNo, materialSpec, supplier, inspector, substrate, thicknessUm].forEach((el) => {
    el.addEventListener('input', schedulePersist);
    el.addEventListener('change', schedulePersist);
  });

  loadPersistedState();
})();
