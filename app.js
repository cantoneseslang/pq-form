(() => {
  // 同一オリジン運用に対応: ローカル(5508)のみ 5013 を指し、それ以外は空(=同一オリジン)
  const isLocal = ['localhost','127.0.0.1'].includes(location.hostname);
  const isStaticDev = location.port === '5508';
  const API_BASE = (isLocal && isStaticDev) ? 'http://localhost:5013' : '';
  const tableBody = document.getElementById('tableBody');
  const addRowBtn = document.getElementById('addRowBtn');
  const removeRowBtn = document.getElementById('removeRowBtn');
  const clearBtn = document.getElementById('clearBtn');
  const saveLocalBtn = document.getElementById('saveLocalBtn');
  const loadByDateBtn = document.getElementById('loadByDateBtn');

  function pad(n){return n.toString().padStart(2,'0');}

  function setToday(){
    const now = new Date();
    document.getElementById('year').value = now.getFullYear();
    document.getElementById('month').value = pad(now.getMonth()+1);
    document.getElementById('day').value = pad(now.getDate());
  }

  function createRow(){
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="time" /></td>
      <td><input type="time" /></td>
      <td><input type="text" value="" placeholder="" /></td>
      <td><input type="text" placeholder="厚度" /></td>
      <td><input type="text" placeholder="圓度" /></td>
      <td><input type="text" placeholder="高度" /></td>
      <td class="name"><input type="text" value="" placeholder="" /></td>
      <td><input type="text" placeholder="長度" /></td>
      <td class="chk"><input type="checkbox" /></td>
      <td class="chk"><input type="checkbox" /></td>
      <td class="chk"><input type="checkbox" /></td>
      <td class="chk"><input type="checkbox" /></td>
      <td class="chk"><input type="checkbox" /></td>
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
      <td><input type="time" /></td>
      <td><input type="text" placeholder="備註" /></td>
      <td><input type="text" placeholder="其他" /></td>
      <td class="w-160 btns">
        <button class="btn btn-secondary btn-row-save" type="button">暫存</button>
        <button class="btn btn-primary btn-row-send" type="button">送出</button>
      </td>
    `;
    return tr;
  }

  function addRow(n=1){
    for(let i=0;i<n;i++) tableBody.appendChild(createRow());
    persistLocal();
  }

  function removeRow(){
    const last = tableBody.lastElementChild;
    if(last) tableBody.removeChild(last);
    persistLocal();
  }

  function clearAll(){
    tableBody.innerHTML='';
    persistLocal();
  }

  function persistLocal(){
    try{
      const data = {
        y: document.getElementById('year').value,
        m: document.getElementById('month').value,
        d: document.getElementById('day').value,
        rows: [...tableBody.querySelectorAll('tr')].map(tr=>[...tr.querySelectorAll('input,select')].map(el=>el.value))
      };
      localStorage.setItem('pq-form-ui', JSON.stringify(data));
    }catch(e){/* noop */}
  }

  // 產品名稱の列幅を内容に合わせて自動調整
  function adjustNameColumnWidth(){
    const nameHeader = document.querySelector('thead th:nth-child(7)'); // 7列目: 產品名稱
    const nameCells = document.querySelectorAll('#tableBody td.name input');
    if(!nameHeader || nameCells.length === 0) return;

    // 最長テキストの幅を計測
    let maxPx = 0;
    const measurer = document.createElement('span');
    measurer.style.visibility = 'hidden';
    measurer.style.whiteSpace = 'pre';
    measurer.style.position = 'absolute';
    measurer.style.left = '-9999px';
    measurer.style.font = getComputedStyle(nameCells[0]).font;
    document.body.appendChild(measurer);
    nameCells.forEach(inp => {
      measurer.textContent = inp.value || inp.placeholder || '';
      maxPx = Math.max(maxPx, measurer.getBoundingClientRect().width);
    });
    measurer.remove();

    // 余白（入力の左右パディング＋枠＋ゆとり）
    const padding = 48; // px
    let width = Math.ceil(maxPx + padding);
    // 指定に合わせて 1/4 縮小（75%）
    width = Math.ceil(width * 0.75);
    // 下限/上限（画面幅に応じてクランプ）
    const minW = 420;
    const maxW = Math.min(1000, window.innerWidth - 600);
    width = Math.min(Math.max(width, minW), maxW);

    nameHeader.style.width = width + 'px';
    // 各セル自体は100%幅入力なので、セル幅も合わせる
    document.querySelectorAll('#tableBody td.name').forEach(td => {
      td.style.width = width + 'px';
    });
  }

  function restoreLocal(){
    try{
      const raw = localStorage.getItem('pq-form-ui');
      if(!raw){ addRow(3); return; }
      const data = JSON.parse(raw);
      document.getElementById('year').value = data.y || '';
      document.getElementById('month').value = data.m || '';
      document.getElementById('day').value = data.d || '';
      tableBody.innerHTML='';
      (data.rows||[]).forEach(r=>{
        const tr = createRow();
        [...tr.querySelectorAll('input,select')].forEach((el,i)=>{ el.value = r[i]||''; });
        tableBody.appendChild(tr);
      });
      if(tableBody.children.length===0) addRow(3);
    }catch(e){ addRow(3); }
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

  // 行内ボタン: 暫存/送出（ひとまずローカル保存とコンソール出力の雛形）
  document.addEventListener('click', (e)=>{
    const target = e.target;
    if(target.classList.contains('btn-row-save')){
      persistLocal();
      try{ alert('此行已暫存（目前僅保存在本機）'); }catch(_){}
    }
    if(target.classList.contains('btn-row-send')){
      // まずツールバーのヘッダー値をシートに反映
      const headerPayload = {
        date: {
          y: document.getElementById('year').value,
          m: document.getElementById('month').value,
          d: document.getElementById('day').value
        },
        types: collectChecks('input[name="type"]'),
        machines: collectChecks('input[name="machine"]')
      };
      fetch(`${API_BASE}/api/pq_form/update_header`,{
        method:'POST', headers:{'Content-Type':'application/json'}, cache:'no-store', body: JSON.stringify(headerPayload)
      }).catch(()=>{});

      // 次に行データを送信（シート列マッピング固定：写真の列順に合わせる）
      const tr = target.closest('tr');
      const cells = [...tr.querySelectorAll('input,select')];
      // 取得（UIの並び順）
      const tLoad  = cells[0]?.value || '';
      const tStart = cells[1]?.value || '';
      const productNo = cells[2]?.value || '';
      const thickness = cells[3]?.value || '';
      const roundness = cells[4]?.value || '';
      const height = cells[5]?.value || '';
      const name = cells[6]?.value || '';
      const lengthVal = cells[7]?.value || '';
      const toTF = (v)=> v ? 'TRUE' : 'FALSE';
      const chkLenTol    = toTF((cells[8]?.type === 'checkbox') ? !!cells[8].checked : false);
      const chkCutDim    = toTF((cells[9]?.type === 'checkbox') ? !!cells[9].checked : false);
      const chkLeftRight = toTF((cells[10]?.type === 'checkbox') ? !!cells[10].checked : false);
      const chkUpDown    = toTF((cells[11]?.type === 'checkbox') ? !!cells[11].checked : false);
      const chkTwist     = toTF((cells[12]?.type === 'checkbox') ? !!cells[12].checked : false);
      const operator = cells[13]?.value || '';
      const tFinish  = cells[14]?.value || '';
      const note     = cells[15]?.value || '';
      const other    = cells[16]?.value || '';

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
        alert(res.success? '已送出（伺服器寫入）' : '送出失敗: '+ (res.error||''));
      }).catch(err=>{
        console.error(err);
        alert('送出失敗');
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

  // Init
  setToday();
  restoreLocal();
  // 初期計測（フォント読み込み後）
  window.addEventListener('load', adjustNameColumnWidth);
  window.addEventListener('resize', adjustNameColumnWidth);
})();


