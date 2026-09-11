// ==================== 데이터 저장소 ====================
let masterMap = {};
let stockMap = {};
let scanCount = {};
let unknownBarcodes = {};
let scanHistory = [];
let currentFilter = 'all';
let deferredInstallPrompt = null;

// ==================== 유틸 ====================
function getStyleSizeKey(styleCode, sizeCode) {
  return `${styleCode}|||${sizeCode}`;
}

function parseStyleSizeKey(key) {
  const [styleCode, sizeCode] = key.split('|||');
  return { styleCode, sizeCode };
}

function cleanText(value) {
  return String(value ?? '').trim();
}

function cleanBarcode(value) {
  // 바코드는 숫자 계산을 하지 않고 문자열 그대로 취급한다.
  // 스캐너가 실수로 넣는 앞뒤 공백/개행만 제거한다.
  return String(value ?? '').trim();
}

function stripLeadingZeros(value) {
  const s = cleanBarcode(value);
  const stripped = s.replace(/^0+/, '');
  return stripped || '0';
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function findHeaderIndex(header, candidates) {
  const normalized = header.map(h => cleanText(h).replace(/\s/g, ''));
  return normalized.findIndex(h => candidates.some(c => h.includes(c)));
}

function setUploadStatus(message, type = '') {
  const el = document.getElementById('uploadStatus');
  el.className = `msg ${type}`.trim();
  el.textContent = message;
}

function updateDataChips() {
  const masterChip = document.getElementById('masterChip');
  const stockChip = document.getElementById('stockChip');
  const masterN = Object.keys(masterMap).length;
  const stockN = Object.keys(stockMap).length;
  masterChip.textContent = `기준정보 ${masterN.toLocaleString()}개`;
  stockChip.textContent = `현재고 ${stockN.toLocaleString()}개`;
  masterChip.className = `chip${masterN ? ' good' : ''}`;
  stockChip.className = `chip${stockN ? ' good' : ''}`;
}

// ==================== 엑셀 파싱 ====================
function parseFile1(data, fileName) {
  const workbook = XLSX.read(data, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });

  if (!rows.length) throw new Error(`${fileName}: 빈 파일입니다.`);

  const header = rows[0];
  const barcodeCol = findHeaderIndex(header, ['바코드']);
  const styleCol = findHeaderIndex(header, ['스타일코드', '스타일']);
  const sizeCol = findHeaderIndex(header, ['사이즈코드', '사이즈']);

  if (barcodeCol < 0 || styleCol < 0) {
    throw new Error(`${fileName}: 바코드 또는 스타일코드 열을 찾지 못했습니다.`);
  }

  let added = 0;
  let duplicated = 0;

  for (let i = 1; i < rows.length; i++) {
    const barcode = cleanBarcode(rows[i][barcodeCol]);
    const styleCode = cleanText(rows[i][styleCol]);
    const sizeCode = sizeCol >= 0 ? cleanText(rows[i][sizeCol]) : '';

    if (!barcode || !styleCode) continue;

    if (masterMap[barcode]) {
      duplicated++;
      continue;
    }

    masterMap[barcode] = {
      styleCode,
      sizeCode,
      styleSizeKey: getStyleSizeKey(styleCode, sizeCode)
    };
    added++;
  }

  console.log(`기준정보 추가: ${fileName} → ${added}개, 중복 ${duplicated}개`);
  return { added, duplicated };
}

function parseFile2(data, fileName = '현재고') {
  const workbook = XLSX.read(data, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });

  if (!rows.length) throw new Error(`${fileName}: 빈 파일입니다.`);

  const header = rows[0];
  const styleCol = findHeaderIndex(header, ['스타일코드', '스타일']);
  const sizeCol = findHeaderIndex(header, ['사이즈코드', '사이즈']);
  const stockCol = findHeaderIndex(header, ['현재고']);

  if (styleCol < 0 || stockCol < 0) {
    throw new Error(`${fileName}: 스타일코드 또는 현재고 열을 찾지 못했습니다.`);
  }

  const nextStockMap = {};

  for (let i = 1; i < rows.length; i++) {
    const styleCode = cleanText(rows[i][styleCol]);
    const sizeCode = sizeCol >= 0 ? cleanText(rows[i][sizeCol]) : '';
    const rawStock = cleanText(rows[i][stockCol]).replaceAll(',', '');
    const stock = Number(rawStock) || 0;

    if (!styleCode) continue;

    const key = getStyleSizeKey(styleCode, sizeCode);
    if (!nextStockMap[key]) nextStockMap[key] = { currentStock: 0 };
    nextStockMap[key].currentStock += stock;
  }

  stockMap = nextStockMap;
  console.log(`현재고 로드: ${Object.keys(stockMap).length}개`);
}

// ==================== 바코드 매칭 ====================
function resolveBarcode(inputBarcode) {
  const barcode = cleanBarcode(inputBarcode);
  if (!barcode) return null;

  // 1) 가장 안전한 정확 일치
  if (masterMap[barcode]) {
    return { matched: true, usedBarcode: barcode, method: 'exact' };
  }

  const keys = Object.keys(masterMap);
  if (!keys.length) return { matched: false, usedBarcode: barcode, method: 'none' };

  // 2) 앞의 0이 빠진 경우 대응
  // 예: 기준 0000070608 / 스캐너 70608
  const sig = stripLeadingZeros(barcode);
  const zeroEquivalent = keys.filter(k => stripLeadingZeros(k) === sig);
  if (zeroEquivalent.length === 1) {
    return { matched: true, usedBarcode: zeroEquivalent[0], method: 'leading-zero' };
  }

  // 3) 예전 호환용: 짧게 들어온 값이 기준 바코드의 유일한 끝자리인 경우
  const suffixCandidates = keys.filter(k => k.length > barcode.length && k.endsWith(barcode));
  if (suffixCandidates.length === 1) {
    return { matched: true, usedBarcode: suffixCandidates[0], method: 'suffix' };
  }

  return { matched: false, usedBarcode: barcode, method: 'none' };
}

function aggregateScannedByStyleSize() {
  const map = {};
  for (const [barcode, count] of Object.entries(scanCount)) {
    const info = masterMap[barcode];
    if (!info) continue;
    map[info.styleSizeKey] = (map[info.styleSizeKey] || 0) + count;
  }
  return map;
}

// ==================== 표시 ====================
function render() {
  const allKeys = new Set(Object.keys(stockMap));
  const scannedByKey = aggregateScannedByStyleSize();
  Object.keys(scannedByKey).forEach(k => allKeys.add(k));
  Object.keys(unknownBarcodes).forEach(barcode => allKeys.add(`__UNKNOWN__${barcode}`));

  let totalScanned = 0;
  let totalStock = 0;
  let totalDiff = 0;
  const rows = [];

  allKeys.forEach(key => {
    if (key.startsWith('__UNKNOWN__')) {
      const barcode = key.slice('__UNKNOWN__'.length);
      const scanned = unknownBarcodes[barcode] || 0;
      totalScanned += scanned;
      totalDiff += scanned;
      rows.push({ type: 'unknown', barcode, scanned, stock: null, diff: scanned });
      return;
    }

    const { styleCode, sizeCode } = parseStyleSizeKey(key);
    const stock = stockMap[key]?.currentStock || 0;
    const scanned = scannedByKey[key] || 0;
    const diff = scanned - stock;

    totalScanned += scanned;
    totalStock += stock;
    totalDiff += diff;
    rows.push({ type: 'normal', styleCode, sizeCode, scanned, stock, diff });
  });

  let filteredRows = rows;
  if (currentFilter === 'scanned') {
    filteredRows = rows.filter(r => r.type === 'unknown' || r.scanned > 0);
  } else if (currentFilter === 'diff') {
    filteredRows = rows.filter(r => r.diff !== 0);
  }

  filteredRows.sort((a, b) => {
    if (a.type === 'unknown' && b.type !== 'unknown') return 1;
    if (a.type !== 'unknown' && b.type === 'unknown') return -1;
    if (a.type === 'normal' && b.type === 'normal') {
      return a.styleCode.localeCompare(b.styleCode, 'ko', { numeric: true }) ||
             a.sizeCode.localeCompare(b.sizeCode, 'ko', { numeric: true });
    }
    return 0;
  });

  document.getElementById('sumScanned').textContent = totalScanned.toLocaleString();
  document.getElementById('sumStock').textContent = totalStock.toLocaleString();
  const diffEl = document.getElementById('sumDiff');
  diffEl.textContent = `${totalDiff > 0 ? '+' : ''}${totalDiff.toLocaleString()}`;
  diffEl.className = `v ${totalDiff > 0 ? 'diff-plus' : totalDiff < 0 ? 'diff-minus' : 'diff-zero'}`;

  const area = document.getElementById('resultArea');
  if (!filteredRows.length) {
    area.innerHTML = '<div style="padding:14px;color:#8a929c;">표시할 데이터가 없습니다.</div>';
  } else {
    area.innerHTML = filteredRows.map(r => {
      if (r.type === 'unknown') {
        return `<div class="result-row unknown">
          <div class="code">미등록 바코드 · ${escapeHtml(r.barcode)}</div>
          <div class="nums">실사 ${r.scanned} · 현재고 - · <strong>차이 +${r.diff}</strong></div>
        </div>`;
      }

      const diffClass = r.diff > 0 ? 'diff-plus' : r.diff < 0 ? 'diff-minus' : 'diff-zero';
      const diffStr = `${r.diff > 0 ? '+' : ''}${r.diff}`;
      return `<div class="result-row">
        <div class="code">${escapeHtml(r.styleCode)} / ${escapeHtml(r.sizeCode || '-')}</div>
        <div class="nums">실사 ${r.scanned} · 현재고 ${r.stock} · <span class="${diffClass}">차이 ${diffStr}</span></div>
      </div>`;
    }).join('');
  }

  updateSaveList();
  updateDataChips();
}

// ==================== 바코드 처리 ====================
function processBarcode(rawBarcode, source = 'scanner') {
  const barcode = cleanBarcode(rawBarcode);
  if (!barcode) return;

  if (!Object.keys(masterMap).length) {
    document.getElementById('lastScan').innerHTML = '<span class="warn-text">⚠️ 기준정보 파일을 먼저 불러오세요.</span>';
    return;
  }

  const result = resolveBarcode(barcode);

  if (result.matched) {
    const usedBarcode = result.usedBarcode;
    scanCount[usedBarcode] = (scanCount[usedBarcode] || 0) + 1;
    scanHistory.push({ type: 'matched', input: barcode, barcode: usedBarcode, source });

    const info = masterMap[usedBarcode];
    const key = info.styleSizeKey;
    const stock = stockMap[key]?.currentStock || 0;
    const totalForKey = Object.entries(scanCount)
      .filter(([bc]) => masterMap[bc]?.styleSizeKey === key)
      .reduce((sum, [, count]) => sum + count, 0);

    let correction = '';
    if (result.method === 'leading-zero') correction = ` <span class="help">(${escapeHtml(barcode)} → ${escapeHtml(usedBarcode)} 자동보정)</span>`;
    if (result.method === 'suffix') correction = ` <span class="help">(${escapeHtml(barcode)} → ${escapeHtml(usedBarcode)} 끝자리 매칭)</span>`;

    document.getElementById('lastScan').innerHTML =
      `<span class="ok">✅ ${escapeHtml(info.styleCode)} / ${escapeHtml(info.sizeCode || '-')} → 실사 ${totalForKey} · 현재고 ${stock} · 차이 ${totalForKey - stock >= 0 ? '+' : ''}${totalForKey - stock}</span>${correction}`;
  } else {
    unknownBarcodes[barcode] = (unknownBarcodes[barcode] || 0) + 1;
    scanHistory.push({ type: 'unknown', input: barcode, barcode, source });
    document.getElementById('lastScan').innerHTML =
      `<span class="warn-text">⚠️ 미등록 바코드: ${escapeHtml(barcode)} (${unknownBarcodes[barcode]}회)</span>`;
  }

  render();
}

// ==================== 되돌리기 ====================
function normalizeHistoryItem(item) {
  // 구버전 세션은 문자열 히스토리였으므로 호환 처리
  if (typeof item === 'string') {
    const resolved = resolveBarcode(item);
    if (resolved?.matched) return { type: 'matched', input: item, barcode: resolved.usedBarcode, source: 'legacy' };
    return { type: 'unknown', input: item, barcode: item, source: 'legacy' };
  }
  return item;
}

function undoLastScan() {
  if (!scanHistory.length) {
    document.getElementById('lastScan').innerHTML = '<span class="warn-text">⚠️ 되돌릴 스캔이 없습니다.</span>';
    return;
  }

  const item = normalizeHistoryItem(scanHistory.pop());
  if (!item) return;

  if (item.type === 'matched' && scanCount[item.barcode]) {
    scanCount[item.barcode]--;
    if (scanCount[item.barcode] <= 0) delete scanCount[item.barcode];
  } else if (item.type === 'unknown' && unknownBarcodes[item.barcode]) {
    unknownBarcodes[item.barcode]--;
    if (unknownBarcodes[item.barcode] <= 0) delete unknownBarcodes[item.barcode];
  }

  document.getElementById('lastScan').innerHTML = `<span class="cancel-msg">⚠️ ${escapeHtml(item.input || item.barcode)} 취소됨!</span>`;
  render();
  document.getElementById('barcodeInput').focus();
}

// ==================== 중간 저장 ====================
function saveSession(name) {
  if (!name) {
    alert('저장 이름을 입력하세요!');
    return;
  }

  const data = {
    version: 2,
    scanCount,
    unknownBarcodes,
    scanHistory,
    timestamp: new Date().toISOString()
  };

  const allSessions = JSON.parse(localStorage.getItem('rbo_sessions') || '{}');
  allSessions[name] = data;
  localStorage.setItem('rbo_sessions', JSON.stringify(allSessions));
  document.getElementById('saveMsg').innerHTML = `<span class="ok">✅ "${escapeHtml(name)}" 저장 완료!</span>`;
  updateSaveList();
}

function loadSession(name) {
  const allSessions = JSON.parse(localStorage.getItem('rbo_sessions') || '{}');
  const session = allSessions[name];
  if (!session) {
    alert('저장된 세션이 없습니다.');
    return;
  }

  scanCount = session.scanCount || {};
  unknownBarcodes = session.unknownBarcodes || {};
  scanHistory = (session.scanHistory || []).map(normalizeHistoryItem).filter(Boolean);
  document.getElementById('saveMsg').innerHTML = `<span class="ok">✅ "${escapeHtml(name)}" 불러옴!</span>`;
  render();
}

function deleteSession(name) {
  if (!confirm(`"${name}" 세션을 정말 삭제할까요?`)) return;
  const allSessions = JSON.parse(localStorage.getItem('rbo_sessions') || '{}');
  delete allSessions[name];
  localStorage.setItem('rbo_sessions', JSON.stringify(allSessions));
  document.getElementById('saveMsg').innerHTML = `<span class="warn-text">🗑️ "${escapeHtml(name)}" 삭제 완료!</span>`;
  updateSaveList();
}

function mergeSessions() {
  const allSessions = JSON.parse(localStorage.getItem('rbo_sessions') || '{}');
  const names = Object.keys(allSessions);
  if (names.length < 2) {
    alert('합칠 세션이 2개 이상 필요합니다.');
    return;
  }

  const mergedScan = {};
  const mergedUnknown = {};
  const mergedHistory = [];

  names.forEach(name => {
    const s = allSessions[name];
    Object.entries(s.scanCount || {}).forEach(([k, v]) => {
      mergedScan[k] = (mergedScan[k] || 0) + v;
    });
    Object.entries(s.unknownBarcodes || {}).forEach(([k, v]) => {
      mergedUnknown[k] = (mergedUnknown[k] || 0) + v;
    });
    (s.scanHistory || []).forEach(item => mergedHistory.push(normalizeHistoryItem(item)));
  });

  const newName = prompt('합친 세션 이름:', names.join('+'));
  if (!newName) return;

  allSessions[newName] = {
    version: 2,
    scanCount: mergedScan,
    unknownBarcodes: mergedUnknown,
    scanHistory: mergedHistory,
    timestamp: new Date().toISOString()
  };

  localStorage.setItem('rbo_sessions', JSON.stringify(allSessions));
  scanCount = mergedScan;
  unknownBarcodes = mergedUnknown;
  scanHistory = mergedHistory;
  document.getElementById('saveMsg').innerHTML = `<span class="ok">✅ "${escapeHtml(newName)}" 합치기 완료!</span>`;
  render();
}

function updateSaveList() {
  const allSessions = JSON.parse(localStorage.getItem('rbo_sessions') || '{}');
  const names = Object.keys(allSessions);
  const listEl = document.getElementById('saveList');

  if (!names.length) {
    listEl.innerHTML = '';
    return;
  }

  listEl.innerHTML = names.map(name => `
    <span class="session-item" data-name="${escapeHtml(name)}">
      <span class="session-load">📦 ${escapeHtml(name)}</span>
      <button class="session-delete" title="삭제" aria-label="${escapeHtml(name)} 삭제">✕</button>
    </span>
  `).join('');

  listEl.querySelectorAll('.session-load').forEach(el => {
    el.addEventListener('click', () => loadSession(el.parentElement.dataset.name));
  });

  listEl.querySelectorAll('.session-delete').forEach(el => {
    el.addEventListener('click', e => {
      e.stopPropagation();
      deleteSession(el.parentElement.dataset.name);
    });
  });
}

// ==================== 엑셀 내보내기 ====================
function buildExportData(filterDiffOnly = false) {
  const rows = [['구분', '스타일코드', '사이즈코드', '현재고', '실사', '차이']];
  const allKeys = new Set(Object.keys(stockMap));
  const scannedByKey = aggregateScannedByStyleSize();
  Object.keys(scannedByKey).forEach(k => allKeys.add(k));
  Object.keys(unknownBarcodes).forEach(barcode => allKeys.add(`__UNKNOWN__${barcode}`));

  const normalRows = [];
  const unknownRows = [];

  allKeys.forEach(key => {
    if (key.startsWith('__UNKNOWN__')) {
      const barcode = key.slice('__UNKNOWN__'.length);
      const scanned = unknownBarcodes[barcode] || 0;
      if (filterDiffOnly && scanned === 0) return;
      unknownRows.push(['미등록바코드', barcode, '-', '-', scanned, scanned]);
      return;
    }

    const { styleCode, sizeCode } = parseStyleSizeKey(key);
    const stock = stockMap[key]?.currentStock || 0;
    const scanned = scannedByKey[key] || 0;
    const diff = scanned - stock;
    if (filterDiffOnly && diff === 0) return;
    normalRows.push(['정상', styleCode, sizeCode, stock, scanned, diff]);
  });

  normalRows.sort((a, b) => String(a[1]).localeCompare(String(b[1]), 'ko', { numeric: true }) || String(a[2]).localeCompare(String(b[2]), 'ko', { numeric: true }));
  unknownRows.sort((a, b) => String(a[1]).localeCompare(String(b[1]), 'ko', { numeric: true }));
  rows.push(...normalRows, ...unknownRows);
  return rows;
}

function exportExcel(filterDiffOnly = false) {
  if (!Object.keys(stockMap).length && !Object.keys(scanCount).length && !Object.keys(unknownBarcodes).length) {
    alert('저장할 데이터가 없습니다.');
    return;
  }

  const data = buildExportData(filterDiffOnly);
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [
    { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 10 }, { wch: 10 }, { wch: 10 }
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, filterDiffOnly ? '차이내역' : '전체');

  const now = new Date();
  const stamp = [now.getFullYear(), String(now.getMonth() + 1).padStart(2, '0'), String(now.getDate()).padStart(2, '0')].join('');
  XLSX.writeFile(wb, filterDiffOnly ? `재고조사_차이내역_${stamp}.xlsx` : `재고조사_전체_${stamp}.xlsx`);
}

// ==================== 카메라 바코드 스캔 ====================
let cameraStream = null;
let cameraDetector = null;
let cameraLoopRunning = false;
let cameraDetectBusy = false;
let lastCameraValue = '';
let lastCameraAt = 0;

async function startCameraScan() {
  if (!('BarcodeDetector' in window)) {
    alert('이 브라우저는 카메라 자동 바코드 인식을 지원하지 않습니다.\n\n블루투스/USB 바코드 스캐너는 그대로 사용할 수 있습니다. 카메라 기능은 Edge/Chrome 계열에서 이용해 보세요.');
    return;
  }

  if (!navigator.mediaDevices?.getUserMedia) {
    alert('카메라를 사용할 수 없는 환경입니다. HTTPS 주소에서 열었는지 확인해주세요.');
    return;
  }

  try {
    const supported = typeof BarcodeDetector.getSupportedFormats === 'function'
      ? await BarcodeDetector.getSupportedFormats()
      : [];
    const wanted = ['ean_13', 'ean_8', 'upc_a', 'upc_e', 'code_128', 'code_39', 'itf', 'codabar'];
    const formats = supported.length ? wanted.filter(f => supported.includes(f)) : undefined;
    cameraDetector = formats?.length ? new BarcodeDetector({ formats }) : new BarcodeDetector();

    cameraStream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode: { ideal: 'environment' },
        width: { ideal: 1920 },
        height: { ideal: 1080 }
      },
      audio: false
    });

    const video = document.getElementById('cameraVideo');
    video.srcObject = cameraStream;
    await video.play();

    const layer = document.getElementById('cameraLayer');
    layer.classList.add('open');
    layer.setAttribute('aria-hidden', 'false');
    document.getElementById('cameraMsg').textContent = '바코드를 사각형 안에 맞춰주세요. 같은 바코드는 약 1.2초 후 다시 셀 수 있습니다.';

    cameraLoopRunning = true;
    cameraDetectLoop();
  } catch (err) {
    console.error(err);
    stopCameraScan();
    alert('카메라를 열지 못했습니다. 브라우저의 카메라 권한을 확인해주세요.');
  }
}

async function cameraDetectLoop() {
  if (!cameraLoopRunning) return;
  const video = document.getElementById('cameraVideo');

  if (!cameraDetectBusy && cameraDetector && video.readyState >= 2) {
    cameraDetectBusy = true;
    try {
      const codes = await cameraDetector.detect(video);
      if (codes.length) {
        const value = cleanBarcode(codes[0].rawValue);
        const now = Date.now();
        const duplicateTooSoon = value === lastCameraValue && now - lastCameraAt < 1200;

        if (value && !duplicateTooSoon) {
          lastCameraValue = value;
          lastCameraAt = now;
          processBarcode(value, 'camera');
          document.getElementById('cameraMsg').textContent = `✅ ${value} 스캔됨`;
          if (navigator.vibrate) navigator.vibrate(50);
        }
      }
    } catch (err) {
      console.warn('Barcode detect error:', err);
    } finally {
      cameraDetectBusy = false;
    }
  }

  setTimeout(cameraDetectLoop, 180);
}

function stopCameraScan() {
  cameraLoopRunning = false;
  cameraDetectBusy = false;
  if (cameraStream) {
    cameraStream.getTracks().forEach(track => track.stop());
    cameraStream = null;
  }
  const video = document.getElementById('cameraVideo');
  if (video) video.srcObject = null;
  const layer = document.getElementById('cameraLayer');
  if (layer) {
    layer.classList.remove('open');
    layer.setAttribute('aria-hidden', 'true');
  }
  document.getElementById('barcodeInput')?.focus();
}

// ==================== 파일 이벤트 ====================
async function loadMasterFiles(files) {
  if (!files?.length) return;
  let addedTotal = 0;
  let duplicateTotal = 0;

  try {
    for (const file of files) {
      const data = await file.arrayBuffer();
      const result = parseFile1(data, file.name);
      addedTotal += result.added;
      duplicateTotal += result.duplicated;
    }
    setUploadStatus(`✅ 기준정보 ${addedTotal.toLocaleString()}개 추가${duplicateTotal ? ` · 중복 ${duplicateTotal.toLocaleString()}개 제외` : ''}`, 'ok');
    render();
  } catch (err) {
    console.error(err);
    setUploadStatus(`⚠️ ${err.message}`, 'warn-text');
  }
}

document.getElementById('file1').addEventListener('change', e => loadMasterFiles(e.target.files));

document.getElementById('file1Extra').addEventListener('change', async e => {
  await loadMasterFiles(e.target.files);
  e.target.value = '';
});

document.getElementById('addFileBtn').addEventListener('click', () => {
  document.getElementById('file1Extra').click();
});

document.getElementById('file2').addEventListener('change', async e => {
  const file = e.target.files?.[0];
  if (!file) return;

  try {
    const data = await file.arrayBuffer();
    parseFile2(data, file.name);
    setUploadStatus(`✅ 현재고 로드 완료 · ${Object.keys(stockMap).length.toLocaleString()}개 스타일/사이즈`, 'ok');
    render();
  } catch (err) {
    console.error(err);
    setUploadStatus(`⚠️ ${err.message}`, 'warn-text');
  }
});

// ==================== UI 이벤트 ====================
document.getElementById('barcodeInput').addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    e.preventDefault();
    processBarcode(e.target.value, 'scanner');
    e.target.value = '';
  }
});

document.getElementById('undoBtn').addEventListener('click', undoLastScan);
document.getElementById('cameraBtn').addEventListener('click', startCameraScan);
document.getElementById('cameraCloseBtn').addEventListener('click', stopCameraScan);

document.querySelectorAll('.filter-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentFilter = btn.dataset.filter;
    render();
  });
});

document.getElementById('saveBtn').addEventListener('click', () => {
  saveSession(document.getElementById('saveName').value.trim());
});

document.getElementById('loadBtn').addEventListener('click', () => {
  const allSessions = JSON.parse(localStorage.getItem('rbo_sessions') || '{}');
  const names = Object.keys(allSessions);
  if (!names.length) {
    alert('저장된 세션이 없습니다.');
    return;
  }
  const name = prompt(`불러올 세션 이름:\n${names.join('\n')}`, names[0]);
  if (name) loadSession(name);
});

document.getElementById('mergeBtn').addEventListener('click', mergeSessions);
document.getElementById('exportFull').addEventListener('click', () => exportExcel(false));
document.getElementById('exportDiff').addEventListener('click', () => exportExcel(true));

document.getElementById('resetBtn').addEventListener('click', () => {
  if (!confirm('실사 수량과 미등록 바코드를 모두 초기화할까요? 저장된 중간 세션은 삭제되지 않습니다.')) return;
  scanCount = {};
  unknownBarcodes = {};
  scanHistory = [];
  document.getElementById('lastScan').textContent = '';
  render();
  document.getElementById('barcodeInput').focus();
});

// ==================== PWA 설치 ====================
window.addEventListener('beforeinstallprompt', event => {
  event.preventDefault();
  deferredInstallPrompt = event;
  document.getElementById('installBtn').style.display = '';
});

document.getElementById('installBtn').addEventListener('click', async () => {
  if (!deferredInstallPrompt) return;
  deferredInstallPrompt.prompt();
  await deferredInstallPrompt.userChoice;
  deferredInstallPrompt = null;
  document.getElementById('installBtn').style.display = 'none';
});

window.addEventListener('appinstalled', () => {
  deferredInstallPrompt = null;
  document.getElementById('installBtn').style.display = 'none';
});

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./service-worker.js').catch(err => console.warn('Service worker 등록 실패:', err));
  });
}

window.addEventListener('pagehide', stopCameraScan);

// ==================== 시작 ====================
updateSaveList();
updateDataChips();
document.getElementById('barcodeInput').focus();
