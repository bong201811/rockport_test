/*
 * 재고조사 추가 기능 v1.4.1
 * 1) 최종 엑셀에 스캔원본 시트를 추가 (입력순서 및 자동보정 이전 값 유지)
 * 2) TXT/CSV/TSV/XLSX/XLS 파일의 바코드를 기존 실사에 일괄 누적
 *
 * index.html에서 app.js, partial-mode.js, (사용한다면 barcode-fix.js) 다음에 로드할 것.
 * 기존 app.js / partial-mode.js / 중간저장 포맷은 변경하지 않음.
 */
(() => {
  'use strict';
  if (typeof scanHistory === 'undefined' || typeof resolveBarcode !== 'function' ||
      typeof render !== 'function' || typeof exportExcel !== 'function' ||
      typeof masterMap === 'undefined' || typeof XLSX === 'undefined') {
    console.error('scan-tools.js: app.js 및 엑셀 라이브러리 로드 후 실행해야 합니다.');
    return;
  }

  const MAX_IMPORT_ROWS = 200000;
  const BARCODE_HEADERS = new Set([
    '바코드', '바코드번호', '바코드값', '스캔바코드', '스캔원본',
    'barcode', 'barcodenumber', 'barcodevalue', 'ean', 'ean13', 'ean8', 'jan', 'upc'
  ]);
  let importing = false;

  function asBarcode(value) {
    if (value === null || value === undefined) return '';
    let s = String(value).replace(/^\uFEFF/, '').trim();
    // CSV에서 앞자리 0을 유지하기 위해 ="00123"로 저장한 셀 대응.
    if (/^="[^"]*"$/.test(s)) s = s.slice(2, -1);
    return s.trim();
  }

  function isBarcodeHeader(value) {
    return BARCODE_HEADERS.has(String(value ?? '').normalize('NFKC').replace(/[\s_\-]/g, '').toLowerCase());
  }

  function normalizeHistoryRecord(record, index) {
    const old = typeof record === 'string';
    const raw = old ? record : (record?.input ?? record?.barcode ?? '');
    const matchedBarcode = old ? '' : (record?.barcode ?? '');
    const type = old ? 'legacy' : (record?.type ?? '');
    const master = type === 'matched' ? masterMap[matchedBarcode] : null;
    const source = old ? '구형 기록' : String(record?.source ?? '');
    let label = '확인 필요';
    if (type === 'matched') label = '정상';
    else if (type === 'unknown') label = '미등록';
    else if (type === 'legacy') label = '구형 기록';
    return [index + 1, String(raw), String(matchedBarcode), label,
      master?.styleCode ?? '', master?.sizeCode ?? '', source];
  }

  function appendScanHistorySheet(workbook) {
    if (!Array.isArray(scanHistory) || !scanHistory.length) return false;
    if (workbook.SheetNames?.includes('스캔원본')) return true;
    if (scanHistory.length + 1 > 1048576) {
      throw new Error('엑셀 한 시트의 최대 행 수를 초과했습니다. 원본 기록을 나누어 저장해야 합니다.');
    }
    const data = [['순서', '스캔원본', '매칭바코드', '인식결과', '스타일코드', '사이즈코드', '입력경로']];
    for (let i = 0; i < scanHistory.length; i++) data.push(normalizeHistoryRecord(scanHistory[i], i));
    const sheet = XLSX.utils.aoa_to_sheet(data);
    // 긴 숫자와 앞자리 0은 엑셀에서도 반드시 문자로 유지한다.
    for (let row = 2; row <= data.length; row++) {
      for (const col of ['B', 'C']) {
        const addr = `${col}${row}`;
        sheet[addr] = { t: 's', v: String(data[row - 1][col === 'B' ? 1 : 2]), z: '@' };
      }
    }
    sheet['!cols'] = [
      { wch: 8 }, { wch: 23 }, { wch: 23 }, { wch: 12 },
      { wch: 18 }, { wch: 14 }, { wch: 29 }
    ];
    sheet['!autofilter'] = { ref: `A1:G${data.length}` };
    XLSX.utils.book_append_sheet(workbook, sheet, '스캔원본');
    return true;
  }

  // 전체실사와 간이실사의 기존 내보내기를 그대로 사용해 첫 시트가 바뀌지 않게 한다.
  // 최종 저장 버튼 동작 중에만 writeFile을 감싸 원본 시트를 붙인다.
  const previousExportExcel = exportExcel;
  exportExcel = function (...args) {
    const originalWriteFile = XLSX.writeFile;
    XLSX.writeFile = function (workbook, filename, options) {
      appendScanHistorySheet(workbook);
      return originalWriteFile.call(XLSX, workbook, filename, options);
    };
    try {
      return previousExportExcel.apply(this, args);
    } finally {
      XLSX.writeFile = originalWriteFile;
    }
  };

  function saveRawOnly() {
    if (!scanHistory.length) {
      alert('저장할 스캔 원본이 없습니다.');
      return;
    }
    const workbook = XLSX.utils.book_new();
    appendScanHistorySheet(workbook);
    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    XLSX.writeFile(workbook, `바코드_스캔원본_${stamp}.xlsx`);
  }

  // 큰따옴표로 감싼 콤마, 이중따옴표, CRLF를 포함한 일반적인 CSV 처리.
  function parseDelimited(text, delimiter) {
    const rows = [];
    let current = [];
    let cell = '';
    let quoted = false;
    text = String(text).replace(/^\uFEFF/, '');
    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      if (quoted) {
        if (char === '"' && text[i + 1] === '"') { cell += '"'; i++; }
        else if (char === '"') quoted = false;
        else cell += char;
      } else if (char === '"' && cell === '') {
        quoted = true;
      } else if (char === delimiter) {
        current.push(cell); cell = '';
      } else if (char === '\r' || char === '\n') {
        current.push(cell);
        if (current.some(x => String(x).trim())) rows.push(current);
        current = []; cell = '';
        if (char === '\r' && text[i + 1] === '\n') i++;
      } else {
        cell += char;
      }
    }
    if (quoted) throw new Error('CSV 따옴표가 닫히지 않았습니다.');
    current.push(cell);
    if (current.some(x => String(x).trim())) rows.push(current);
    return rows;
  }

  function extractColumn(rows) {
    if (!rows.length) throw new Error('파일에 바코드가 없습니다.');
    const first = rows[0];
    const headerIndex = first.findIndex(isBarcodeHeader);
    const oneColumn = rows.every(row => row.slice(1).every(cell => !String(cell ?? '').trim()));
    if (headerIndex < 0 && !oneColumn) {
      throw new Error('여러 열이 있는 파일입니다. 첫 행에 "바코드" 열 제목을 넣어 주세요.');
    }
    const barcodeColumn = headerIndex >= 0 ? headerIndex : 0;
    const startRow = headerIndex >= 0 ? 1 : 0;
    const data = rows.slice(startRow).map(row => asBarcode(row[barcodeColumn])).filter(Boolean);
    if (!data.length) throw new Error('읽을 수 있는 바코드가 없습니다.');
    if (data.some(value => /^\d+(?:\.\d+)?e[+-]?\d+$/i.test(value))) {
      throw new Error('일부 바코드가 1.98E+12 같은 지수 표기로 저장되어 있습니다. 원본 문자열을 유지한 파일로 다시 저장해 주세요.');
    }
    if (data.length > MAX_IMPORT_ROWS) throw new Error(`파일이 너무 큽니다. 한 번에 최대 ${MAX_IMPORT_ROWS.toLocaleString()}건까지 불러올 수 있습니다.`);
    return data;
  }

  async function readBarcodeFile(file) {
    const ext = String(file.name).split('.').pop().toLowerCase();
    if (ext === 'txt') {
      // TXT: 바코드 1줄 = 실물 1개. 첫 줄의 "바코드" 제목은 자동으로 제외.
      return extractColumn((await file.text()).replace(/^\uFEFF/, '').split(/\r\n|\r|\n/).map(line => [line]));
    }
    if (ext === 'csv' || ext === 'tsv') {
      const text = await file.text();
      return extractColumn(parseDelimited(text, ext === 'tsv' ? '\t' : ','));
    }
    if (ext === 'xlsx' || ext === 'xls') {
      const book = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      if (!book.SheetNames?.length) throw new Error('엑셀에 시트가 없습니다.');
      const sheet = book.Sheets[book.SheetNames[0]];
      // 숫자 셀의 화면 표시값(0으로 채운 셀 서식 포함)을 우선 사용.
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
      return extractColumn(rows);
    }
    throw new Error('TXT, CSV, TSV, XLSX, XLS 형식만 지원합니다.');
  }

  function setStatus(message, isError = false) {
    const el = document.getElementById('rboBulkStatus');
    if (!el) return;
    el.textContent = message;
    el.style.color = isError ? 'var(--danger,#c62828)' : 'var(--muted,#68707a)';
  }

  function nextFrame() {
    return new Promise(resolve => setTimeout(resolve, 0));
  }

  async function importBarcodeFile(file) {
    if (importing || !file) return;
    if (!Object.keys(masterMap).length) {
      setStatus('⚠️ 기준정보 파일을 먼저 업로드해 주세요.', true);
      return;
    }
    importing = true;
    const fileInput = document.getElementById('rboBulkFile');
    const runButton = document.getElementById('rboBulkRun');
    const watchedIds = ['barcodeInput', 'undoBtn', 'resetBtn', 'saveBtn', 'loadBtn',
      'mergeBtn', 'exportFull', 'exportDiff', 'cameraBtn'];
    const watched = watchedIds.map(id => document.getElementById(id)).filter(Boolean);
    const oldDisabled = watched.map(el => el.disabled);
    fileInput.disabled = true;
    runButton.disabled = true;
    try {
      setStatus(`📖 ${file.name} 파일 읽는 중...`);
      const barcodes = await readBarcodeFile(file);
      const okay = confirm(`${file.name}\n총 ${barcodes.length.toLocaleString()}회 스캔 기록을 기존 실사에 추가할까요?\n\n동일한 바코드가 여러 줄이면 그 횟수만큼 누적됩니다.\n같은 파일을 다시 가져오면 중복 누적되니 주의하세요.`);
      if (!okay) {
        setStatus('파일 불러오기를 취소했습니다. 기존 실사 기록은 유지됩니다.');
        return;
      }
      watched.forEach(el => { el.disabled = true; });
      const stagedMatched = Object.create(null);
      const stagedUnknown = Object.create(null);
      const stagedHistory = [];
      const cache = new Map();
      let matchedCount = 0;
      let unknownCount = 0;
      for (let i = 0; i < barcodes.length; i++) {
        const raw = barcodes[i];
        let result = cache.get(raw);
        if (!result) {
          result = resolveBarcode(raw);
          cache.set(raw, result || { matched: false, usedBarcode: raw });
        }
        if (result?.matched && masterMap[result.usedBarcode]) {
          const key = result.usedBarcode;
          stagedMatched[key] = (stagedMatched[key] || 0) + 1;
          stagedHistory.push({ type: 'matched', input: raw, barcode: key, source: `file:${file.name}` });
          matchedCount++;
        } else {
          stagedUnknown[raw] = (stagedUnknown[raw] || 0) + 1;
          stagedHistory.push({ type: 'unknown', input: raw, barcode: raw, source: `file:${file.name}` });
          unknownCount++;
        }
        if ((i + 1) % 150 === 0) {
          setStatus(`⏳ 분석 중 ${((i + 1) / barcodes.length * 100).toFixed(0)}% · ${i + 1}/${barcodes.length}회`);
          await nextFrame();
        }
      }
      // 분석을 끝낸 후 한 번에 반영하여 실패나 취소 시 기존 실사 보존.
      Object.entries(stagedMatched).forEach(([code, amount]) => {
        scanCount[code] = (scanCount[code] || 0) + amount;
      });
      Object.entries(stagedUnknown).forEach(([code, amount]) => {
        unknownBarcodes[code] = (unknownBarcodes[code] || 0) + amount;
      });
      for (const item of stagedHistory) scanHistory.push(item);
      render();
      const last = document.getElementById('lastScan');
      if (last) last.textContent = `📁 ${file.name} · ${barcodes.length.toLocaleString()}건 추가 (정상 ${matchedCount.toLocaleString()} / 미등록 ${unknownCount.toLocaleString()})`;
      setStatus(`✅ 추가 완료! ${barcodes.length.toLocaleString()}회 · 정상 ${matchedCount.toLocaleString()} / 미등록 ${unknownCount.toLocaleString()} · 기존 실사에 누적됨. 중간 저장을 눌러 기록을 보관하세요.`);
    } catch (error) {
      console.error('바코드 파일 일괄 불러오기 오류', error);
      setStatus(`⚠️ ${error.message || '파일을 처리하지 못했습니다.'}`, true);
    } finally {
      watched.forEach((el, i) => { el.disabled = oldDisabled[i]; });
      fileInput.disabled = false;
      runButton.disabled = false;
      importing = false;
    }
  }

  function installUI() {
    const scanRow = document.querySelector('.scan-row');
    if (!scanRow) {
      console.error('scan-tools.js: 기존 바코드 스캔 화면을 찾지 못했습니다.');
      return;
    }
    const panel = document.createElement('div');
    panel.id = 'rboBulkPanel';
    panel.style.cssText = 'margin-top:12px;padding:11px;border:1px solid var(--line,#dce2ea);border-radius:10px;background:#f8fbff;';
    panel.innerHTML = `
      <div style="font-weight:700;margin-bottom:6px;">📂 바코드 파일 일괄 실사 <small style="color:var(--muted);font-weight:400;">v1.7</small></div>
      <div class="help" style="margin-bottom:8px;">TXT는 바코드 1줄당 1개, CSV·엑셀은 바코드 1행당 1개로 집계합니다. 기존 실사 기록에 누적됩니다.</div>
      <input type="file" id="rboBulkFile" accept=".txt,.csv,.tsv,.xlsx,.xls" aria-label="바코드 기록 파일 선택" style="display:block;width:100%;margin-bottom:7px;" />
      <button type="button" id="rboBulkRun" class="primary" style="width:100%;">📥 바코드 파일 불러와 실사에 추가</button>
      <div id="rboBulkStatus" role="status" class="msg" aria-live="polite">기준정보·현재고 파일을 먼저 올린 뒤 사용하세요.</div>
    `;
    scanRow.insertAdjacentElement('afterend', panel);
    document.getElementById('rboBulkRun').addEventListener('click', () => {
      const file = document.getElementById('rboBulkFile').files?.[0];
      if (!file) { setStatus('⚠️ 바코드 파일을 선택해 주세요.', true); return; }
      importBarcodeFile(file);
    });

    const exportButtons = document.querySelector('.export-actions');
    if (exportButtons) {
      const rawOnly = document.createElement('button');
      rawOnly.type = 'button';
      rawOnly.id = 'rboExportRawOnly';
      rawOnly.textContent = '🧾 스캔 원본만 엑셀 저장';
      rawOnly.title = '입력순서와 자동보정 이전 원본 바코드를 따로 저장합니다.';
      rawOnly.addEventListener('click', saveRawOnly);
      exportButtons.appendChild(rawOnly);
    }
  }

  installUI();
  console.info('✅ 재고조사 v1.7: 원본 스캔 순서 엑셀 + 바코드 파일 일괄 실사');
})();
