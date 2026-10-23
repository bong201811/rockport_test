/*
 * RBO 재고조사 - 간이 재고실사 추가 기능 v1.0
 * 원본 app.js 다음에 로드하는 보조 스크립트입니다.
 * 원본 바코드 스캔, 저장/불러오기/합치기, 되돌리기 코드는 변경하지 않습니다.
 */
(() => {
  'use strict';

  // 원본과 서로 다른 이름을 사용해 충돌을 피합니다.
  const MODE_KEY = 'rbo_inventory_compare_mode_v1';
  let compareMode = localStorage.getItem(MODE_KEY) === 'partial' ? 'partial' : 'full';
  const baseRender = render;
  const baseBuildExportData = buildExportData;
  const baseExportExcel = exportExcel;

  function activeStyles() {
    const result = new Set();
    for (const [barcode, qty] of Object.entries(scanCount)) {
      const info = masterMap[barcode];
      if (info && Number(qty) > 0) result.add(info.styleCode);
    }
    return result;
  }

  function partialRows() {
    const styles = activeStyles();
    const scannedByKey = aggregateScannedByStyleSize();
    const allKeys = new Set();

    // 스캔된 스타일에 속한 현재고 전체 사이즈를 가져옵니다.
    for (const key of Object.keys(stockMap)) {
      if (styles.has(parseStyleSizeKey(key).styleCode)) allKeys.add(key);
    }
    // 전산 파일에 사이즈 행이 없어도 실제 스캔한 상품은 표시합니다.
    for (const key of Object.keys(scannedByKey)) {
      if (styles.has(parseStyleSizeKey(key).styleCode)) allKeys.add(key);
    }

    const rows = [...allKeys].map(key => {
      const { styleCode, sizeCode } = parseStyleSizeKey(key);
      const stock = stockMap[key]?.currentStock || 0;
      const scanned = scannedByKey[key] || 0;
      return { type: 'normal', styleCode, sizeCode, scanned, stock, diff: scanned - stock };
    });
    rows.sort((a, b) =>
      a.styleCode.localeCompare(b.styleCode, 'ko', { numeric: true }) ||
      a.sizeCode.localeCompare(b.sizeCode, 'ko', { numeric: true })
    );

    // 매칭하지 못한 바코드는 스타일을 유추하지 않고 별도 유지합니다.
    const unknownRows = Object.entries(unknownBarcodes)
      .filter(([, qty]) => Number(qty) > 0)
      .map(([barcode, qty]) => ({
        type: 'unknown', barcode, scanned: Number(qty), stock: null, diff: Number(qty)
      }))
      .sort((a, b) => a.barcode.localeCompare(b.barcode, 'ko', { numeric: true }));
    rows.push(...unknownRows);
    return { rows, styles };
  }

  function renderPartial() {
    const { rows, styles } = partialRows();
    let scannedSum = 0;
    let stockSum = 0;
    let diffSum = 0;
    for (const row of rows) {
      scannedSum += row.scanned;
      stockSum += row.stock || 0;
      diffSum += row.diff;
    }

    document.getElementById('sumScanned').textContent = scannedSum.toLocaleString();
    document.getElementById('sumStock').textContent = stockSum.toLocaleString();
    const diffEl = document.getElementById('sumDiff');
    diffEl.textContent = `${diffSum > 0 ? '+' : ''}${diffSum.toLocaleString()}`;
    diffEl.className = `v ${diffSum > 0 ? 'diff-plus' : diffSum < 0 ? 'diff-minus' : 'diff-zero'}`;

    let filteredRows = rows;
    if (currentFilter === 'scanned') {
      filteredRows = rows.filter(r => r.type === 'unknown' || r.scanned > 0);
    } else if (currentFilter === 'diff') {
      filteredRows = rows.filter(r => r.diff !== 0);
    }

    const area = document.getElementById('resultArea');
    if (!styles.size && !rows.length) {
      area.innerHTML = '<div style="padding:14px;color:#8a929c;">아직 스캔한 스타일이 없습니다. 첫 바코드를 찍으면 해당 스타일의 모든 사이즈가 자동으로 비교됩니다.</div>';
    } else if (!filteredRows.length) {
      area.innerHTML = '<div style="padding:14px;color:#8a929c;">선택한 필터에 표시할 결과가 없습니다.</div>';
    } else {
      area.innerHTML = filteredRows.map(row => {
        if (row.type === 'unknown') {
          return `<div class="result-row unknown">
            <div class="code">미등록 바코드 · ${escapeHtml(row.barcode)}</div>
            <div class="nums">실사 ${row.scanned} · 현재고 - · <strong>확인 필요</strong></div>
          </div>`;
        }
        const cls = row.diff > 0 ? 'diff-plus' : row.diff < 0 ? 'diff-minus' : 'diff-zero';
        const d = `${row.diff > 0 ? '+' : ''}${row.diff}`;
        return `<div class="result-row">
          <div class="code">${escapeHtml(row.styleCode)} / ${escapeHtml(row.sizeCode || '-')}</div>
          <div class="nums">실사 ${row.scanned} · 현재고 ${row.stock} · <span class="${cls}">차이 ${d}</span></div>
        </div>`;
      }).join('');
    }

    updateSaveList();
    updateDataChips();
    updatePartialModeInfo(styles);
  }

  function updatePartialModeInfo(styles) {
    const summary = document.getElementById('partialStyleSummary');
    if (!summary) return;
    if (compareMode !== 'partial') {
      summary.style.display = 'none';
      return;
    }
    const arr = [...styles].sort((a, b) => a.localeCompare(b, 'ko', { numeric: true }));
    summary.style.display = '';
    if (!arr.length) {
      summary.textContent = '조사 스타일 0개 · 스캔으로 자동 선택됩니다.';
    } else {
      const listed = arr.slice(0, 12).join(', ');
      summary.textContent = `조사 스타일 ${arr.length}개 · ${listed}${arr.length > 12 ? ` 외 ${arr.length - 12}개` : ''}`;
    }
  }

  // 기존 기능을 그대로 호출하되 간이실사 모드에서만 변경된 집계를 사용합니다.
  render = function () {
    if (compareMode === 'partial') renderPartial();
    else {
      baseRender();
      updatePartialModeInfo(new Set());
    }
  };

  buildExportData = function (filterDiffOnly = false) {
    if (compareMode === 'full') return baseBuildExportData(filterDiffOnly);

    const data = [['구분', '스타일코드', '사이즈코드', '현재고', '실사', '차이']];
    for (const row of partialRows().rows) {
      if (filterDiffOnly && row.diff === 0) continue;
      if (row.type === 'unknown') {
        data.push(['미등록바코드', row.barcode, '-', '-', row.scanned, row.diff]);
      } else {
        data.push(['정상', row.styleCode, row.sizeCode, row.stock, row.scanned, row.diff]);
      }
    }
    return data;
  };

  exportExcel = function (filterDiffOnly = false) {
    if (compareMode === 'full') return baseExportExcel(filterDiffOnly);
    const { rows, styles } = partialRows();
    if (!rows.length) {
      alert('간이실사 결과가 없습니다. 바코드를 먼저 스캔하세요.');
      return;
    }
    if (!Object.keys(stockMap).length) {
      alert('현재고 파일을 먼저 업로드하세요.');
      return;
    }
    const data = buildExportData(filterDiffOnly);
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [
      { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 10 }, { wch: 10 }, { wch: 10 }
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, filterDiffOnly ? '차이내역' : '간이실사');
    const now = new Date();
    const stamp = [now.getFullYear(), String(now.getMonth() + 1).padStart(2, '0'), String(now.getDate()).padStart(2, '0')].join('');
    XLSX.writeFile(wb, `간이재고실사_${filterDiffOnly ? '차이내역' : '전체'}_${stamp}.xlsx`);
  };

  function switchMode(next) {
    compareMode = next === 'partial' ? 'partial' : 'full';
    localStorage.setItem(MODE_KEY, compareMode);
    document.querySelectorAll('.compare-mode-btn').forEach(button => {
      const active = button.dataset.mode === compareMode;
      button.classList.toggle('active', active);
      button.setAttribute('aria-pressed', active ? 'true' : 'false');
    });
    document.getElementById('modeHelp').textContent = compareMode === 'partial'
      ? '한 번이라도 찍은 스타일만 조사합니다. 선택된 스타일의 미스캔 사이즈는 차이가 계산되고, 다른 스타일은 제외됩니다.'
      : '현재고 파일의 모든 스타일·사이즈를 비교합니다.';
    document.getElementById('exportFull').textContent = compareMode === 'partial' ? '📊 간이실사 엑셀 저장' : '📊 전체 엑셀 저장';
    render();
    document.getElementById('barcodeInput').focus();
  }

  function installModeUI() {
    const scanningRow = document.querySelector('.scan-row');
    if (!scanningRow) throw new Error('RBO 원본 스캔 화면을 찾지 못했습니다.');

    const css = document.createElement('style');
    css.textContent = `
      .compare-mode { display:flex; flex-wrap:wrap; gap:7px; margin-bottom:8px; }
      .compare-mode-btn { flex:1 1 130px; min-height:44px; }
      .compare-mode-btn.active { background:var(--primary); color:#fff; border-color:var(--primary); }
      .compare-mode-help { margin:0 0 12px; font-size:12px; color:var(--muted); line-height:1.5; }
      #partialStyleSummary { margin:10px 0 5px; padding:9px 10px; border-radius:10px;
        background:var(--primary-soft); color:#165bb7; overflow-wrap:anywhere; font-size:12px; font-weight:650; }
    `;
    document.head.appendChild(css);

    const container = document.createElement('div');
    container.innerHTML = `
      <div class="compare-mode" role="group" aria-label="재고실사 방식">
        <button type="button" class="compare-mode-btn" data-mode="full">📦 전체 재고실사</button>
        <button type="button" class="compare-mode-btn" data-mode="partial">🔍 간이 재고실사</button>
      </div>
      <div id="modeHelp" class="compare-mode-help"></div>
    `;
    scanningRow.parentNode.insertBefore(container, scanningRow);

    const filters = document.querySelector('.filters');
    const summary = document.createElement('div');
    summary.id = 'partialStyleSummary';
    summary.style.display = 'none';
    filters.parentNode.insertBefore(summary, filters.nextSibling);

    document.querySelectorAll('.compare-mode-btn').forEach(button => {
      button.addEventListener('click', () => switchMode(button.dataset.mode));
    });
    switchMode(compareMode);
  }

  installModeUI();
})();
