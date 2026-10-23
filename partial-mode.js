/*
 * RBO 재고조사 - 간이 재고실사 스타일별 사이즈 상세 v1.4
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

  // 서로 다른 엑셀의 스타일/사이즈 문자열 표기를 통일합니다.
  // 예: 기준정보 rw0001 ↔ 현재고 RW0001, 전각/반각/공백 차이
  function normalizeCode(value) {
    return String(value ?? '').normalize('NFKC').replace(/\s+/g, '').toUpperCase();
  }

  function activeStyles() {
    const selected = new Map();
    for (const [barcode, qty] of Object.entries(scanCount)) {
      const info = masterMap[barcode];
      if (!info || Number(qty) <= 0) continue;
      const styleKey = normalizeCode(info.styleCode);
      if (styleKey && !selected.has(styleKey)) selected.set(styleKey, String(info.styleCode).trim());
    }
    return selected;
  }

  function partialRows() {
    // 조사 대상: 실제 스캔한 '스타일' 전체 (스캔한 사이즈만이 아님)
    const selected = activeStyles();
    const styles = new Set(selected.keys());
    const rowsByCode = new Map();
    const stylesFoundInStock = new Set();

    // 중요: 현재고 파일에서 스캔한 스타일의 *모든 사이즈*를 먼저 가져옵니다.
    // 미스캔 사이즈도 scanned=0으로 남겨 -현재고 차이가 나오게 합니다.
    for (const [originalKey, record] of Object.entries(stockMap)) {
      const { styleCode, sizeCode } = parseStyleSizeKey(originalKey);
      const styleKey = normalizeCode(styleCode);
      if (!styles.has(styleKey)) continue;
      stylesFoundInStock.add(styleKey);
      const rowKey = JSON.stringify([styleKey, normalizeCode(sizeCode)]);
      let row = rowsByCode.get(rowKey);
      if (!row) {
        row = {
          type: 'normal', styleCode: selected.get(styleKey) || String(styleCode),
          sizeCode: String(sizeCode ?? ''), scanned: 0, stock: 0, diff: 0
        };
        rowsByCode.set(rowKey, row);
      }
      row.stock += Number(record?.currentStock) || 0;
    }

    // 기준정보에 존재하지만 현재고 파일에 없는 사이즈는 실사 초과로 추가합니다.
    for (const [barcode, qty] of Object.entries(scanCount)) {
      const info = masterMap[barcode];
      const scanned = Number(qty);
      if (!info || !(scanned > 0)) continue;
      const styleKey = normalizeCode(info.styleCode);
      if (!styles.has(styleKey)) continue;
      const rowKey = JSON.stringify([styleKey, normalizeCode(info.sizeCode)]);
      let row = rowsByCode.get(rowKey);
      if (!row) {
        row = {
          type: 'normal', styleCode: selected.get(styleKey) || String(info.styleCode),
          sizeCode: String(info.sizeCode ?? ''), scanned: 0, stock: 0, diff: 0
        };
        rowsByCode.set(rowKey, row);
      }
      row.scanned += scanned;
    }

    const rows = [...rowsByCode.values()];
    rows.forEach(row => { row.diff = row.scanned - row.stock; });
    rows.sort((a, b) =>
      a.styleCode.localeCompare(b.styleCode, 'ko', { numeric: true }) ||
      a.sizeCode.localeCompare(b.sizeCode, 'ko', { numeric: true })
    );

    // 매칭하지 못한 바코드는 임의의 스타일로 배정하지 않습니다.
    for (const [barcode, qty] of Object.entries(unknownBarcodes)) {
      if (!(Number(qty) > 0)) continue;
      rows.push({ type: 'unknown', barcode, scanned: Number(qty), stock: null, diff: Number(qty) });
    }
    // 화면과 엑셀에서 완전히 동일한 집계/누락 사이즈를 사용합니다.
    const missingStockStyles = [...styles].filter(key => !stylesFoundInStock.has(key))
      .map(key => selected.get(key) || key);
    return { rows, styles: new Set([...selected.values()]), missingStockStyles };
  }

  // 스타일 전체의 사이즈별 수량을 그룹화합니다.
  // 총 차이(+/- 상쇄)와 부족수량을 혼동하지 않도록 부족/초과를 별도 합산합니다.
  function groupByStyle(rows) {
    const grouped = new Map();
    for (const row of rows) {
      if (row.type !== 'normal') continue;
      const key = normalizeCode(row.styleCode);
      if (!grouped.has(key)) {
        grouped.set(key, { styleCode: row.styleCode, rows: [], stock: 0, scanned: 0, shortage: 0, excess: 0, shortageSizes: 0 });
      }
      const group = grouped.get(key);
      group.rows.push(row);
      group.stock += row.stock;
      group.scanned += row.scanned;
      if (row.diff < 0) { group.shortage += -row.diff; group.shortageSizes++; }
      if (row.diff > 0) group.excess += row.diff;
    }
    return grouped;
  }

  function renderPartial() {
    const { rows, styles, missingStockStyles } = partialRows();
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

    const allGroups = groupByStyle(rows);
    let filteredRows = rows;
    if (currentFilter === 'scanned') {
      // 미실사: 조사 스타일에 존재하지만 찍지 않은 사이즈
      filteredRows = rows.filter(r => r.type === 'normal' && r.scanned === 0 && r.stock > 0);
    } else if (currentFilter === 'diff') {
      filteredRows = rows.filter(r => r.diff !== 0);
    } else if (currentFilter === 'shortage') {
      // 부족만: 일부만 찍은 사이즈까지 포함하여 부족한 모든 사이즈 표시
      filteredRows = rows.filter(r => r.type === 'normal' && r.diff < 0);
    }
    const filteredGroups = groupByStyle(filteredRows);
    const area = document.getElementById('resultArea');
    if (!styles.size && !rows.length) {
      area.innerHTML = '<div style="padding:14px;color:#8a929c;">아직 스캔한 스타일이 없습니다. 첫 바코드를 찍으면 해당 스타일의 모든 사이즈가 자동으로 비교됩니다.</div>';
    } else if (!filteredRows.length) {
      area.innerHTML = '<div style="padding:14px;color:#8a929c;">선택한 필터에 표시할 결과가 없습니다.</div>';
    } else {
      const panels = [...filteredGroups.entries()].map(([key, visible]) => {
        const full = allGroups.get(key);
        const shortageSizes = full.rows.filter(r => r.diff < 0)
          .map(r => `${escapeHtml(r.sizeCode || '-')} ${-r.diff}개`).join(' · ');
        const status = full.shortage
          ? `<strong class="rbo-shortage">부족 ${full.shortage.toLocaleString()}개 (${full.shortageSizes}개 사이즈)</strong>`
          : '<strong class="rbo-match">부족 없음</strong>';
        const extra = full.excess ? `<span class="rbo-excess">초과 ${full.excess.toLocaleString()}개</span>` : '';
        const rowHtml = visible.rows.map(r => {
          const result = r.diff < 0
            ? `<span class="rbo-shortage">부족 ${(-r.diff).toLocaleString()}개</span>`
            : r.diff > 0
              ? `<span class="rbo-excess">초과 ${r.diff.toLocaleString()}개</span>`
              : '<span class="rbo-match">일치</span>';
          return `<div class="rbo-size-row${r.diff < 0 ? ' rbo-size-row-short' : ''}">
            <div class="rbo-size-name">${escapeHtml(r.sizeCode || '-')}</div>
            <div class="rbo-count">${r.stock.toLocaleString()}</div>
            <div class="rbo-count">${r.scanned.toLocaleString()}</div>
            <div class="rbo-count rbo-result">${result}</div>
          </div>`;
        }).join('');
        return `<section class="rbo-style-panel" aria-label="${escapeHtml(full.styleCode)} 사이즈별 결과">
          <div class="rbo-style-heading">
            <strong class="rbo-style-name">${escapeHtml(full.styleCode)}</strong>
            <span class="rbo-style-stats">${status}${extra}</span>
          </div>
          <div class="rbo-style-totals">전산 ${full.stock.toLocaleString()}개 · 실사 ${full.scanned.toLocaleString()}개</div>
          ${full.shortage ? `<div class="rbo-shortage-detail">부족 사이즈 · ${shortageSizes}</div>` : ''}
          <div class="rbo-size-head"><span>사이즈</span><span>전산</span><span>실사</span><span>결과</span></div>
          ${rowHtml}
        </section>`;
      });
      const unknownHtml = filteredRows.filter(r => r.type === 'unknown').map(r =>
        `<div class="result-row unknown"><div class="code">미등록 바코드 · ${escapeHtml(r.barcode)}</div>
        <div class="nums">실사 ${r.scanned} · 현재고 - · <strong>확인 필요</strong></div></div>`
      ).join('');
      area.innerHTML = panels.join('') + unknownHtml;
    }

    updateSaveList();
    updateDataChips();
    updatePartialModeInfo(styles, missingStockStyles);
  }

  function updatePartialModeInfo(styles, missingStockStyles = []) {
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
      if (!Object.keys(stockMap).length) {
        summary.textContent += ' · ⚠️ 현재고 파일을 불러와야 누락 사이즈를 비교할 수 있습니다.';
      } else if (missingStockStyles.length) {
        summary.textContent += ` · ⚠️ 현재고에 없는 스타일: ${missingStockStyles.slice(0, 5).join(', ')}${missingStockStyles.length > 5 ? ' 외' : ''}`;
      }
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

    // 간이실사 엑셀은 기존 전체 실사와 동일한 6개 열만 사용합니다.
    // '차이'가 음수면 부족, 양수면 초과, 0이면 일치입니다.
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
    const { rows } = partialRows();
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
    // 모드 전환 시 숨겨져 있던 필터를 해제해서 전체 사이즈를 즉시 보여줍니다.
    currentFilter = 'all';
    document.querySelectorAll('.filter-btn').forEach(button => {
      button.classList.toggle('active', button.dataset.filter === 'all');
      if (button.dataset.filter === 'scanned') {
        button.textContent = compareMode === 'partial' ? '미실사' : '실사목록';
        button.title = compareMode === 'partial'
          ? '선택된 스타일에서 아직 스캔하지 않은 사이즈만 보기'
          : '실제로 스캔한 항목만 보기';
      }
    });
    const shortageButton = document.querySelector('.partial-shortage-filter');
    if (shortageButton) shortageButton.style.display = compareMode === 'partial' ? 'inline-block' : 'none';
    document.getElementById('modeHelp').textContent = compareMode === 'partial'
      ? '바코드를 하나라도 찍은 스타일의 전산재고 전체 사이즈를 비교합니다. 안 찍은 사이즈도 차이(-)로 표시됩니다. 스타일별로 사이즈·전산·실사·부족/초과를 표시합니다. 부족만 버튼으로 부족한 사이즈만 확인할 수 있습니다.'
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
      .rbo-style-panel { padding:0 0 10px; border-bottom:7px solid #f1f4f9; }
      .rbo-style-panel:last-child { border-bottom:0; }
      .rbo-style-heading { display:flex; gap:8px; flex-wrap:wrap; justify-content:space-between;
        align-items:center; padding:12px 10px 3px; }
      .rbo-style-name { font-size:16px; color:#18212f; overflow-wrap:anywhere; }
      .rbo-style-stats { display:flex; flex-wrap:wrap; gap:7px; align-items:center; font-size:12px; }
      .rbo-style-totals { font-size:12px; color:#6b7280; padding:2px 10px 6px; }
      .rbo-shortage-detail { font-size:12px; color:var(--danger,#c62828); padding:5px 9px;
        margin:0 9px 9px; background:#fff1f0; border-radius:6px; overflow-wrap:anywhere; }
      .rbo-shortage { color:var(--danger,#c62828); font-weight:800; }
      .rbo-excess { color:var(--plus,#1e66cc); font-weight:800; }
      .rbo-match { color:var(--ok,#16814a); font-weight:750; }
      .rbo-size-head, .rbo-size-row { display:grid;
        grid-template-columns:minmax(52px,1.1fr) minmax(41px,.7fr) minmax(41px,.7fr) minmax(82px,1.4fr);
        gap:5px; align-items:center; padding:9px 10px; font-variant-numeric:tabular-nums; }
      .rbo-size-head { font-size:11px; font-weight:750; background:#f3f6fa; color:#55606c; }
      .rbo-size-row { font-size:13px; border-bottom:1px solid #f0f1f3; }
      .rbo-size-row-short { background:#fff8f7; }
      .rbo-size-name { font-weight:800; }
      .rbo-count { text-align:right; white-space:nowrap; }
      .rbo-result { font-size:12px; }
      .partial-shortage-filter { display:none; }
      @media(max-width:380px) { .rbo-size-head, .rbo-size-row { padding:8px 6px; gap:3px; }
        .rbo-size-row { font-size:12px; } .rbo-result { font-size:11px; } }
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
      <div class="rbo-version-tag" style="font-size:11px;color:#165bb7;font-weight:800;margin:2px 0 6px;">✅ 간이실사 패치 v1.5 · 엑셀 간소화</div>
      <div id="modeHelp" class="compare-mode-help"></div>
    `;
    scanningRow.parentNode.insertBefore(container, scanningRow);

    const filters = document.querySelector('.filters');
    const shortageButton = document.createElement('button');
    shortageButton.type = 'button';
    shortageButton.className = 'filter-btn partial-shortage-filter';
    shortageButton.dataset.filter = 'shortage';
    shortageButton.textContent = '부족만';
    shortageButton.title = '해당 스타일에서 부족한 모든 사이즈와 각각 부족한 개수 보기';
    shortageButton.addEventListener('click', () => {
      if (compareMode !== 'partial') return;
      currentFilter = 'shortage';
      document.querySelectorAll('.filter-btn').forEach(button => {
        button.classList.toggle('active', button.dataset.filter === 'shortage');
      });
      render();
    });
    filters.appendChild(shortageButton);
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
