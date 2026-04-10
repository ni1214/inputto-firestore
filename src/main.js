import './style.css';
import { FIELD_DEFS, PROJECT_TEMPLATE, DRAWING_TEMPLATE, ROW_TEMPLATE } from './schema.js';
import { listProjects, loadProjectBundle, loadDrawingRows, loadProjectSymbols, saveDrawingBundle, saveDrawingSymbolBatch } from './store.js';
import {
  clearStoredGeminiApiKey,
  extractHandaiDataFromPdf,
  loadStoredGeminiApiKey,
  saveGeminiApiKey
} from './gemini.js';

const SAVE_ACTOR = 'system';
const SIDEBAR_STORAGE_KEY = 'inputto_sidebar_collapsed';
const MODE_STORAGE_KEY = 'inputto_active_mode';
const AUTO_SAVE_DELAY_MS = 1200;

const state = {
  env: 'production',
  activeMode: localStorage.getItem(MODE_STORAGE_KEY) || 'editor',
  sidebarCollapsed: localStorage.getItem(SIDEBAR_STORAGE_KEY) === 'true',
  loading: false,
  saving: false,
  geminiApiKey: loadStoredGeminiApiKey(),
  projects: [],
  project: { ...PROJECT_TEMPLATE },
  drawings: [],
  selectedDrawingId: '',
  drawing: { ...DRAWING_TEMPLATE },
  rows: [],
  selectedRowIds: new Set(),
  filterText: '',
  search: {
    projectC2: '',
    keyword: '',
    floor: '',
    insideOutside: '',
    resultCountText: '0件',
    hint: '検索結果はここに表示されます。',
    results: [],
    cacheProjectC2: '',
    cacheRows: []
  }
};

const elements = {
  appShell: document.getElementById('appShell'),
  sidebarToggle: document.getElementById('sidebarToggle'),
  navItems: Array.from(document.querySelectorAll('[data-mode]')),
  modePanels: Array.from(document.querySelectorAll('[data-mode-panel]')),
  projectSelect: document.getElementById('projectSelect'),
  refreshProjectsButton: document.getElementById('refreshProjectsButton'),
  newProjectButton: document.getElementById('newProjectButton'),
  projectC2Input: document.getElementById('projectC2Input'),
  projectNameInput: document.getElementById('projectNameInput'),
  projectShortNameInput: document.getElementById('projectShortNameInput'),
  projectContactInput: document.getElementById('projectContactInput'),
  drawingNumberInput: document.getElementById('drawingNumberInput'),
  drawingStatusInput: document.getElementById('drawingStatusInput'),
  loadDrawingButton: document.getElementById('loadDrawingButton'),
  drawingTabs: document.getElementById('drawingTabs'),
  bulkSymbolsInput: document.getElementById('bulkSymbolsInput'),
  applyBulkSymbolsButton: document.getElementById('applyBulkSymbolsButton'),
  clearBulkSymbolsButton: document.getElementById('clearBulkSymbolsButton'),
  pdfDropZone: document.getElementById('pdfDropZone'),
  pdfDropHint: document.getElementById('pdfDropHint'),
  pdfFileInput: document.getElementById('pdfFileInput'),
  setGeminiKeyButton: document.getElementById('setGeminiKeyButton'),
  filterInput: document.getElementById('filterInput'),
  addRowButton: document.getElementById('addRowButton'),
  duplicateRowButton: document.getElementById('duplicateRowButton'),
  deleteRowsButton: document.getElementById('deleteRowsButton'),
  tableHead: document.getElementById('tableHead'),
  tableBody: document.getElementById('tableBody'),
  searchProjectSelect: document.getElementById('searchProjectSelect'),
  searchKeywordInput: document.getElementById('searchKeywordInput'),
  searchFloorInput: document.getElementById('searchFloorInput'),
  searchInsideOutsideInput: document.getElementById('searchInsideOutsideInput'),
  searchButton: document.getElementById('searchButton'),
  searchResultCount: document.getElementById('searchResultCount'),
  searchResultHint: document.getElementById('searchResultHint'),
  searchResultsBody: document.getElementById('searchResultsBody'),
  summaryProject: document.getElementById('summaryProject'),
  summaryDrawing: document.getElementById('summaryDrawing'),
  summaryRows: document.getElementById('summaryRows'),
  summaryLabels: document.getElementById('summaryLabels'),
  inspectionTableBody: document.getElementById('inspectionTableBody'),
  labelPages: document.getElementById('labelPages'),
  printButton: document.getElementById('printButton'),
  toastHost: document.getElementById('toastHost')
};

let autoSaveTimerId = null;
let lastSavedSignature = '';

function createUiRow(overrides = {}) {
  return {
    ...ROW_TEMPLATE,
    uiId: crypto.randomUUID(),
    ...overrides
  };
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function normalizeText(value) {
  return String(value || '')
    .normalize('NFKC')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');
}

function setBusy(mode, message) {
  state.loading = mode === 'loading';
  state.saving = mode === 'saving';
  if (elements.busyState) {
    elements.busyState.textContent = mode === 'loading' ? '読み込み中' : mode === 'saving' ? '保存中' : '待機中';
    elements.busyState.dataset.mode = mode || 'idle';
  }
  if (elements.statusText) {
    elements.statusText.textContent = message;
  }
}

function setStatus(message) {
  setBusy('', message);
}

function showToast(message, kind = 'info') {
  if (!elements.toastHost || !message) {
    return;
  }
  const toast = document.createElement('div');
  toast.className = `toast${kind === 'error' ? ' is-error' : kind === 'success' ? ' is-success' : ''}`;
  toast.textContent = message;
  elements.toastHost.appendChild(toast);
  window.setTimeout(() => {
    toast.remove();
  }, 3200);
}

function setActiveMode(mode) {
  state.activeMode = mode;
  localStorage.setItem(MODE_STORAGE_KEY, mode);
  renderChrome();
}

function toggleSidebar() {
  state.sidebarCollapsed = !state.sidebarCollapsed;
  localStorage.setItem(SIDEBAR_STORAGE_KEY, String(state.sidebarCollapsed));
  renderChrome();
}

function renderChrome() {
  elements.appShell.classList.toggle('sidebar-collapsed', state.sidebarCollapsed);
  elements.navItems.forEach((item) => item.classList.toggle('is-active', item.dataset.mode === state.activeMode));
  elements.modePanels.forEach((panel) => panel.classList.toggle('is-active', panel.dataset.modePanel === state.activeMode));
}

function syncProjectFromForm() {
  state.project = {
    c2: elements.projectC2Input.value.trim(),
    projectName: elements.projectNameInput.value.trim(),
    shortName: elements.projectShortNameInput.value.trim(),
    contact: elements.projectContactInput.value.trim()
  };
}

function syncDrawingFromForm() {
  state.drawing = {
    ...state.drawing,
    drawingNumber: elements.drawingNumberInput.value.trim(),
    drawingStatus: elements.drawingStatusInput.value.trim()
  };
}

function syncSearchFromForm() {
  state.search.projectC2 = elements.searchProjectSelect.value.trim();
  state.search.keyword = elements.searchKeywordInput.value.trim();
  state.search.floor = elements.searchFloorInput.value.trim();
  state.search.insideOutside = elements.searchInsideOutsideInput.value.trim();
}

function updateFormInputs() {
  elements.projectC2Input.value = state.project.c2 || '';
  elements.projectNameInput.value = state.project.projectName || '';
  elements.projectShortNameInput.value = state.project.shortName || '';
  elements.projectContactInput.value = state.project.contact || '';
  elements.drawingNumberInput.value = state.drawing.drawingNumber || '';
  elements.drawingStatusInput.value = state.drawing.drawingStatus || '';
  elements.filterInput.value = state.filterText || '';
  elements.searchProjectSelect.value = state.search.projectC2 || '';
  elements.searchKeywordInput.value = state.search.keyword || '';
  elements.searchFloorInput.value = state.search.floor || '';
  elements.searchInsideOutsideInput.value = state.search.insideOutside || '';
}

function buildProjectOptions(selectedValue) {
  return ['<option value="">工事を選択</option>']
    .concat(
      state.projects.map((project) => {
        const label = [project.projectName || '-', `工事 ${project.c2}`];
        if (project.drawingCount || project.symbolCount) {
          label.push(`手配書 ${project.drawingCount || 0}件 / 符号 ${project.symbolCount || 0}件`);
        }
        const selected = selectedValue && selectedValue === project.c2 ? ' selected' : '';
        return `<option value="${escapeHtml(project.c2)}"${selected}>${escapeHtml(label.join(' / '))}</option>`;
      })
    )
    .join('');
}

function renderProjectSelects() {
  elements.projectSelect.innerHTML = buildProjectOptions(state.project.c2 || '');
  elements.searchProjectSelect.innerHTML = buildProjectOptions(state.search.projectC2 || '');
}

function renderDrawingTabs() {
  if (!state.project.c2) {
    elements.drawingTabs.innerHTML = '<p class="empty-text">工事を選ぶと手配書タブが出ます。</p>';
    return;
  }
  if (!state.drawings.length) {
    elements.drawingTabs.innerHTML = '<p class="empty-text">まだ手配書がありません。手配書Noを入れて読込すると、新しい手配書として入力できます。</p>';
    return;
  }

  elements.drawingTabs.innerHTML = state.drawings
    .map((drawing) => {
      const active = drawing.id === state.selectedDrawingId ? ' is-active' : '';
      return `
        <button type="button" class="drawing-chip${active}" data-drawing-id="${escapeHtml(drawing.id)}">
          <span>${escapeHtml(drawing.drawingNumber || '-')}</span>
          <small>${escapeHtml(String(drawing.rowCount || 0))}件</small>
        </button>
      `;
    })
    .join('');
}

function renderTableHead() {
  const headCells = ['<th class="checkbox-cell"><input id="toggleAllRows" type="checkbox"></th>']
    .concat(FIELD_DEFS.map((field) => `<th style="min-width:${field.width};">${escapeHtml(field.label)}</th>`))
    .join('');
  elements.tableHead.innerHTML = `<tr>${headCells}</tr>`;
}

function matchesFilter(row) {
  const term = state.filterText.trim().toLowerCase();
  if (!term) {
    return true;
  }
  return [row.symbol, row.name, row.floor, row.insideOutside, row.bakeColor]
    .join(' ')
    .toLowerCase()
    .includes(term);
}

function hasRowContent(row) {
  return FIELD_DEFS.some((field) => String(row[field.key] || '').trim() !== '');
}

function getMeaningfulRows(rows = state.rows) {
  return rows.filter(hasRowContent);
}

function hasIncompleteRows(rows = state.rows) {
  return getMeaningfulRows(rows).some((row) => !String(row.symbol || '').trim());
}

function normalizeSymbolInput(value) {
  return String(value || '')
    .normalize('NFKC')
    .trim()
    .replace(/\s+/g, '');
}

function expandSymbolRange(value) {
  const normalized = normalizeSymbolInput(value);
  const rangeMatch = normalized.match(/^(.*?)(\d+)\s*[~〜～]\s*(.*?)(\d+)$/);
  if (!rangeMatch) {
    return [normalized].filter(Boolean);
  }

  const [, prefixLeft, startText, prefixRight, endText] = rangeMatch;
  if (normalizeText(prefixLeft) !== normalizeText(prefixRight)) {
    return [normalized].filter(Boolean);
  }

  const startNumber = Number(startText);
  const endNumber = Number(endText);
  if (!Number.isFinite(startNumber) || !Number.isFinite(endNumber) || endNumber < startNumber) {
    return [normalized].filter(Boolean);
  }

  const padWidth = startText.startsWith('0') ? startText.length : 0;
  const items = [];
  for (let number = startNumber; number <= endNumber; number += 1) {
    const suffix = padWidth ? String(number).padStart(padWidth, '0') : String(number);
    items.push(`${prefixLeft}${suffix}`);
  }
  return items.filter(Boolean);
}

function parseBulkSymbols(text) {
  const seen = new Set();
  const symbols = [];

  String(text || '')
    .split(/[\n,、]+/)
    .map((line) => normalizeSymbolInput(line))
    .filter(Boolean)
    .forEach((line) => {
      expandSymbolRange(line).forEach((symbol) => {
        const normalized = normalizeText(symbol);
        if (!normalized || seen.has(normalized)) {
          return;
        }
        seen.add(normalized);
        symbols.push(symbol);
      });
    });

  return symbols;
}

function syncBulkSymbolsInput(rows = state.rows) {
  if (!elements.bulkSymbolsInput) {
    return;
  }
  elements.bulkSymbolsInput.value = getMeaningfulRows(rows)
    .map((row) => String(row.symbol || '').trim())
    .filter(Boolean)
    .join('\n');
}

function buildPersistedSnapshot() {
  syncProjectFromForm();
  syncDrawingFromForm();
  return {
    project: {
      c2: state.project.c2 || '',
      projectName: state.project.projectName || '',
      shortName: state.project.shortName || '',
      contact: state.project.contact || ''
    },
    drawing: {
      id: state.selectedDrawingId || '',
      drawingNumber: state.drawing.drawingNumber || '',
      drawingStatus: state.drawing.drawingStatus || ''
    },
    rows: getMeaningfulRows().map((row) => {
      const payload = { docId: row.docId || '' };
      FIELD_DEFS.forEach((field) => {
        payload[field.key] = row[field.key] || '';
      });
      return payload;
    })
  };
}

function computeSaveSignature() {
  return JSON.stringify(buildPersistedSnapshot());
}

function syncSavedSignature() {
  lastSavedSignature = computeSaveSignature();
}

function clearAutoSaveTimer() {
  if (autoSaveTimerId) {
    clearTimeout(autoSaveTimerId);
    autoSaveTimerId = null;
  }
}

function canAutoSave() {
  const projectCode = String(state.project.c2 || '').trim();
  const projectName = String(state.project.projectName || '').trim();
  const drawingNumber = String(state.drawing.drawingNumber || '').trim();
  const meaningfulRows = getMeaningfulRows();

  if (!projectCode || !projectName) {
    return false;
  }
  if (!drawingNumber && meaningfulRows.length) {
    return false;
  }
  if (hasIncompleteRows(meaningfulRows)) {
    return false;
  }
  return true;
}

function scheduleAutoSave() {
  clearAutoSaveTimer();
  if (!canAutoSave()) {
    return;
  }
  const nextSignature = computeSaveSignature();
  if (nextSignature === lastSavedSignature) {
    return;
  }
  autoSaveTimerId = window.setTimeout(() => {
    autoSaveTimerId = null;
    void saveCurrent({ auto: true });
  }, AUTO_SAVE_DELAY_MS);
}

async function applyBulkSymbols() {
  const symbols = parseBulkSymbols(elements.bulkSymbolsInput?.value || '');
  syncProjectFromForm();
  syncDrawingFromForm();
  if (!symbols.length && !state.selectedDrawingId) {
    setStatus('追加する符号がありません。');
    return;
  }

  if (!state.project.c2 || !state.project.projectName) {
    showToast('先に工事を入れてください。', 'error');
    return;
  }
  if (!state.drawing.drawingNumber) {
    showToast('先に手配書Noを入れてください。', 'error');
    return;
  }

  try {
    const response = await saveDrawingSymbolBatch({
      env: state.env,
      project: state.project,
      drawing: { ...state.drawing, id: state.selectedDrawingId },
      symbols,
      operator: SAVE_ACTOR
    });

    state.project = {
      ...state.project,
      ...response.project
    };
    if (response.drawings) {
      state.drawings = response.drawings.sort((left, right) => String(left.drawingNumber || '').localeCompare(String(right.drawingNumber || ''), 'ja', { numeric: true }));
    }
    if (response.drawing) {
      state.selectedDrawingId = response.drawing.id;
      state.drawing = {
        ...state.drawing,
        ...response.drawing
      };
    }
    state.rows = Array.isArray(response.rows) && response.rows.length ? response.rows : [createUiRow()];
    state.selectedRowIds = new Set();
    syncBulkSymbolsInput(state.rows);
    renderAll();
    syncSavedSignature();
    await refreshProjects();
    showToast('符号を反映しました。', 'success');
    return;
  } catch (error) {
    console.error(error);
    showToast(`符号の反映に失敗しました: ${error.message}`, 'error');
    return;
  }

  const existingRows = getMeaningfulRows();
  const nextRows = existingRows.length ? state.rows.filter(hasRowContent) : [];
  const seen = new Set(nextRows.map((row) => normalizeText(row.symbol)));

  symbols.forEach((symbol) => {
    const normalized = normalizeText(symbol);
    if (!normalized || seen.has(normalized)) {
      return;
    }
    seen.add(normalized);
    nextRows.push(createUiRow({ symbol }));
  });

  state.rows = nextRows.length ? nextRows : [createUiRow()];
  state.selectedRowIds = new Set();
  renderRows();
  if (elements.bulkSymbolsInput) {
    elements.bulkSymbolsInput.value = '';
  }
  setStatus(`${symbols.length}件の符号を追加しました。`);
  showToast(`${symbols.length}件の符号を追加しました。`, 'success');
  scheduleAutoSave();
}

function clearBulkSymbols() {
  if (elements.bulkSymbolsInput) {
    elements.bulkSymbolsInput.value = '';
    elements.bulkSymbolsInput.focus();
  }
  showToast('一括入力を空にしました。');
}

function promptForGeminiKey() {
  const nextKey = window.prompt('Gemini APIキーを入力してください', state.geminiApiKey || '');
  if (nextKey === null) {
    return;
  }
  if (!String(nextKey).trim()) {
    clearStoredGeminiApiKey();
    state.geminiApiKey = '';
    showToast('Gemini APIキーを削除しました。');
    return;
  }
  const savedKey = saveGeminiApiKey(nextKey);
  state.geminiApiKey = savedKey;
  showToast('Gemini APIキーを保存しました。', 'success');
}

async function handlePdfImport(file) {
  if (!file) {
    return;
  }

  if (!state.geminiApiKey) {
    const nextKey = window.prompt('Gemini APIキーを入力してください');
    if (!nextKey) {
      showToast('Gemini APIキーがないためPDFを読み込めません。', 'error');
      return;
    }
    state.geminiApiKey = saveGeminiApiKey(nextKey);
  }

  setBusy('loading', 'PDFを解析しています...');
  elements.pdfDropZone?.classList.add('is-busy');
  try {
    const extracted = await extractHandaiDataFromPdf(file, {
      apiKey: state.geminiApiKey
    });

    const bundle = extracted.project.c2
      ? await loadProjectBundle(state.env, extracted.project.c2)
      : { project: { ...PROJECT_TEMPLATE }, drawings: [] };

    state.project = {
      c2: extracted.project.c2 || bundle.project.c2 || '',
      projectName: extracted.project.projectName || bundle.project.projectName || '',
      shortName: extracted.project.shortName || bundle.project.shortName || '',
      contact: extracted.project.contact || bundle.project.contact || ''
    };
    state.drawings = bundle.drawings || [];
    state.drawing = {
      ...DRAWING_TEMPLATE,
      drawingNumber: extracted.drawing.drawingNumber || '',
      drawingStatus: extracted.drawing.drawingStatus || ''
    };
    state.selectedDrawingId =
      state.drawings.find((item) => item.drawingNumber === state.drawing.drawingNumber)?.id || '';
    state.rows = extracted.rows.length
      ? extracted.rows.map((row) => createUiRow({ ...row }))
      : [createUiRow()];
    state.selectedRowIds = new Set();
    syncBulkSymbolsInput(state.rows);
    if (!state.search.projectC2) {
      state.search.projectC2 = state.project.c2;
    }
    renderAll();
    setActiveMode('report');
    syncSavedSignature();
    scheduleAutoSave();
    showToast('PDFの内容を読み込みました。', 'success');
  } catch (error) {
    console.error(error);
    showToast(`PDFの読み込みに失敗しました: ${error.message}`, 'error');
  } finally {
    elements.pdfDropZone?.classList.remove('is-busy');
    setBusy('', '');
  }
}

function getReportRows() {
  return state.rows.filter(hasRowContent);
}

function buildLabelItems(rows) {
  const shortName = state.project.shortName || state.project.projectName || '未設定';
  const drawingLabel = state.drawing.drawingNumber ? `手配書 ${state.drawing.drawingNumber}` : '';
  const labels = [];

  rows.forEach((row) => {
    const copies = Number(row.labelCount || 0);
    if (!copies || !row.symbol) {
      return;
    }
    const line1 = [shortName, drawingLabel, row.floor, row.insideOutside].filter(Boolean).join(' ');
    const line2 = [row.symbol, row.name].filter(Boolean).join(' ');
    const line3 = row.bakeColor ? `焼付色 ${row.bakeColor}` : '';

    for (let count = 0; count < copies; count += 1) {
      labels.push({ line1, line2, line3 });
    }
  });

  return labels;
}

function renderLabelPages(rows) {
  const labels = buildLabelItems(rows);
  if (!labels.length) {
    elements.labelPages.innerHTML = '<p class="empty-text">ラベル枚数が入った行だけここに並びます。</p>';
    return;
  }

  const pages = [];
  for (let index = 0; index < labels.length; index += 20) {
    pages.push(labels.slice(index, index + 20));
  }

  elements.labelPages.innerHTML = pages
    .map((page) => {
      const items = Array.from({ length: 20 }, (_, index) => page[index] || null)
        .map((label) => {
          if (!label) {
            return '<article class="label-card label-card-empty"></article>';
          }
          return `
            <article class="label-card">
              <div class="label-line label-line-top">${escapeHtml(label.line1)}</div>
              <div class="label-line label-line-middle">${escapeHtml(label.line2)}</div>
              <div class="label-line label-line-bottom">${escapeHtml(label.line3)}</div>
            </article>
          `;
        })
        .join('');
      return `<section class="label-sheet">${items}</section>`;
    })
    .join('');
}

function renderSummary(rows) {
  const labels = buildLabelItems(rows);
  elements.summaryProject.textContent = state.project.shortName || state.project.projectName || '-';
  elements.summaryDrawing.textContent = state.drawing.drawingNumber || '-';
  elements.summaryRows.textContent = `${rows.length}件`;
  elements.summaryLabels.textContent = `${labels.length}枚`;
  renderLabelPages(rows);
}

function renderInspectionTable(rows) {
  if (!rows.length) {
    elements.inspectionTableBody.innerHTML = '<tr><td colspan="8" class="empty-text">手配書を読み込むと検査表が表示されます。</td></tr>';
    return;
  }

  elements.inspectionTableBody.innerHTML = rows
    .map((row) => `
      <tr>
        <td>${escapeHtml(state.drawing.drawingNumber || '-')}</td>
        <td>${escapeHtml(row.symbol || '')}</td>
        <td>${escapeHtml(row.name || '')}</td>
        <td>${escapeHtml(row.floor || '')}</td>
        <td>${escapeHtml(row.insideOutside || '')}</td>
        <td>${escapeHtml(row.labelCount || '')}</td>
        <td>${escapeHtml(row.bakeColor || '')}</td>
        <td>${escapeHtml(row.draftAssignee || '')}</td>
      </tr>
    `)
    .join('');
}

function renderReport() {
  const reportRows = getReportRows();
  renderSummary(reportRows);
  renderInspectionTable(reportRows);
}

function renderRows() {
  const visibleRows = state.rows.filter(matchesFilter);
  if (!visibleRows.length) {
    elements.tableBody.innerHTML = '<tr><td colspan="34" class="empty-text">表示できる符号がありません。</td></tr>';
    renderReport();
    return;
  }

  elements.tableBody.innerHTML = visibleRows
    .map((row) => {
      const checkbox = `
        <td class="checkbox-cell">
          <input type="checkbox" data-row-select="${escapeHtml(row.uiId)}"${state.selectedRowIds.has(row.uiId) ? ' checked' : ''}>
        </td>
      `;
      const cells = FIELD_DEFS.map((field) => {
        const value = row[field.key] || '';
        const inputType = field.type || 'text';
        const step = inputType === 'number' ? ' step="1"' : '';
        return `
          <td>
            <input
              class="table-input"
              data-row-id="${escapeHtml(row.uiId)}"
              data-field-key="${escapeHtml(field.key)}"
              type="${escapeHtml(inputType)}"
              value="${escapeHtml(value)}"${step}>
          </td>
        `;
      }).join('');
      return `<tr>${checkbox}${cells}</tr>`;
    })
    .join('');

  renderReport();
}

function renderSearchResults() {
  elements.searchResultCount.textContent = state.search.resultCountText;
  elements.searchResultHint.textContent = state.search.hint;

  if (!state.search.results.length) {
    elements.searchResultsBody.innerHTML = '<tr><td colspan="7" class="empty-text">一致するデータがありません。</td></tr>';
    return;
  }

  elements.searchResultsBody.innerHTML = state.search.results
    .map((row, index) => `
      <tr>
        <td>${escapeHtml(row.drawingNumber || '')}</td>
        <td>${escapeHtml(row.symbol || '')}</td>
        <td>${escapeHtml(row.name || '')}</td>
        <td>${escapeHtml(row.floor || '')}</td>
        <td>${escapeHtml(row.insideOutside || '')}</td>
        <td>${escapeHtml(row.bakeColor || '')}</td>
        <td>
          <button type="button" class="ghost-button compact-button" data-open-search-result="${index}">
            <span class="material-symbols-outlined">open_in_new</span>
            <span>編集で開く</span>
          </button>
        </td>
      </tr>
    `)
    .join('');
}

function renderAll() {
  updateFormInputs();
  renderChrome();
  renderProjectSelects();
  renderDrawingTabs();
  renderRows();
  renderSearchResults();
}

async function refreshProjects() {
  setBusy('loading', '工事一覧を読み込んでいます。');
  try {
    state.projects = await listProjects(state.env);
    renderProjectSelects();
    setStatus('工事一覧を更新しました。');
  } catch (error) {
    console.error(error);
    setStatus(`工事一覧の読込に失敗しました: ${error.message}`);
  }
}

function resetDrawingState() {
  state.selectedDrawingId = '';
  state.drawing = { ...DRAWING_TEMPLATE };
  state.rows = [];
  state.selectedRowIds = new Set();
  syncBulkSymbolsInput([]);
}

function clearProjectState() {
  clearAutoSaveTimer();
  state.project = { ...PROJECT_TEMPLATE };
  state.drawings = [];
  resetDrawingState();
  syncSavedSignature();
  renderAll();
}

async function selectProject(c2) {
  if (!c2) {
    clearProjectState();
    renderAll();
    return;
  }

  setBusy('loading', `工事 ${c2} を読み込んでいます。`);
  try {
    const { project, drawings } = await loadProjectBundle(state.env, c2);
    state.project = {
      c2: project.c2 || '',
      projectName: project.projectName || '',
      shortName: project.shortName || '',
      contact: project.contact || ''
    };
    state.drawings = drawings;
    if (!state.search.projectC2) {
      state.search.projectC2 = state.project.c2;
    }
    resetDrawingState();
    renderAll();
    syncSavedSignature();
    setStatus(`工事 ${c2} を読み込みました。`);
  } catch (error) {
    console.error(error);
    setStatus(`工事の読込に失敗しました: ${error.message}`);
  }
}

async function loadCurrentDrawing() {
  syncProjectFromForm();
  syncDrawingFromForm();

  if (!state.project.c2) {
    setStatus('先に工事番号を入れてください。');
    return;
  }
  if (!state.drawing.drawingNumber) {
    setStatus('手配書Noを入れてください。');
    return;
  }

  const drawing = state.drawings.find((item) => item.drawingNumber === state.drawing.drawingNumber);
  const drawingId = drawing?.id || '';

  setBusy('loading', `手配書No ${state.drawing.drawingNumber} の符号を読み込んでいます。`);
  try {
    state.selectedDrawingId = drawingId;
    state.drawing = {
      id: drawingId,
      drawingNumber: state.drawing.drawingNumber,
      drawingStatus: drawing?.drawingStatus || state.drawing.drawingStatus || ''
    };
    state.rows = drawingId ? await loadDrawingRows(state.env, state.project.c2, drawingId) : [];
    state.rows = state.rows.length ? state.rows : [createUiRow()];
    state.selectedRowIds = new Set();
    syncBulkSymbolsInput(state.rows);
    renderAll();
    syncSavedSignature();
    if (drawingId) {
      setStatus(`登録済みの手配書No ${state.drawing.drawingNumber} を読み込みました。`);
    } else {
      setStatus(`手配書No ${state.drawing.drawingNumber} は未登録です。そのまま入力して保存できます。`);
    }
  } catch (error) {
    console.error(error);
    setStatus(`手配書の読込に失敗しました: ${error.message}`);
  }
}

async function saveCurrent(options = {}) {
  const { auto = false } = options;
  syncProjectFromForm();
  syncDrawingFromForm();

  if (!canAutoSave()) {
    return;
  }
  const nextSignature = computeSaveSignature();
  if (auto && nextSignature === lastSavedSignature) {
    return;
  }

  setBusy('saving', '保存しています。');
  try {
    const response = await saveDrawingBundle({
      env: state.env,
      project: state.project,
      drawing: { ...state.drawing, id: state.selectedDrawingId },
      rows: state.rows,
      operator: SAVE_ACTOR
    });

    state.project = {
      ...state.project,
      ...response.project
    };
    if (response.drawings) {
      state.drawings = response.drawings.sort((left, right) => String(left.drawingNumber || '').localeCompare(String(right.drawingNumber || ''), 'ja', { numeric: true }));
    }
    if (response.drawing) {
      state.selectedDrawingId = response.drawing.id;
      state.drawing = {
        ...state.drawing,
        ...response.drawing
      };
    }
    state.rows = Array.isArray(response.rows) ? response.rows : state.rows;
    if (!state.rows.length) {
      state.rows = [createUiRow()];
    }
    state.selectedRowIds = new Set();
    syncBulkSymbolsInput(state.rows);
    renderAll();
    syncSavedSignature();
    await refreshProjects();
    setStatus(auto ? '自動保存しました。' : '保存しました。');
  } catch (error) {
    console.error(error);
    setStatus(`保存に失敗しました: ${error.message}`);
  }
}

function addRow() {
  state.rows.push(createUiRow());
  renderRows();
  scheduleAutoSave();
}

function duplicateSelectedRows() {
  const targets = state.rows.filter((row) => state.selectedRowIds.has(row.uiId));
  if (!targets.length) {
    setStatus('複製する行を選択してください。');
    return;
  }

  targets.forEach((row) => {
    state.rows.push(
      createUiRow({
        ...row,
        docId: '',
        symbol: ''
      })
    );
  });
  state.selectedRowIds = new Set();
  renderRows();
  setStatus(`${targets.length}行を複製しました。`);
  scheduleAutoSave();
}

function deleteSelectedRows() {
  if (!state.selectedRowIds.size) {
    setStatus('削除する行を選択してください。');
    return;
  }
  state.rows = state.rows.filter((row) => !state.selectedRowIds.has(row.uiId));
  if (!state.rows.length) {
    state.rows = [createUiRow()];
  }
  state.selectedRowIds = new Set();
  renderRows();
  setStatus('選択行を削除しました。');
  scheduleAutoSave();
}

async function performSearch() {
  syncSearchFromForm();

  if (!state.search.projectC2) {
    state.search.results = [];
    state.search.resultCountText = '0件';
    state.search.hint = '先に工事を選択してください。';
    renderSearchResults();
    setStatus('検索する工事を選択してください。');
    return;
  }

  try {
    if (state.search.cacheProjectC2 !== state.search.projectC2) {
      setBusy('loading', `工事 ${state.search.projectC2} の検索用データを読み込んでいます。`);
      state.search.cacheRows = await loadProjectSymbols(state.env, state.search.projectC2);
      state.search.cacheProjectC2 = state.search.projectC2;
    } else {
      setBusy('loading', '検索条件を絞り込んでいます。');
    }

    const keyword = normalizeText(state.search.keyword);
    const floor = normalizeText(state.search.floor);
    const insideOutside = normalizeText(state.search.insideOutside);

    state.search.results = state.search.cacheRows.filter((row) => {
      const keywordOk = !keyword || normalizeText([
        row.drawingNumber,
        row.symbol,
        row.name,
        row.floor,
        row.insideOutside,
        row.bakeColor,
        row.drawingStatus
      ].join(' ')).includes(keyword);
      const floorOk = !floor || normalizeText(row.floor).includes(floor);
      const inoutOk = !insideOutside || normalizeText(row.insideOutside).includes(insideOutside);
      return keywordOk && floorOk && inoutOk;
    });

    state.search.resultCountText = `${state.search.results.length}件`;
    state.search.hint = `${state.search.projectC2} の中だけを検索しています。`;
    renderSearchResults();
    setStatus(`${state.search.results.length}件の検索結果を表示しました。`);
  } catch (error) {
    console.error(error);
    state.search.results = [];
    state.search.resultCountText = '0件';
    state.search.hint = '検索中にエラーが発生しました。';
    renderSearchResults();
    setStatus(`検索に失敗しました: ${error.message}`);
  }
}

async function openSearchResult(index) {
  const result = state.search.results[Number(index)];
  if (!result) {
    return;
  }

  await selectProject(result.c2);
  state.selectedDrawingId = result.drawingId || '';
  state.drawing = {
    id: result.drawingId || '',
    drawingNumber: result.drawingNumber || '',
    drawingStatus: result.drawingStatus || ''
  };
  updateFormInputs();
  await loadCurrentDrawing();
  setActiveMode('report');
}

function bindEvents() {
  elements.sidebarToggle.addEventListener('click', toggleSidebar);

  elements.navItems.forEach((item) => {
    item.addEventListener('click', () => setActiveMode(item.dataset.mode));
  });

  elements.projectSelect.addEventListener('change', async () => {
    await selectProject(elements.projectSelect.value);
  });

  elements.refreshProjectsButton.addEventListener('click', refreshProjects);
  elements.newProjectButton.addEventListener('click', () => {
    clearProjectState();
    elements.projectSelect.value = '';
    elements.projectC2Input.focus();
    setStatus('新規工事入力に切り替えました。');
  });
  elements.loadDrawingButton.addEventListener('click', loadCurrentDrawing);
  elements.applyBulkSymbolsButton.addEventListener('click', applyBulkSymbols);
  elements.clearBulkSymbolsButton.addEventListener('click', clearBulkSymbols);
  elements.addRowButton.addEventListener('click', addRow);
  elements.duplicateRowButton.addEventListener('click', duplicateSelectedRows);
  elements.deleteRowsButton.addEventListener('click', deleteSelectedRows);
  elements.filterInput.addEventListener('input', () => {
    state.filterText = elements.filterInput.value;
    renderRows();
  });

  elements.searchProjectSelect.addEventListener('change', () => {
    syncSearchFromForm();
  });
  elements.searchKeywordInput.addEventListener('input', syncSearchFromForm);
  elements.searchFloorInput.addEventListener('input', syncSearchFromForm);
  elements.searchInsideOutsideInput.addEventListener('input', syncSearchFromForm);
  elements.searchButton.addEventListener('click', performSearch);
  elements.printButton.addEventListener('click', () => {
    setActiveMode('report');
    window.print();
  });

  elements.drawingTabs.addEventListener('click', async (event) => {
    const button = event.target.closest('[data-drawing-id]');
    if (!button) {
      return;
    }
    const drawing = state.drawings.find((item) => item.id === button.dataset.drawingId);
    if (!drawing) {
      return;
    }
    state.selectedDrawingId = drawing.id;
    state.drawing = {
      id: drawing.id,
      drawingNumber: drawing.drawingNumber || '',
      drawingStatus: drawing.drawingStatus || ''
    };
    updateFormInputs();
    await loadCurrentDrawing();
  });

  elements.tableBody.addEventListener('input', (event) => {
    const input = event.target.closest('[data-row-id][data-field-key]');
    if (!input) {
      return;
    }
    const row = state.rows.find((item) => item.uiId === input.dataset.rowId);
    if (!row) {
      return;
    }
    row[input.dataset.fieldKey] = input.value;
    renderReport();
    scheduleAutoSave();
  });

  elements.tableBody.addEventListener('change', (event) => {
    const checkbox = event.target.closest('[data-row-select]');
    if (!checkbox) {
      return;
    }
    const rowId = checkbox.dataset.rowSelect;
    if (checkbox.checked) {
      state.selectedRowIds.add(rowId);
    } else {
      state.selectedRowIds.delete(rowId);
    }
  });

  elements.tableHead.addEventListener('change', (event) => {
    const checkbox = event.target.closest('#toggleAllRows');
    if (!checkbox) {
      return;
    }
    const visibleRows = state.rows.filter(matchesFilter);
    state.selectedRowIds = checkbox.checked ? new Set(visibleRows.map((row) => row.uiId)) : new Set();
    renderRows();
  });

  elements.searchResultsBody.addEventListener('click', async (event) => {
    const button = event.target.closest('[data-open-search-result]');
    if (!button) {
      return;
    }
    await openSearchResult(button.dataset.openSearchResult);
  });

  elements.setGeminiKeyButton.addEventListener('click', promptForGeminiKey);

  elements.pdfDropZone.addEventListener('click', () => {
    elements.pdfFileInput.click();
  });
  elements.pdfDropZone.addEventListener('dragenter', (event) => {
    event.preventDefault();
    elements.pdfDropZone.classList.add('is-dragover');
  });
  elements.pdfDropZone.addEventListener('dragover', (event) => {
    event.preventDefault();
    elements.pdfDropZone.classList.add('is-dragover');
  });
  elements.pdfDropZone.addEventListener('dragleave', () => {
    elements.pdfDropZone.classList.remove('is-dragover');
  });
  elements.pdfDropZone.addEventListener('drop', async (event) => {
    event.preventDefault();
    elements.pdfDropZone.classList.remove('is-dragover');
    const file = event.dataTransfer?.files?.[0];
    if (file) {
      await handlePdfImport(file);
    }
  });

  elements.pdfFileInput.addEventListener('change', async () => {
    const file = elements.pdfFileInput.files?.[0];
    if (file) {
      await handlePdfImport(file);
    }
    elements.pdfFileInput.value = '';
  });

  [
    elements.projectC2Input,
    elements.projectNameInput,
    elements.projectShortNameInput,
    elements.projectContactInput,
    elements.drawingNumberInput,
    elements.drawingStatusInput
  ].forEach((input) => {
    input.addEventListener('input', () => {
      syncProjectFromForm();
      syncDrawingFromForm();
      renderReport();
      scheduleAutoSave();
    });
  });

  elements.bulkSymbolsInput?.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      applyBulkSymbols();
    }
  });
}

function renderInitialState() {
  renderTableHead();
  state.rows = [createUiRow()];
  syncSavedSignature();
  renderSearchResults();
  renderAll();
}

async function bootstrap() {
  renderInitialState();
  bindEvents();
  setActiveMode(state.activeMode);
  setStatus('Firestore 版を初期化しました。');
  await refreshProjects();
}

bootstrap();
