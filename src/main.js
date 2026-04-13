import './style.css';
import { FIELD_DEFS, PROJECT_TEMPLATE, DRAWING_TEMPLATE, ROW_TEMPLATE } from './schema.js';
import { assignSymbolsToDrawing, listProjects, loadAssignmentHistory, loadProjectBundle, loadDrawingRows, loadProjectSymbols, saveDrawingBundle } from './store.js';
import { extractHandaiDataFromPdf } from './gemini.js';

const SAVE_ACTOR = 'system';
const SIDEBAR_STORAGE_KEY = 'inputto_sidebar_collapsed';
const MODE_STORAGE_KEY = 'inputto_active_mode';
const CONTACT_OPTIONS = ['高橋', '髙林', '小島', '佐野'];
const DISALLOWED_CONTACTS = new Set(['鈴木', '鈴木様']);
const DEFAULT_MODE = 'register';
const ASSIGNMENT_ALL_GROUP = '__all__';
const ASSIGNMENT_STEPS = ['project', 'load', 'boxes', 'symbols', 'confirm'];

const state = {
  env: 'production',
  activeMode: normalizeMode(localStorage.getItem(MODE_STORAGE_KEY)),
  sidebarCollapsed: localStorage.getItem(SIDEBAR_STORAGE_KEY) === 'true',
  loading: false,
  saving: false,
  analyzingPdf: false,
  projects: [],
  project: { ...PROJECT_TEMPLATE },
  drawings: [],
  selectedDrawingId: '',
  drawing: { ...DRAWING_TEMPLATE },
  rows: [],
  selectedRowIds: new Set(),
  filterText: '',
  assignment: {
    rows: [],
    selectedDocIds: new Set(),
    targetDrawingNumber: '',
    filterText: '',
    collapsedGroupKeys: new Set(),
    activeGroupKeys: new Set(),
    boxes: [],
    activeBoxId: '',
    step: 'project',
    history: []
  },
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
  assignmentProjectSelect: document.getElementById('assignmentProjectSelect'),
  reportProjectSelect: document.getElementById('reportProjectSelect'),
  refreshProjectsButton: document.getElementById('refreshProjectsButton'),
  assignmentRefreshProjectsButton: document.getElementById('assignmentRefreshProjectsButton'),
  newProjectButton: document.getElementById('newProjectButton'),
  registerButton: document.getElementById('registerButton'),
  projectC2Input: document.getElementById('projectC2Input'),
  projectNameInput: document.getElementById('projectNameInput'),
  projectShortNameInput: document.getElementById('projectShortNameInput'),
  projectContactInput: document.getElementById('projectContactInput'),
  drawingNumberInput: document.getElementById('drawingNumberInput'),
  drawingStatusInput: document.getElementById('drawingStatusInput'),
  drawingTabs: document.getElementById('drawingTabs'),
  assignmentDrawingTabs: document.getElementById('assignmentDrawingTabs'),
  assignmentStage: document.getElementById('assignmentStage'),
  assignmentStepButtons: Array.from(document.querySelectorAll('[data-assignment-step]')),
  assignmentStepPanels: Array.from(document.querySelectorAll('[data-assignment-step-panel]')),
  assignmentSymbolsList: document.getElementById('assignmentSymbolsList'),
  assignmentHistoryList: document.getElementById('assignmentHistoryList'),
  assignmentConfirmSummary: document.getElementById('assignmentConfirmSummary'),
  assignmentFilterInput: document.getElementById('assignmentFilterInput'),
  assignmentTargetDrawingInput: document.getElementById('assignmentTargetDrawingInput'),
  assignmentBoxCountSelect: document.getElementById('assignmentBoxCountSelect'),
  assignmentSelectedCount: document.getElementById('assignmentSelectedCount'),
  assignmentBoxesList: document.getElementById('assignmentBoxesList'),
  assignmentBoxesLists: Array.from(document.querySelectorAll('[data-assignment-boxes-list]')),
  loadAssignmentButton: document.getElementById('loadAssignmentButton'),
  addAssignmentBoxButton: document.getElementById('addAssignmentBoxButton'),
  addSelectionToBoxButton: document.getElementById('addSelectionToBoxButton'),
  addVisibleToBoxButton: document.getElementById('addVisibleToBoxButton'),
  applyAssignmentBoxesButton: document.getElementById('applyAssignmentBoxesButton'),
  clearAssignmentBoxesButton: document.getElementById('clearAssignmentBoxesButton'),
  selectAllAssignmentButton: document.getElementById('selectAllAssignmentButton'),
  clearAssignmentSelectionButton: document.getElementById('clearAssignmentSelectionButton'),
  bulkSymbolsInput: document.getElementById('bulkSymbolsInput'),
  applyBulkSymbolsButton: document.getElementById('applyBulkSymbolsButton'),
  clearBulkSymbolsButton: document.getElementById('clearBulkSymbolsButton'),
  pickPdfButton: document.getElementById('pickPdfButton'),
  pdfFileInput: document.getElementById('pdfFileInput'),
  appDropOverlay: document.getElementById('appDropOverlay'),
  pdfBusyOverlay: document.getElementById('pdfBusyOverlay'),
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
  reportRegisterButton: document.getElementById('reportRegisterButton'),
  printButton: document.getElementById('printButton'),
  toastHost: document.getElementById('toastHost')
};

let autoSaveTimerId = null;
let lastSavedSignature = '';
let appDragDepth = 0;

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

function sanitizeContact(value) {
  const contact = String(value || '').trim();
  return DISALLOWED_CONTACTS.has(contact) ? '' : contact;
}

function normalizeMode(mode) {
  if (mode === 'editor') {
    return DEFAULT_MODE;
  }
  return ['register', 'assignment-edit', 'search', 'report'].includes(mode) ? mode : DEFAULT_MODE;
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

function canEnterAssignmentStep(step) {
  if (step === 'project') {
    return true;
  }
  if (!state.project.c2) {
    return false;
  }
  if (step === 'load') {
    return true;
  }
  if (step === 'boxes' || step === 'symbols') {
    return Boolean(state.assignment.rows.length);
  }
  if (step === 'confirm') {
    return state.assignment.boxes.some((box) => box.rowKeys?.size);
  }
  return false;
}

function setAssignmentStep(step, options = {}) {
  const nextStep = ASSIGNMENT_STEPS.includes(step) ? step : 'project';
  if (!canEnterAssignmentStep(nextStep)) {
    if (!options.silent) {
      if (!state.project.c2) {
        showToast('先に工事を選んでください。', 'error');
      } else if (nextStep === 'boxes' || nextStep === 'symbols') {
        showToast('先に登録済みを読込してください。', 'error');
      } else if (nextStep === 'confirm') {
        showToast('新手配書番号に符号を入れてください。', 'error');
      }
    }
    return false;
  }
  const stepChanged = state.assignment.step !== nextStep;
  state.assignment.step = nextStep;
  if (stepChanged) {
    renderAssignmentBoxes();
  }
  renderAssignmentWizard();
  if (stepChanged && elements.assignmentStage) {
    elements.assignmentStage.scrollTo({ top: 0, behavior: 'smooth' });
  }
  return true;
}

async function navigateAssignmentStep(step) {
  const changed = setAssignmentStep(step);
  if (!changed) {
    return;
  }
  if (step === 'load' && state.project.c2 && !state.assignment.rows.length) {
    await loadAssignmentSymbols();
  }
}

function getAiLogicSetupUrl(message) {
  const matched = String(message || '').match(/https:\/\/console\.firebase\.google\.com\/project\/[^\s]+\/ailogic\/?/);
  return matched ? matched[0] : '';
}

function showToast(message, kind = 'info') {
  if (!elements.toastHost || !message) {
    return;
  }
  const toast = document.createElement('div');
  const strongSuccess = kind === 'success' && ['登録完了', 'PDFから入力'].some((text) => String(message).includes(text));
  toast.className = `toast${kind === 'error' ? ' is-error' : kind === 'success' ? ' is-success' : ''}${strongSuccess ? ' is-strong-success' : ''}`;
  const setupUrl = getAiLogicSetupUrl(message);

  if (setupUrl) {
    const text = document.createElement('span');
    text.textContent = 'Firebase AI Logic が未設定です。';
    const link = document.createElement('a');
    link.href = setupUrl;
    link.target = '_blank';
    link.rel = 'noreferrer';
    link.textContent = '設定を開く';
    toast.append(text, link);
  } else {
    toast.textContent = message;
  }

  elements.toastHost.appendChild(toast);
  window.setTimeout(() => {
    toast.remove();
  }, setupUrl ? 15000 : strongSuccess ? 5200 : 3200);
}

function setActiveMode(mode) {
  state.activeMode = normalizeMode(mode);
  localStorage.setItem(MODE_STORAGE_KEY, state.activeMode);
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
    contact: sanitizeContact(elements.projectContactInput.value)
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

function renderContactOptions(value = '') {
  if (!elements.projectContactInput) {
    return;
  }

  const currentValue = sanitizeContact(value);
  const options = [...CONTACT_OPTIONS];
  if (currentValue && !options.includes(currentValue)) {
    options.unshift(currentValue);
  }

  elements.projectContactInput.innerHTML = [
    '<option value="">担当を選択</option>',
    ...options.map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option)}</option>`)
  ].join('');
}

function updateFormInputs() {
  const contact = sanitizeContact(state.project.contact || '');
  state.project.contact = contact;
  renderContactOptions(contact);
  elements.projectC2Input.value = state.project.c2 || '';
  elements.projectNameInput.value = state.project.projectName || '';
  elements.projectShortNameInput.value = state.project.shortName || '';
  elements.projectContactInput.value = contact;
  elements.drawingNumberInput.value = state.drawing.drawingNumber || '';
  elements.drawingStatusInput.value = state.drawing.drawingStatus || '';
  elements.filterInput.value = state.filterText || '';
  if (elements.reportProjectSelect) {
    elements.reportProjectSelect.value = state.project.c2 || '';
  }
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
  if (elements.assignmentProjectSelect) {
    elements.assignmentProjectSelect.innerHTML = buildProjectOptions(state.project.c2 || '');
  }
  if (elements.reportProjectSelect) {
    elements.reportProjectSelect.innerHTML = buildProjectOptions(state.project.c2 || '');
  }
  elements.searchProjectSelect.innerHTML = buildProjectOptions(state.search.projectC2 || '');
}

function renderDrawingTabs() {
  let reportHtml = '';
  if (!state.project.c2) {
    reportHtml = '<p class="empty-text">工事を選ぶと手配書タブが出ます。</p>';
  } else if (!state.drawings.length) {
    reportHtml = '<p class="empty-text">まだ登録済みの手配書がありません。</p>';
  } else {
    reportHtml = state.drawings
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

  if (elements.drawingTabs) {
    elements.drawingTabs.innerHTML = reportHtml;
  }

  if (!elements.assignmentDrawingTabs) {
    return;
  }
  if (!state.project.c2) {
    elements.assignmentDrawingTabs.innerHTML = '<p class="empty-text">工事を選ぶと手配書タブが出ます。</p>';
    return;
  }
  if (!state.assignment.rows.length) {
    elements.assignmentDrawingTabs.innerHTML = '<p class="empty-text">登録済み符号を読み込んでいます。</p>';
    return;
  }

  const groups = getAssignmentGroups();
  const activeGroupKeys = getActiveAssignmentGroupKeys(groups);
  const allActive = groups.length > 0 && groups.every((group) => activeGroupKeys.has(group.key)) ? ' is-active' : '';
  const allPreview = state.assignment.rows
    .slice(0, 8)
    .map((row) => row.symbol)
    .filter(Boolean)
    .join('、');
  const tabs = [
    `
      <button type="button" class="drawing-chip assignment-tab-chip${allActive}" data-assignment-tab="${ASSIGNMENT_ALL_GROUP}">
        <span>全て</span>
        <small>${state.assignment.rows.length}件${allPreview ? ` / ${escapeHtml(allPreview)}` : ''}</small>
      </button>
    `,
    ...groups.map((group) => {
      const active = activeGroupKeys.has(group.key) ? ' is-active' : '';
      const preview = group.rows
        .slice(0, 8)
        .map((row) => row.symbol)
        .filter(Boolean)
        .join('、');
      return `
        <button type="button" class="drawing-chip assignment-tab-chip${active}" data-assignment-tab="${escapeHtml(group.key)}">
          <span>${escapeHtml(group.drawingNumber || '-')}</span>
          <small>${group.rows.length}件${preview ? ` / ${escapeHtml(preview)}` : ''}</small>
        </button>
      `;
    })
  ];
  elements.assignmentDrawingTabs.innerHTML = tabs.join('');
}

function getAssignmentGroups() {
  const groups = new Map();
  state.assignment.rows.forEach((row) => {
    const key = row.drawingId || '__unassigned';
    const group = groups.get(key) || {
      key,
      drawingNumber: row.drawingNumber || '未割当',
      drawingId: row.drawingId || '',
      rows: []
    };
    group.rows.push(row);
    groups.set(key, group);
  });
  return Array.from(groups.values()).sort((left, right) =>
    String(left.drawingNumber || '').localeCompare(String(right.drawingNumber || ''), 'ja', { numeric: true })
  );
}

function getActiveAssignmentGroupKeys(groups = null) {
  const sourceGroups = groups || getAssignmentGroups();
  const validKeys = new Set(sourceGroups.map((group) => group.key));
  const nextKeys = new Set(Array.from(state.assignment.activeGroupKeys || []).filter((key) => validKeys.has(key)));

  if (nextKeys.size) {
    state.assignment.activeGroupKeys = nextKeys;
    return nextKeys;
  }

  const firstDrawingId = state.drawings[0]?.id || '';
  const firstDrawingExists = sourceGroups.some((group) => group.key === firstDrawingId);
  const fallbackKey = firstDrawingExists ? firstDrawingId : sourceGroups[0]?.key || firstDrawingId || '';
  state.assignment.activeGroupKeys = fallbackKey ? new Set([fallbackKey]) : new Set();
  return state.assignment.activeGroupKeys;
}

function assignmentRowMatches(row) {
  const term = normalizeText(state.assignment.filterText);
  if (!term) {
    return true;
  }
  return normalizeText([row.drawingNumber, row.symbol, row.name, row.floor, row.insideOutside].join(' ')).includes(term);
}

function getVisibleAssignmentRows() {
  const groups = getAssignmentGroups();
  const activeGroupKeys = getActiveAssignmentGroupKeys(groups);
  return state.assignment.rows.filter((row) => {
    const key = row.drawingId || '__unassigned';
    return activeGroupKeys.has(key) && assignmentRowMatches(row);
  });
}

function getAssignmentSelectionKey(row) {
  return [
    row.docId || '',
    row.drawingId || '',
    row.drawingNumber || '',
    row.symbolN || normalizeText(row.symbol),
    row.uiId || ''
  ].join('__');
}

function getSelectedAssignmentDocIds() {
  const selectedKeys = state.assignment.selectedDocIds;
  return getAssignmentDocIdsForKeys(selectedKeys);
}

function getAssignmentRowsForKeys(rowKeys) {
  const keySet = rowKeys instanceof Set ? rowKeys : new Set(rowKeys);
  return state.assignment.rows.filter((row) => keySet.has(getAssignmentSelectionKey(row)));
}

function getAssignmentDocIdsForKeys(rowKeys) {
  return Array.from(new Set(getAssignmentRowsForKeys(rowKeys).map((row) => row.docId).filter(Boolean)));
}

function getSelectedAssignmentGroups() {
  const groups = getAssignmentGroups();
  const activeKeys = getActiveAssignmentGroupKeys(groups);
  return groups.filter((group) => activeKeys.has(group.key));
}

function getDefaultAssignmentTargetBase() {
  const selectedGroups = getSelectedAssignmentGroups();
  const numbers = selectedGroups.map((group) => group.drawingNumber).filter(Boolean);
  if (!numbers.length) {
    return '';
  }
  const bases = numbers.map((number) => String(number).split('-')[0]).filter(Boolean);
  if (bases.length && bases.every((base) => base === bases[0])) {
    return bases[0];
  }
  return numbers[0];
}

function buildAssignmentTargetNumbers(base, count) {
  const normalizedBase = String(base || '').trim();
  const total = Math.max(1, Math.min(10, Number(count) || 1));
  if (!normalizedBase) {
    return [];
  }
  if (total === 1) {
    return [normalizedBase];
  }
  return Array.from({ length: total }, (_, index) => `${normalizedBase}-${index + 1}`);
}

function createAssignmentBox(targetDrawingNumber) {
  const normalizedTarget = String(targetDrawingNumber || '').trim();
  if (!normalizedTarget) {
    return null;
  }
  const existing = state.assignment.boxes.find((box) => normalizeText(box.targetDrawingNumber) === normalizeText(normalizedTarget));
  if (existing) {
    state.assignment.activeBoxId = existing.id;
    return existing;
  }
  const box = {
    id: crypto.randomUUID(),
    targetDrawingNumber: normalizedTarget,
    rowKeys: new Set()
  };
  state.assignment.boxes.push(box);
  state.assignment.activeBoxId = box.id;
  return box;
}

function createAssignmentBoxesFromControls() {
  const base = state.assignment.targetDrawingNumber || getDefaultAssignmentTargetBase();
  const targets = buildAssignmentTargetNumbers(base, elements.assignmentBoxCountSelect?.value || 1);
  if (!targets.length) {
    showToast('新手配書番号を入力してください。', 'error');
    elements.assignmentTargetDrawingInput?.focus();
    return false;
  }
  targets.forEach(createAssignmentBox);
  const firstBox = state.assignment.boxes.find((box) => normalizeText(box.targetDrawingNumber) === normalizeText(targets[0]));
  if (firstBox) {
    state.assignment.activeBoxId = firstBox.id;
    state.assignment.targetDrawingNumber = firstBox.targetDrawingNumber;
  }
  renderAssignmentBoxes();
  renderAssignmentWizard();
  return true;
}

function getActiveAssignmentBox() {
  return state.assignment.boxes.find((box) => box.id === state.assignment.activeBoxId) || state.assignment.boxes[0] || null;
}

function addRowsToSpecificAssignmentBox(rowKeys, box) {
  if (!box) {
    showToast('入れ先の新手配書番号を選んでください。', 'error');
    return false;
  }
  const keys = Array.from(rowKeys || []).filter(Boolean);
  if (!keys.length) {
    showToast('入れる符号を選んでください。', 'error');
    return false;
  }

  state.assignment.boxes.forEach((item) => {
    keys.forEach((key) => item.rowKeys.delete(key));
  });
  keys.forEach((key) => box.rowKeys.add(key));
  state.assignment.activeBoxId = box.id;
  state.assignment.targetDrawingNumber = box.targetDrawingNumber;
  state.assignment.selectedDocIds = new Set();
  renderAssignmentBoxes();
  renderAssignmentList();
  renderAssignmentWizard();
  return true;
}

function addRowsToAssignmentBox(rowKeys, targetDrawingNumber = '') {
  const normalizedTarget = String(targetDrawingNumber || '').trim();
  const box = normalizedTarget ? createAssignmentBox(normalizedTarget) : getActiveAssignmentBox();
  if (!box && !state.assignment.boxes.length) {
    showToast('先に新手配書番号を作成してください。', 'error');
    elements.assignmentTargetDrawingInput?.focus();
    return false;
  }
  return addRowsToSpecificAssignmentBox(rowKeys, box);
}

function getAssignmentBoxForRowKey(rowKey) {
  return state.assignment.boxes.find((box) => box.rowKeys?.has(rowKey)) || null;
}

function removeAssignmentBox(boxId) {
  state.assignment.boxes = state.assignment.boxes.filter((box) => box.id !== boxId);
  if (state.assignment.activeBoxId === boxId) {
    state.assignment.activeBoxId = state.assignment.boxes[0]?.id || '';
  }
  renderAssignmentBoxes();
  renderAssignmentList();
  renderAssignmentWizard();
}

function renderAssignmentList() {
  if (!elements.assignmentSymbolsList) {
    return;
  }

  if (elements.assignmentFilterInput) {
    elements.assignmentFilterInput.value = state.assignment.filterText || '';
  }
  if (elements.assignmentTargetDrawingInput) {
    elements.assignmentTargetDrawingInput.value = state.assignment.targetDrawingNumber || '';
  }
  if (elements.assignmentSelectedCount) {
    const visibleCount = getVisibleAssignmentRows().length;
    elements.assignmentSelectedCount.textContent = `${state.assignment.selectedDocIds.size}件選択 / ${visibleCount}件表示`;
  }

  if (!state.project.c2) {
    elements.assignmentSymbolsList.innerHTML = '<p class="empty-text">工事を選ぶと登録済み符号を読み込めます。</p>';
    return;
  }
  if (!state.assignment.rows.length) {
    elements.assignmentSymbolsList.innerHTML = '<p class="empty-text">登録済みを読込すると、符号をチェックして手配書Noを変更できます。</p>';
    return;
  }

  const filtering = Boolean(normalizeText(state.assignment.filterText));
  const allGroups = getAssignmentGroups();
  const activeGroupKeys = getActiveAssignmentGroupKeys(allGroups);
  renderDrawingTabs();
  const groups = allGroups
    .filter((group) => activeGroupKeys.has(group.key))
    .map((group) => ({
      ...group,
      totalRows: group.rows.length,
      rows: group.rows.filter(assignmentRowMatches)
    }))
    .filter((group) => !filtering || group.rows.length);

  if (!groups.length) {
    elements.assignmentSymbolsList.innerHTML = '<p class="empty-text">一致する手配書No・符号がありません。</p>';
    return;
  }

  elements.assignmentSymbolsList.innerHTML = groups
    .map((group) => {
      const collapsed = !filtering && state.assignment.collapsedGroupKeys.has(group.key);
      const rows = group.rows
        .map((row) => {
          const rowKey = getAssignmentSelectionKey(row);
          const checked = state.assignment.selectedDocIds.has(rowKey) ? ' checked' : '';
          const pendingBox = getAssignmentBoxForRowKey(rowKey);
          const detail = [row.name, row.floor, row.insideOutside].filter(Boolean).join(' / ');
          return `
            <label class="assignment-row${pendingBox ? ' is-boxed' : ''}" draggable="true" data-assignment-drag-row="${escapeHtml(rowKey)}">
              <input type="checkbox" data-assignment-symbol="${escapeHtml(rowKey)}"${checked}>
              <span class="assignment-symbol">${escapeHtml(row.symbol || '-')}</span>
              <small>${escapeHtml(detail)}${pendingBox ? `<b>新番号 ${escapeHtml(pendingBox.targetDrawingNumber)}</b>` : ''}</small>
            </label>
          `;
        })
        .join('');

      return `
        <article class="assignment-group${collapsed ? ' is-collapsed' : ''}">
          <div class="assignment-group-head">
            <div class="assignment-group-title">
              <strong>${escapeHtml(group.drawingNumber)}</strong>
              <small>${filtering ? `${group.rows.length}/${group.totalRows}件` : `${group.rows.length}件`}</small>
            </div>
            <button
              type="button"
              class="assignment-toggle"
              data-assignment-toggle="${escapeHtml(group.key)}"
              aria-expanded="${collapsed ? 'false' : 'true'}">
              <span class="material-symbols-outlined">${collapsed ? 'expand_more' : 'expand_less'}</span>
              <span>${collapsed ? '表示' : '非表示'}</span>
            </button>
          </div>
          <div class="assignment-group-rows"${collapsed ? ' hidden' : ''}>${rows}</div>
        </article>
      `;
    })
    .join('');
}

function formatHistoryDate(value) {
  if (!value) {
    return '';
  }
  const date = value.seconds ? new Date(value.seconds * 1000) : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }
  return date.toLocaleString('ja-JP', {
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  });
}

function renderAssignmentHistory() {
  if (!elements.assignmentHistoryList) {
    return;
  }
  const history = state.assignment.history || [];
  if (!history.length) {
    elements.assignmentHistoryList.innerHTML = '<p class="empty-text">分割・統合するとここに履歴が残ります。</p>';
    return;
  }

  elements.assignmentHistoryList.innerHTML = history
    .slice(0, 10)
    .map((item) => {
      const sources = (item.sourceDrawings || [])
        .map((source) => `${source.drawingNumber || '-'} ${source.count || 0}件`)
        .join(' + ');
      const target = item.targetDrawing?.drawingNumber || '-';
      const preview = (item.symbolPreview || []).filter(Boolean).join(', ');
      const action = (item.sourceDrawings || []).length > 1 ? '統合' : '分割/移動';
      return `
        <article class="assignment-history-item">
          <strong>${escapeHtml(action)}: ${escapeHtml(sources)} → ${escapeHtml(target)}</strong>
          <span>${escapeHtml(String(item.symbolCount || 0))}件${preview ? ` / ${escapeHtml(preview)}` : ''}</span>
          <small>${escapeHtml(formatHistoryDate(item.createdAt))}</small>
        </article>
      `;
    })
    .join('');
}

function renderAssignmentBoxes() {
  const lists = elements.assignmentBoxesLists?.length ? elements.assignmentBoxesLists : [elements.assignmentBoxesList].filter(Boolean);
  if (!lists.length) {
    return;
  }
  if (elements.assignmentTargetDrawingInput && !state.assignment.targetDrawingNumber) {
    elements.assignmentTargetDrawingInput.value = getDefaultAssignmentTargetBase();
  }
  if (!state.assignment.boxes.length) {
    lists.forEach((list) => {
      list.innerHTML = '<p class="empty-text">作成数を選ぶと、新手配書番号がここに表示されます。</p>';
    });
    return;
  }

  const html = state.assignment.boxes
    .map((box) => {
      const rows = getAssignmentRowsForKeys(box.rowKeys);
      const active = box.id === state.assignment.activeBoxId ? ' is-active' : '';
      const preview = rows
        .slice(0, 6)
        .map((row) => row.symbol)
        .filter(Boolean)
        .join(', ');
      return `
        <article class="assignment-box${active}" data-assignment-box="${escapeHtml(box.id)}" data-assignment-box-drop="${escapeHtml(box.id)}">
          <div class="assignment-box-main" data-assignment-box-select="${escapeHtml(box.id)}">
            <label>
              <span>新手配書番号</span>
              <input type="text" value="${escapeHtml(box.targetDrawingNumber)}" data-assignment-box-name="${escapeHtml(box.id)}">
            </label>
            <span>${rows.length}件${preview ? ` / ${escapeHtml(preview)}` : ''}</span>
          </div>
          <button type="button" class="assignment-box-delete" data-assignment-box-delete="${escapeHtml(box.id)}">
            <span class="material-symbols-outlined">close</span>
          </button>
        </article>
      `;
    })
    .join('');
  lists.forEach((list) => {
    list.innerHTML = html;
  });
}

function renderAssignmentConfirmSummary() {
  if (!elements.assignmentConfirmSummary) {
    return;
  }
  const boxes = state.assignment.boxes
    .map((box) => ({
      ...box,
      rows: getAssignmentRowsForKeys(box.rowKeys || [])
    }))
    .filter((box) => box.rows.length);

  if (!boxes.length) {
    elements.assignmentConfirmSummary.innerHTML = '<p class="empty-text">新手配書番号に符号を入れると、ここで登録内容を確認できます。</p>';
    return;
  }

  elements.assignmentConfirmSummary.innerHTML = boxes
    .map((box) => {
      const preview = box.rows
        .slice(0, 10)
        .map((row) => row.symbol || '-')
        .join(', ');
      return `
        <article class="assignment-confirm-card">
          <strong>${escapeHtml(box.targetDrawingNumber)}</strong>
          <span>${box.rows.length}件</span>
          <small>${escapeHtml(preview)}${box.rows.length > 10 ? ' ほか' : ''}</small>
        </article>
      `;
    })
    .join('');
}

function renderAssignmentWizard() {
  const requestedStep = ASSIGNMENT_STEPS.includes(state.assignment.step) ? state.assignment.step : 'project';
  const activeStep = canEnterAssignmentStep(requestedStep) ? requestedStep : state.project.c2 ? 'load' : 'project';
  state.assignment.step = activeStep;

  elements.assignmentStepButtons.forEach((button) => {
    const step = button.dataset.assignmentStep;
    const active = step === activeStep;
    const available = canEnterAssignmentStep(step);
    button.classList.toggle('is-active', active);
    button.classList.toggle('is-disabled', !available);
    button.setAttribute('aria-current', active ? 'step' : 'false');
    button.setAttribute('aria-disabled', available ? 'false' : 'true');
  });

  elements.assignmentStepPanels.forEach((panel) => {
    const active = panel.dataset.assignmentStepPanel === activeStep;
    panel.classList.toggle('is-active', active);
    panel.hidden = !active;
  });

  renderAssignmentConfirmSummary();
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
}

async function applyBulkSymbols() {
  const symbols = parseBulkSymbols(elements.bulkSymbolsInput?.value || '');
  applyBulkSymbolsToRows(symbols, { clearInput: true, notify: true });
}

function applyBulkSymbolsToRows(symbols, options = {}) {
  const { clearInput = false, notify = false } = options;
  syncProjectFromForm();
  syncDrawingFromForm();
  if (!symbols.length && !state.selectedDrawingId) {
    if (notify) {
      setStatus('追加する符号がありません。');
      showToast('追加する符号がありません。', 'error');
    }
    return false;
  }

  if (!state.project.c2 || !state.project.projectName) {
    if (notify) {
      showToast('先に工事を入れてください。', 'error');
    }
    return false;
  }
  if (!state.drawing.drawingNumber) {
    if (notify) {
      showToast('先に手配書Noを入れてください。', 'error');
    }
    return false;
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
  if (clearInput && elements.bulkSymbolsInput) {
    elements.bulkSymbolsInput.value = '';
  }
  if (notify) {
    setStatus(`${symbols.length}件の符号を登録準備しました。登録ボタンで保存できます。`);
    showToast(`${symbols.length}件の符号を登録準備しました。`, 'success');
  }
  scheduleAutoSave();
  return true;
}

function clearBulkSymbols() {
  if (elements.bulkSymbolsInput) {
    elements.bulkSymbolsInput.value = '';
    elements.bulkSymbolsInput.focus();
  }
  showToast('一括入力を空にしました。');
}

function isPdfFile(file) {
  if (!file) {
    return false;
  }
  return file.type === 'application/pdf' || /\.pdf$/i.test(String(file.name || ''));
}

function showAppDropOverlay() {
  if (!elements.appDropOverlay) {
    return;
  }
  elements.appDropOverlay.hidden = false;
  window.requestAnimationFrame(() => {
    elements.appDropOverlay?.classList.add('is-active');
  });
}

function hideAppDropOverlay() {
  if (!elements.appDropOverlay) {
    return;
  }
  elements.appDropOverlay.classList.remove('is-active');
  window.setTimeout(() => {
    if (!elements.appDropOverlay?.classList.contains('is-active')) {
      elements.appDropOverlay.hidden = true;
    }
  }, 180);
}

function resetAppDropState() {
  appDragDepth = 0;
  hideAppDropOverlay();
}

function setPdfAnalysisBusy(active) {
  state.analyzingPdf = active;
  document.body.classList.toggle('is-pdf-analyzing', active);

  if (!elements.pdfBusyOverlay) {
    return;
  }

  if (active) {
    elements.pdfBusyOverlay.hidden = false;
    elements.pdfBusyOverlay.setAttribute('aria-hidden', 'false');
    window.requestAnimationFrame(() => {
      if (state.analyzingPdf) {
        elements.pdfBusyOverlay?.classList.add('is-active');
      }
    });
    return;
  }

  elements.pdfBusyOverlay.classList.remove('is-active');
  elements.pdfBusyOverlay.setAttribute('aria-hidden', 'true');
  window.setTimeout(() => {
    if (!elements.pdfBusyOverlay?.classList.contains('is-active')) {
      elements.pdfBusyOverlay.hidden = true;
    }
  }, 180);
}

function isFileDragEvent(event) {
  const types = Array.from(event.dataTransfer?.types || []);
  return types.includes('Files');
}

function getPdfFromFileList(fileList) {
  return Array.from(fileList || []).find((file) => isPdfFile(file)) || null;
}

async function handlePdfImport(file) {
  if (!file) {
    return;
  }
  if (!isPdfFile(file)) {
    showToast('PDFファイルを選んでください。', 'error');
    return;
  }

  setBusy('loading', 'PDFを解析しています。');
  setPdfAnalysisBusy(true);
  document.body.classList.add('is-importing-pdf');
  elements.pickPdfButton?.setAttribute('disabled', 'disabled');
  try {
    const extracted = await extractHandaiDataFromPdf(file);

    const bundle = extracted.project.c2
      ? await loadProjectBundle(state.env, extracted.project.c2)
      : { project: { ...PROJECT_TEMPLATE }, drawings: [] };

    state.project = {
      c2: extracted.project.c2 || bundle.project.c2 || '',
      projectName: extracted.project.projectName || bundle.project.projectName || '',
      shortName: extracted.project.shortName || bundle.project.shortName || '',
      contact: sanitizeContact(extracted.project.contact || bundle.project.contact || '')
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
    showToast('PDFから入力しました。工事と手配書で確認して登録できます。', 'success');
  } catch (error) {
    console.error(error);
    showToast(`PDFの読込に失敗しました: ${error.message}`, 'error');
  } finally {
    elements.pickPdfButton?.removeAttribute('disabled');
    document.body.classList.remove('is-importing-pdf');
    setPdfAnalysisBusy(false);
    resetAppDropState();
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
  renderAssignmentList();
  renderAssignmentBoxes();
  renderAssignmentHistory();
  renderAssignmentWizard();
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

function resetAssignmentState() {
  state.assignment = {
    rows: [],
    selectedDocIds: new Set(),
    targetDrawingNumber: '',
    filterText: '',
    collapsedGroupKeys: new Set(),
    activeGroupKeys: new Set(),
    boxes: [],
    activeBoxId: '',
    step: 'project',
    history: []
  };
}

function clearProjectState() {
  clearAutoSaveTimer();
  state.project = { ...PROJECT_TEMPLATE };
  state.drawings = [];
  state.filterText = '';
  resetDrawingState();
  resetAssignmentState();
  syncSavedSignature();
  renderAll();
}

function clearDirectFormValues() {
  [
    elements.projectC2Input,
    elements.projectNameInput,
    elements.projectShortNameInput,
    elements.projectContactInput,
    elements.reportProjectSelect,
    elements.drawingNumberInput,
    elements.drawingStatusInput,
    elements.bulkSymbolsInput,
    elements.filterInput,
    elements.assignmentFilterInput,
    elements.assignmentTargetDrawingInput,
    elements.pdfFileInput
  ].forEach((input) => {
    if (input) {
      input.value = '';
    }
  });
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
      contact: sanitizeContact(project.contact || '')
    };
    state.drawings = drawings;
    if (!state.search.projectC2) {
      state.search.projectC2 = state.project.c2;
    }
    resetDrawingState();
    resetAssignmentState();
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

async function loadAssignmentSymbols(options = {}) {
  const { silent = false } = options;
  syncProjectFromForm();

  if (!state.project.c2) {
    showToast('先に工事を選んでください。', 'error');
    return;
  }

  if (!silent) {
    setBusy('loading', '登録済み符号を読み込んでいます。');
  }
  try {
    const [rows, history] = await Promise.all([
      loadProjectSymbols(state.env, state.project.c2),
      loadAssignmentHistory(state.env, state.project.c2)
    ]);
    state.assignment.rows = rows;
    state.assignment.history = history;
    getActiveAssignmentGroupKeys(getAssignmentGroups());
    state.assignment.selectedDocIds = new Set(
      Array.from(state.assignment.selectedDocIds).filter((rowKey) =>
        state.assignment.rows.some((row) => getAssignmentSelectionKey(row) === rowKey)
      )
    );
    state.assignment.boxes = state.assignment.boxes
      .map((box) => ({
        ...box,
        rowKeys: new Set(Array.from(box.rowKeys || []).filter((rowKey) => state.assignment.rows.some((row) => getAssignmentSelectionKey(row) === rowKey)))
      }));
    renderAssignmentList();
    renderAssignmentBoxes();
    renderAssignmentHistory();
    renderAssignmentWizard();
    if (!silent) {
      setStatus(`${state.assignment.rows.length}件の登録済み符号を読み込みました。`);
    }
  } catch (error) {
    console.error(error);
    const message = `登録済み符号の読込に失敗しました: ${error.message}`;
    setStatus(message);
    showToast(message, 'error');
  }
}

async function moveSelectedAssignmentSymbols() {
  syncProjectFromForm();
  syncDrawingFromForm();
  const targetDrawingNumber = String(state.assignment.targetDrawingNumber || '').trim();
  const selectedIds = getSelectedAssignmentDocIds();

  if (!state.project.c2 || !state.project.projectName) {
    showToast('先に工事を選んでください。', 'error');
    return;
  }
  if (!selectedIds.length) {
    showToast('移動する符号をチェックしてください。', 'error');
    return;
  }
  if (!targetDrawingNumber) {
    showToast('移動先の手配書Noを入力してください。', 'error');
    elements.assignmentTargetDrawingInput?.focus();
    return;
  }

  setBusy('saving', '手配書Noを変更しています。');
  try {
    const response = await assignSymbolsToDrawing({
      env: state.env,
      project: state.project,
      drawing: {
        drawingNumber: targetDrawingNumber,
        drawingStatus: state.drawing.drawingNumber === targetDrawingNumber ? state.drawing.drawingStatus : ''
      },
      symbolIds: selectedIds,
      operator: SAVE_ACTOR
    });

    state.project = {
      ...state.project,
      ...response.project
    };
    state.drawings = response.drawings || state.drawings;
    state.selectedDrawingId = response.drawing.id;
    state.drawing = {
      id: response.drawing.id,
      drawingNumber: response.drawing.drawingNumber || '',
      drawingStatus: response.drawing.drawingStatus || ''
    };
    state.rows = Array.isArray(response.rows) && response.rows.length ? response.rows : [createUiRow()];
    state.selectedRowIds = new Set();
    state.assignment.selectedDocIds = new Set();
    state.assignment.targetDrawingNumber = '';
    state.assignment.activeGroupKeys = response.drawing.id ? new Set([response.drawing.id]) : state.assignment.activeGroupKeys;
    state.assignment.history = response.history || state.assignment.history;
    syncBulkSymbolsInput(state.rows);
    renderAll();
    syncSavedSignature();
    await refreshProjects();
    await loadAssignmentSymbols({ silent: true });
    const message = `${selectedIds.length}件を手配書No ${targetDrawingNumber} に移動しました。`;
    setStatus(message);
    showToast(message, 'success');
  } catch (error) {
    console.error(error);
    const message = `手配書Noの変更に失敗しました: ${error.message}`;
    setStatus(message);
    showToast(message, 'error');
  }
}

async function applyAssignmentBoxes() {
  syncProjectFromForm();
  const boxes = state.assignment.boxes
    .map((box) => ({
      ...box,
      symbolIds: getAssignmentDocIdsForKeys(box.rowKeys)
    }))
    .filter((box) => box.targetDrawingNumber && box.symbolIds.length);

  if (!state.project.c2 || !state.project.projectName) {
    showToast('先に工事を選んでください。', 'error');
    return;
  }
  if (!boxes.length) {
    showToast('登録する新手配書番号がありません。', 'error');
    return;
  }

  setBusy('saving', '新手配書番号の内容を登録しています。');
  try {
    let latestResponse = null;
    for (const box of boxes) {
      latestResponse = await assignSymbolsToDrawing({
        env: state.env,
        project: state.project,
        drawing: {
          drawingNumber: box.targetDrawingNumber,
          drawingStatus: ''
        },
        symbolIds: box.symbolIds,
        operator: SAVE_ACTOR
      });
      state.project = {
        ...state.project,
        ...latestResponse.project
      };
    }

    if (latestResponse) {
      state.drawings = latestResponse.drawings || state.drawings;
      state.selectedDrawingId = latestResponse.drawing.id;
      state.drawing = {
        id: latestResponse.drawing.id,
        drawingNumber: latestResponse.drawing.drawingNumber || '',
        drawingStatus: latestResponse.drawing.drawingStatus || ''
      };
      state.assignment.history = latestResponse.history || state.assignment.history;
      state.assignment.activeGroupKeys = latestResponse.drawing.id ? new Set([latestResponse.drawing.id]) : state.assignment.activeGroupKeys;
    }

    state.assignment.boxes = [];
    state.assignment.activeBoxId = '';
    state.assignment.selectedDocIds = new Set();
    state.assignment.targetDrawingNumber = '';
    state.assignment.step = 'load';
    renderAll();
    syncSavedSignature();
    await refreshProjects();
    await loadAssignmentSymbols({ silent: true });
    const movedCount = boxes.reduce((sum, box) => sum + box.symbolIds.length, 0);
    const message = `${boxes.length}個の新手配書番号、${movedCount}件の符号を登録しました。`;
    setStatus(message);
    showToast(message, 'success');
  } catch (error) {
    console.error(error);
    const message = `新手配書番号の登録に失敗しました: ${error.message}`;
    setStatus(message);
    showToast(message, 'error');
  }
}

async function saveCurrent(options = {}) {
  const { auto = false, resetAfterSave = !auto } = options;
  syncProjectFromForm();
  syncDrawingFromForm();
  const bulkSymbols = parseBulkSymbols(elements.bulkSymbolsInput?.value || '');
  if (!auto && state.activeMode === 'register' && bulkSymbols.length) {
    applyBulkSymbolsToRows(bulkSymbols);
  }

  if (!canAutoSave()) {
    if (!auto) {
      const message = '登録できません。工事番号・工事名・手配書No・符号の未入力を確認してください。';
      setStatus(message);
      showToast(message, 'error');
    }
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
    const message = auto ? '自動保存しました。' : '登録完了しました。';
    if (resetAfterSave) {
      clearProjectState();
      elements.projectSelect.value = '';
      if (elements.assignmentProjectSelect) {
        elements.assignmentProjectSelect.value = '';
      }
      if (elements.reportProjectSelect) {
        elements.reportProjectSelect.value = '';
      }
      setActiveMode('register');
    }
    setStatus(message);
    showToast(message, 'success');
  } catch (error) {
    console.error(error);
    const message = `登録に失敗しました: ${error.message}`;
    setStatus(message);
    showToast(message, 'error');
  }
}

function addRow() {
  state.rows.push(createUiRow());
  renderRows();
  scheduleAutoSave();
}

function startNewProject() {
  clearProjectState();
  elements.projectSelect.value = '';
  if (elements.assignmentProjectSelect) {
    elements.assignmentProjectSelect.value = '';
  }
  if (elements.reportProjectSelect) {
    elements.reportProjectSelect.value = '';
  }
  state.search.projectC2 = '';
  state.search.cacheProjectC2 = '';
  state.search.cacheRows = [];
  state.search.results = [];
  state.search.resultCountText = '0件';
  state.search.hint = '検索結果はここに表示されます。';
  clearDirectFormValues();
  resetAppDropState();
  setPdfAnalysisBusy(false);
  renderAll();
  clearDirectFormValues();
  renderAssignmentList();
  setActiveMode('register');
  elements.projectC2Input.focus();
  setStatus('新規工事入力に切り替えました。');
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

  elements.assignmentStepButtons.forEach((button) => {
    button.addEventListener('click', () => {
      void navigateAssignmentStep(button.dataset.assignmentStep);
    });
  });

  document.querySelectorAll('[data-assignment-next]').forEach((button) => {
    button.addEventListener('click', () => {
      void navigateAssignmentStep(button.dataset.assignmentNext);
    });
  });

  document.querySelectorAll('[data-assignment-prev]').forEach((button) => {
    button.addEventListener('click', () => {
      setAssignmentStep(button.dataset.assignmentPrev);
    });
  });

  elements.projectSelect.addEventListener('change', async () => {
    await selectProject(elements.projectSelect.value);
  });
  elements.assignmentProjectSelect?.addEventListener('change', async () => {
    await selectProject(elements.assignmentProjectSelect.value);
    renderAssignmentList();
  });
  elements.reportProjectSelect?.addEventListener('change', async () => {
    await selectProject(elements.reportProjectSelect.value);
    setActiveMode('report');
  });

  elements.refreshProjectsButton.addEventListener('click', refreshProjects);
  elements.assignmentRefreshProjectsButton?.addEventListener('click', refreshProjects);
  elements.newProjectButton.addEventListener('click', () => {
    startNewProject();
  });
  elements.registerButton?.addEventListener('click', () => {
    void saveCurrent();
  });
  elements.reportRegisterButton?.addEventListener('click', () => {
    void saveCurrent({ resetAfterSave: false });
  });
  elements.loadAssignmentButton?.addEventListener('click', () => {
    void loadAssignmentSymbols();
  });
  elements.addAssignmentBoxButton?.addEventListener('click', () => {
    createAssignmentBoxesFromControls();
  });
  elements.addSelectionToBoxButton?.addEventListener('click', () => {
    addRowsToAssignmentBox(state.assignment.selectedDocIds);
  });
  elements.addVisibleToBoxButton?.addEventListener('click', () => {
    addRowsToAssignmentBox(getVisibleAssignmentRows().map(getAssignmentSelectionKey));
  });
  elements.applyAssignmentBoxesButton?.addEventListener('click', () => {
    void applyAssignmentBoxes();
  });
  elements.clearAssignmentBoxesButton?.addEventListener('click', () => {
    state.assignment.boxes = [];
    state.assignment.activeBoxId = '';
    renderAssignmentBoxes();
    renderAssignmentList();
    renderAssignmentWizard();
  });
  elements.selectAllAssignmentButton.addEventListener('click', () => {
    state.assignment.selectedDocIds = new Set(getVisibleAssignmentRows().map(getAssignmentSelectionKey));
    renderAssignmentList();
  });
  elements.clearAssignmentSelectionButton.addEventListener('click', () => {
    state.assignment.selectedDocIds = new Set();
    renderAssignmentList();
  });
  elements.assignmentFilterInput.addEventListener('input', () => {
    state.assignment.filterText = elements.assignmentFilterInput.value;
    renderAssignmentList();
  });
  elements.assignmentTargetDrawingInput.addEventListener('input', () => {
    state.assignment.targetDrawingNumber = elements.assignmentTargetDrawingInput.value;
  });
  elements.assignmentBoxesLists.forEach((boxList) => {
    boxList.addEventListener('click', (event) => {
      if (event.target.closest('[data-assignment-box-name]')) {
        return;
      }
      const selectButton = event.target.closest('[data-assignment-box-select]');
      if (selectButton) {
        const box = state.assignment.boxes.find((item) => item.id === selectButton.dataset.assignmentBoxSelect);
        if (box) {
          state.assignment.activeBoxId = box.id;
          state.assignment.targetDrawingNumber = box.targetDrawingNumber;
          if (elements.assignmentTargetDrawingInput) {
            elements.assignmentTargetDrawingInput.value = box.targetDrawingNumber;
          }
          renderAssignmentBoxes();
          renderAssignmentWizard();
        }
        return;
      }
      const deleteButton = event.target.closest('[data-assignment-box-delete]');
      if (deleteButton) {
        removeAssignmentBox(deleteButton.dataset.assignmentBoxDelete);
      }
    });
    boxList.addEventListener('input', (event) => {
      const input = event.target.closest('[data-assignment-box-name]');
      if (!input) {
        return;
      }
      const box = state.assignment.boxes.find((item) => item.id === input.dataset.assignmentBoxName);
      if (!box) {
        return;
      }
      box.targetDrawingNumber = input.value;
      state.assignment.activeBoxId = box.id;
      state.assignment.targetDrawingNumber = input.value;
      renderAssignmentWizard();
    });
    boxList.addEventListener('dragover', (event) => {
      const dropTarget = event.target.closest('[data-assignment-box-drop]');
      if (!dropTarget) {
        return;
      }
      event.preventDefault();
      dropTarget.classList.add('is-drop-target');
    });
    boxList.addEventListener('dragleave', (event) => {
      event.target.closest('[data-assignment-box-drop]')?.classList.remove('is-drop-target');
    });
    boxList.addEventListener('drop', (event) => {
      const dropTarget = event.target.closest('[data-assignment-box-drop]');
      if (!dropTarget) {
        return;
      }
      event.preventDefault();
      dropTarget.classList.remove('is-drop-target');
      const rowKey = event.dataTransfer?.getData('text/plain');
      const box = state.assignment.boxes.find((item) => item.id === dropTarget.dataset.assignmentBoxDrop);
      addRowsToSpecificAssignmentBox(rowKey ? [rowKey] : [], box);
    });
  });
  elements.assignmentSymbolsList.addEventListener('click', (event) => {
    const toggleButton = event.target.closest('[data-assignment-toggle]');
    if (!toggleButton) {
      return;
    }
    const key = toggleButton.dataset.assignmentToggle;
    if (state.assignment.collapsedGroupKeys.has(key)) {
      state.assignment.collapsedGroupKeys.delete(key);
    } else {
      state.assignment.collapsedGroupKeys.add(key);
    }
    renderAssignmentList();
  });
  elements.assignmentSymbolsList.addEventListener('dragstart', (event) => {
    const row = event.target.closest('[data-assignment-drag-row]');
    if (!row) {
      return;
    }
    event.dataTransfer?.setData('text/plain', row.dataset.assignmentDragRow || '');
    event.dataTransfer.effectAllowed = 'move';
  });
  elements.assignmentDrawingTabs?.addEventListener('click', (event) => {
    const tab = event.target.closest('[data-assignment-tab]');
    if (!tab) {
      return;
    }
    const key = tab.dataset.assignmentTab || '';
    if (key === ASSIGNMENT_ALL_GROUP) {
      state.assignment.activeGroupKeys = new Set(getAssignmentGroups().map((group) => group.key));
    } else if (state.assignment.activeGroupKeys.has(key)) {
      state.assignment.activeGroupKeys.delete(key);
      if (!state.assignment.activeGroupKeys.size) {
        state.assignment.activeGroupKeys.add(key);
      }
    } else {
      state.assignment.activeGroupKeys.add(key);
    }
    state.assignment.selectedDocIds = new Set();
    state.assignment.collapsedGroupKeys = new Set();
    renderDrawingTabs();
    renderAssignmentList();
  });
  elements.assignmentSymbolsList.addEventListener('change', (event) => {
    const symbolCheckbox = event.target.closest('[data-assignment-symbol]');
    if (symbolCheckbox) {
      const rowKey = symbolCheckbox.dataset.assignmentSymbol;
      if (symbolCheckbox.checked) {
        state.assignment.selectedDocIds.add(rowKey);
      } else {
        state.assignment.selectedDocIds.delete(rowKey);
      }
      renderAssignmentList();
      return;
    }

  });
  elements.applyBulkSymbolsButton?.addEventListener('click', applyBulkSymbols);
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

  window.addEventListener('beforeunload', (event) => {
    if (!state.analyzingPdf) {
      return undefined;
    }
    event.preventDefault();
    event.returnValue = '';
    return '';
  });

  const handleDrawingTabClick = async (event) => {
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
  };
  elements.drawingTabs.addEventListener('click', handleDrawingTabClick);

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

  elements.pickPdfButton.addEventListener('click', () => {
    elements.pdfFileInput.click();
  });

  document.addEventListener('dragenter', (event) => {
    if (!isFileDragEvent(event)) {
      return;
    }
    event.preventDefault();
    appDragDepth += 1;
    showAppDropOverlay();
  });

  document.addEventListener('dragover', (event) => {
    if (!isFileDragEvent(event)) {
      return;
    }
    event.preventDefault();
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = 'copy';
    }
    showAppDropOverlay();
  });

  document.addEventListener('dragleave', (event) => {
    if (!isFileDragEvent(event)) {
      return;
    }
    appDragDepth = Math.max(0, appDragDepth - 1);
    if (appDragDepth === 0) {
      hideAppDropOverlay();
    }
  });

  document.addEventListener('drop', async (event) => {
    if (!isFileDragEvent(event)) {
      return;
    }
    event.preventDefault();
    const file = getPdfFromFileList(event.dataTransfer?.files);
    resetAppDropState();
    if (file) {
      await handlePdfImport(file);
      return;
    }
    showToast('PDFファイルをドロップしてください。', 'error');
  });

  window.addEventListener('blur', resetAppDropState);

  elements.pdfFileInput.addEventListener('change', async () => {
    const file = getPdfFromFileList(elements.pdfFileInput.files);
    if (file) {
      await handlePdfImport(file);
    }
    elements.pdfFileInput.value = '';
  });

  const formInputs = [
    elements.projectC2Input,
    elements.projectNameInput,
    elements.projectShortNameInput,
    elements.projectContactInput,
    elements.drawingNumberInput,
    elements.drawingStatusInput
  ];
  formInputs.forEach((input) => {
    const handleFormChange = () => {
      syncProjectFromForm();
      syncDrawingFromForm();
      renderReport();
      scheduleAutoSave();
    };
    input.addEventListener('input', handleFormChange);
    if (input.tagName === 'SELECT') {
      input.addEventListener('change', handleFormChange);
    }
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
