import './style.css';
import { FIELD_DEFS, PROJECT_TEMPLATE, DRAWING_TEMPLATE, ROW_TEMPLATE } from './schema.js';
import { listProjects, loadProjectBundle, loadDrawingRows, saveDrawingBundle } from './store.js';

const state = {
  env: 'production',
  operator: localStorage.getItem('inputto_operator') || '',
  loading: false,
  saving: false,
  projects: [],
  project: { ...PROJECT_TEMPLATE },
  drawings: [],
  selectedDrawingId: '',
  drawing: { ...DRAWING_TEMPLATE },
  rows: [],
  selectedRowIds: new Set(),
  filterText: ''
};

const elements = {
  operatorNameInput: document.getElementById('operatorNameInput'),
  busyState: document.getElementById('busyState'),
  statusText: document.getElementById('statusText'),
  projectSelect: document.getElementById('projectSelect'),
  refreshProjectsButton: document.getElementById('refreshProjectsButton'),
  newProjectButton: document.getElementById('newProjectButton'),
  saveButton: document.getElementById('saveButton'),
  projectC2Input: document.getElementById('projectC2Input'),
  projectNameInput: document.getElementById('projectNameInput'),
  projectShortNameInput: document.getElementById('projectShortNameInput'),
  projectContactInput: document.getElementById('projectContactInput'),
  drawingNumberInput: document.getElementById('drawingNumberInput'),
  drawingStatusInput: document.getElementById('drawingStatusInput'),
  loadDrawingButton: document.getElementById('loadDrawingButton'),
  drawingTabs: document.getElementById('drawingTabs'),
  filterInput: document.getElementById('filterInput'),
  addRowButton: document.getElementById('addRowButton'),
  duplicateRowButton: document.getElementById('duplicateRowButton'),
  deleteRowsButton: document.getElementById('deleteRowsButton'),
  tableHead: document.getElementById('tableHead'),
  tableBody: document.getElementById('tableBody'),
  summaryProject: document.getElementById('summaryProject'),
  summaryDrawing: document.getElementById('summaryDrawing'),
  summaryRows: document.getElementById('summaryRows'),
  summaryLabels: document.getElementById('summaryLabels'),
  labelPages: document.getElementById('labelPages'),
  printButton: document.getElementById('printButton')
};

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

function setBusy(mode, message) {
  state.loading = mode === 'loading';
  state.saving = mode === 'saving';
  elements.busyState.textContent = mode === 'loading' ? '読み込み中' : mode === 'saving' ? '保存中' : '待機中';
  elements.busyState.dataset.mode = mode || 'idle';
  elements.statusText.textContent = message;
}

function setStatus(message) {
  setBusy('', message);
}

function updateFormInputs() {
  elements.operatorNameInput.value = state.operator;
  elements.projectC2Input.value = state.project.c2 || '';
  elements.projectNameInput.value = state.project.projectName || '';
  elements.projectShortNameInput.value = state.project.shortName || '';
  elements.projectContactInput.value = state.project.contact || '';
  elements.drawingNumberInput.value = state.drawing.drawingNumber || '';
  elements.drawingStatusInput.value = state.drawing.drawingStatus || '';
  elements.filterInput.value = state.filterText || '';
}

function renderProjectSelect() {
  const currentValue = state.project.c2 || '';
  const options = ['<option value="">工事を選択</option>']
    .concat(
      state.projects.map((project) => {
        const label = [project.projectName || '-', `工事 ${project.c2}`];
        if (project.drawingCount || project.symbolCount) {
          label.push(`図面 ${project.drawingCount || 0}件 / 符号 ${project.symbolCount || 0}件`);
        }
        const selected = currentValue && currentValue === project.c2 ? ' selected' : '';
        return `<option value="${escapeHtml(project.c2)}"${selected}>${escapeHtml(label.join(' / '))}</option>`;
      })
    )
    .join('');

  elements.projectSelect.innerHTML = options;
}

function renderDrawingTabs() {
  if (!state.project.c2) {
    elements.drawingTabs.innerHTML = '<p class="empty-text">工事を選ぶと図面タブが出ます。</p>';
    return;
  }
  if (!state.drawings.length) {
    elements.drawingTabs.innerHTML = '<p class="empty-text">まだ図面がありません。図面番号を入れて読込すると、新しい図面として入力できます。</p>';
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

function buildLabelItems(rows) {
  const shortName = state.project.shortName || state.project.projectName || '未設定';
  const drawingLabel = state.drawing.drawingNumber ? `図面 ${state.drawing.drawingNumber}` : '';
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

function renderRows() {
  const visibleRows = state.rows.filter(matchesFilter);
  if (!visibleRows.length) {
    elements.tableBody.innerHTML = '<tr><td colspan="34" class="empty-text">表示できる符号がありません。</td></tr>';
    renderSummary([]);
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

  renderSummary(visibleRows);
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

function resetDrawingState() {
  state.selectedDrawingId = '';
  state.drawing = { ...DRAWING_TEMPLATE };
  state.rows = [];
  state.selectedRowIds = new Set();
}

function clearProjectState() {
  state.project = { ...PROJECT_TEMPLATE };
  state.drawings = [];
  resetDrawingState();
  renderAll();
}

function renderAll() {
  updateFormInputs();
  renderProjectSelect();
  renderDrawingTabs();
  renderRows();
}

async function refreshProjects() {
  setBusy('loading', '工事一覧を読み込んでいます。');
  try {
    state.projects = await listProjects(state.env);
    renderProjectSelect();
    setStatus('工事一覧を更新しました。');
  } catch (error) {
    console.error(error);
    setStatus(`工事一覧の読込に失敗しました: ${error.message}`);
  }
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
    resetDrawingState();
    renderAll();
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
    setStatus('図面番号を入れてください。');
    return;
  }

  const drawing = state.drawings.find((item) => item.drawingNumber === state.drawing.drawingNumber);
  const drawingId = drawing?.id || '';

  setBusy('loading', `${state.drawing.drawingNumber} の符号を読み込んでいます。`);
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
    renderAll();
    if (drawingId) {
      setStatus(`図面 ${state.drawing.drawingNumber} を読み込みました。`);
    } else {
      setStatus(`図面 ${state.drawing.drawingNumber} は未登録です。そのまま入力して保存できます。`);
    }
  } catch (error) {
    console.error(error);
    setStatus(`図面の読込に失敗しました: ${error.message}`);
  }
}

async function saveCurrent() {
  syncProjectFromForm();
  syncDrawingFromForm();

  if (!state.operator.trim()) {
    setStatus('作業者名を入れてください。');
    elements.operatorNameInput.focus();
    return;
  }

  setBusy('saving', '保存しています。');
  try {
    const response = await saveDrawingBundle({
      env: state.env,
      project: state.project,
      drawing: { ...state.drawing, id: state.selectedDrawingId },
      rows: state.rows,
      operator: state.operator
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
    renderAll();
    await refreshProjects();
    setStatus('保存しました。');
  } catch (error) {
    console.error(error);
    setStatus(`保存に失敗しました: ${error.message}`);
  }
}

function addRow() {
  state.rows.push(createUiRow());
  renderRows();
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
}

function bindEvents() {
  elements.operatorNameInput.addEventListener('input', () => {
    state.operator = elements.operatorNameInput.value.trim();
    localStorage.setItem('inputto_operator', state.operator);
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
  elements.saveButton.addEventListener('click', saveCurrent);
  elements.addRowButton.addEventListener('click', addRow);
  elements.duplicateRowButton.addEventListener('click', duplicateSelectedRows);
  elements.deleteRowsButton.addEventListener('click', deleteSelectedRows);
  elements.filterInput.addEventListener('input', () => {
    state.filterText = elements.filterInput.value;
    renderRows();
  });
  elements.printButton.addEventListener('click', () => window.print());

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
    renderSummary(state.rows.filter(matchesFilter));
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
      renderSummary(state.rows.filter(matchesFilter));
    });
  });
}

function renderInitialState() {
  renderTableHead();
  state.rows = [createUiRow()];
  renderAll();
}

async function bootstrap() {
  renderInitialState();
  bindEvents();
  setStatus('Firestore 版を初期化しました。');
  await refreshProjects();
}

bootstrap();
