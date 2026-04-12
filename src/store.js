import {
  collection,
  doc,
  getDoc,
  getDocs,
  query,
  setDoc,
  writeBatch,
  where,
  serverTimestamp
} from 'firebase/firestore';
import { db } from './firebase.js';

function normalizeText(value) {
  return String(value || '')
    .normalize('NFKC')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');
}

function hashString(value) {
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}

export function makeDrawingId(drawingNumber) {
  const normalized = normalizeText(drawingNumber);
  return `drw_${hashString(normalized || crypto.randomUUID())}`;
}

export function makeSymbolId(symbol) {
  const normalized = normalizeText(symbol);
  return `sym_${hashString(normalized || crypto.randomUUID())}`;
}

function environmentsRef(env) {
  return doc(db, 'environments', env);
}

function projectsRef(env) {
  return collection(environmentsRef(env), 'projects');
}

function projectRef(env, c2) {
  return doc(projectsRef(env), String(c2).trim());
}

function drawingsRef(env, c2) {
  return collection(projectRef(env, c2), 'drawings');
}

function drawingRef(env, c2, drawingId) {
  return doc(drawingsRef(env, c2), drawingId);
}

function symbolsRef(env, c2) {
  return collection(projectRef(env, c2), 'symbols');
}

function symbolRef(env, c2, symbolId) {
  return doc(symbolsRef(env, c2), symbolId);
}

function sanitizeText(value) {
  return String(value || '').trim();
}

function sanitizeNumberText(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function sanitizeDateText(value) {
  if (!value) {
    return '';
  }
  return String(value).slice(0, 10);
}

function sanitizeProject(project) {
  return {
    c2: sanitizeText(project.c2),
    projectName: sanitizeText(project.projectName),
    shortName: sanitizeText(project.shortName),
    contact: sanitizeText(project.contact)
  };
}

function sanitizeDrawing(drawing) {
  return {
    drawingNumber: sanitizeText(drawing.drawingNumber),
    drawingStatus: sanitizeText(drawing.drawingStatus)
  };
}

function sanitizeRow(row) {
  return {
    docId: sanitizeText(row.docId),
    symbol: sanitizeText(row.symbol),
    name: sanitizeText(row.name),
    floor: sanitizeText(row.floor),
    left: sanitizeNumberText(row.left),
    right: sanitizeNumberText(row.right),
    doubleLeft: sanitizeNumberText(row.doubleLeft),
    doubleRight: sanitizeNumberText(row.doubleRight),
    noHand: sanitizeNumberText(row.noHand),
    width: sanitizeNumberText(row.width),
    height: sanitizeNumberText(row.height),
    frameDepth: sanitizeNumberText(row.frameDepth),
    dwLeft: sanitizeNumberText(row.dwLeft),
    dwRight: sanitizeNumberText(row.dwRight),
    dh: sanitizeNumberText(row.dh),
    insideOutside: sanitizeText(row.insideOutside),
    labelCount: sanitizeNumberText(row.labelCount),
    labelRightCount: sanitizeNumberText(row.labelRightCount),
    labelDoubleCount: sanitizeNumberText(row.labelDoubleCount),
    labelNoHandCount: sanitizeNumberText(row.labelNoHandCount),
    bakeColor: sanitizeText(row.bakeColor),
    floorQuantity: sanitizeNumberText(row.floorQuantity),
    gwDensity: sanitizeNumberText(row.gwDensity),
    gwThickness: sanitizeNumberText(row.gwThickness),
    rwDensity: sanitizeNumberText(row.rwDensity),
    rwThickness: sanitizeNumberText(row.rwThickness),
    draftAssignee: sanitizeText(row.draftAssignee),
    draftFrameAt: sanitizeDateText(row.draftFrameAt),
    draftDoorAt: sanitizeDateText(row.draftDoorAt),
    assemblyFrameCompletedAt: sanitizeDateText(row.assemblyFrameCompletedAt),
    assemblyDoorCompletedAt: sanitizeDateText(row.assemblyDoorCompletedAt),
    frameShipDate: sanitizeDateText(row.frameShipDate),
    doorShipDate: sanitizeDateText(row.doorShipDate)
  };
}

function hasAnyRowContent(row) {
  return Object.entries(row).some(([key, value]) => key !== 'docId' && String(value || '').trim() !== '');
}

const ROW_FIELDS = Object.keys(sanitizeRow({}));
const SYMBOL_PREVIEW_LIMIT = 6;
const MAX_BATCH_WRITES = 450;

function sortSymbolsByLabel(left, right) {
  return String(left.symbol || '').localeCompare(String(right.symbol || ''), 'ja', { numeric: true });
}

function chunkList(items, size) {
  const chunks = [];
  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size));
  }
  return chunks;
}

async function commitChunked(operations) {
  for (const chunk of chunkList(operations, MAX_BATCH_WRITES)) {
    const batch = writeBatch(db);
    chunk.forEach((applyOperation) => applyOperation(batch));
    await batch.commit();
  }
}

function indexSymbolEntries(entries) {
  const byId = new Map();
  const bySymbol = new Map();

  entries.forEach((entry) => {
    byId.set(entry.id, entry);
    const key = normalizeText(entry.symbol);
    if (!key) {
      return;
    }
    const list = bySymbol.get(key) || [];
    list.push(entry);
    bySymbol.set(key, list);
  });

  bySymbol.forEach((list) => {
    list.sort((left, right) => {
      const drawingDiff = Number(Boolean(right.drawingId)) - Number(Boolean(left.drawingId));
      if (drawingDiff !== 0) {
        return drawingDiff;
      }
      const rowDiff = String(left.drawingId || '').localeCompare(String(right.drawingId || ''), 'ja', {
        numeric: true
      });
      if (rowDiff !== 0) {
        return rowDiff;
      }
      return sortSymbolsByLabel(left, right);
    });
  });

  return { byId, bySymbol };
}

function mergeRowData(existingEntry, incomingRow, preserveExistingBlank) {
  const sanitized = sanitizeRow(incomingRow);

  if (!existingEntry || !preserveExistingBlank) {
    return sanitized;
  }

  const merged = { ...existingEntry };
  delete merged.id;

  ROW_FIELDS.forEach((key) => {
    if (sanitized[key] !== '') {
      merged[key] = sanitized[key];
      return;
    }
    if (merged[key] === undefined || merged[key] === null) {
      merged[key] = '';
    }
  });

  return merged;
}

function pickExistingEntry({ row, drawingId, byId, bySymbol }) {
  const rowEntry = row.docId ? byId.get(row.docId) : null;
  const symbolKey = normalizeText(row.symbol);
  const symbolEntries = symbolKey ? bySymbol.get(symbolKey) || [] : [];

  if (symbolEntries.length) {
    return (
      symbolEntries.find((entry) => entry.id === row.docId) ||
      symbolEntries.find((entry) => String(entry.drawingId || '') === drawingId) ||
      symbolEntries[0]
    );
  }

  return rowEntry || null;
}

function buildSymbolPayload({
  existingEntry,
  incomingRow,
  cleanProject,
  cleanDrawing,
  drawingId,
  operator,
  preserveExistingBlank
}) {
  const updatedBy = sanitizeText(operator);
  const merged = mergeRowData(existingEntry, incomingRow, preserveExistingBlank);
  const payload = {
    ...merged,
    c2: cleanProject.c2,
    projectName: cleanProject.projectName,
    shortName: cleanProject.shortName,
    contact: cleanProject.contact,
    drawingId,
    drawingNumber: cleanDrawing.drawingNumber,
    drawingStatus: cleanDrawing.drawingStatus,
    symbolN: normalizeText(merged.symbol),
    floorN: normalizeText(merged.floor),
    insideOutsideN: normalizeText(merged.insideOutside),
    nameN: normalizeText(merged.name),
    updatedAt: serverTimestamp(),
    updatedBy
  };

  if (!existingEntry) {
    payload.createdAt = serverTimestamp();
    payload.createdBy = updatedBy;
  }

  return payload;
}

async function recalculateProjectSummaries(env, c2, operator) {
  const [drawingSnapshot, symbolSnapshot] = await Promise.all([getDocs(drawingsRef(env, c2)), getDocs(symbolsRef(env, c2))]);
  const drawings = drawingSnapshot.docs.map((item) => ({ id: item.id, ...item.data() }));
  const symbols = symbolSnapshot.docs.map((item) => ({ id: item.id, ...item.data() }));
  const assignedSymbols = symbols.filter((row) => sanitizeText(row.drawingId));
  const assignedByDrawingId = new Map();

  assignedSymbols.forEach((row) => {
    const key = sanitizeText(row.drawingId);
    const list = assignedByDrawingId.get(key) || [];
    list.push(row);
    assignedByDrawingId.set(key, list);
  });

  const projectSymbolCount = drawings.reduce((total, drawing) => {
    const grouped = assignedByDrawingId.get(drawing.id) || [];
    return total + grouped.length;
  }, 0);

  const operations = drawings.map((drawing) => (batch) => {
    const grouped = assignedByDrawingId.get(drawing.id) || [];
    const sorted = [...grouped].sort(sortSymbolsByLabel);
    batch.set(
      drawingRef(env, c2, drawing.id),
      {
        rowCount: sorted.length,
        symbolPreview: sorted.slice(0, SYMBOL_PREVIEW_LIMIT).map((row) => row.symbol),
        updatedAt: serverTimestamp(),
        updatedBy: sanitizeText(operator)
      },
      { merge: true }
    );
  });

  operations.push((batch) => {
    batch.set(
      projectRef(env, c2),
      {
        drawingCount: drawings.length,
        symbolCount: projectSymbolCount,
        updatedAt: serverTimestamp(),
        updatedBy: sanitizeText(operator)
      },
      { merge: true }
    );
  });

  await commitChunked(operations);

  const refreshedDrawingsSnapshot = await getDocs(drawingsRef(env, c2));
  const refreshedDrawings = refreshedDrawingsSnapshot.docs
    .map((item) => ({ id: item.id, ...item.data() }))
    .sort((left, right) => String(left.drawingNumber || '').localeCompare(String(right.drawingNumber || ''), 'ja', { numeric: true }));

  return {
    drawings: refreshedDrawings,
    project: {
      c2: sanitizeText(c2),
      drawingCount: drawings.length,
      symbolCount: projectSymbolCount
    }
  };
}

function toComparableDate(value) {
  return value?.seconds ? value.seconds * 1000 : 0;
}

export async function listProjects(env) {
  const snapshot = await getDocs(projectsRef(env));
  return snapshot.docs
    .map((item) => ({ id: item.id, ...item.data() }))
    .sort((left, right) => {
      const dateDiff = toComparableDate(right.updatedAt) - toComparableDate(left.updatedAt);
      if (dateDiff !== 0) {
        return dateDiff;
      }
      return String(left.c2 || '').localeCompare(String(right.c2 || ''), 'ja');
    });
}

export async function loadProjectBundle(env, c2) {
  const projectSnapshot = await getDoc(projectRef(env, c2));
  const drawingsSnapshot = await getDocs(drawingsRef(env, c2));

  const project = projectSnapshot.exists()
    ? { c2: projectSnapshot.id, ...projectSnapshot.data() }
    : { c2: sanitizeText(c2), projectName: '', shortName: '', contact: '' };

  const drawings = drawingsSnapshot.docs
    .map((item) => ({ id: item.id, ...item.data() }))
    .sort((left, right) => String(left.drawingNumber || '').localeCompare(String(right.drawingNumber || ''), 'ja', { numeric: true }));

  return { project, drawings };
}

export async function loadDrawingRows(env, c2, drawingId) {
  const rowsQuery = query(symbolsRef(env, c2), where('drawingId', '==', drawingId));
  const snapshot = await getDocs(rowsQuery);
  return snapshot.docs
    .map((item) => ({
      uiId: crypto.randomUUID(),
      docId: item.id,
      ...item.data()
    }))
    .sort(sortSymbolsByLabel);
}

export async function loadProjectSymbols(env, c2) {
  const snapshot = await getDocs(symbolsRef(env, c2));
  return snapshot.docs
    .map((item) => ({
      uiId: crypto.randomUUID(),
      docId: item.id,
      ...item.data()
    }))
    .filter((item) => sanitizeText(item.drawingId))
    .sort((left, right) => {
      const drawingDiff = String(left.drawingNumber || '').localeCompare(String(right.drawingNumber || ''), 'ja', {
        numeric: true
      });
      if (drawingDiff !== 0) {
        return drawingDiff;
      }
      return String(left.symbol || '').localeCompare(String(right.symbol || ''), 'ja', { numeric: true });
    });
}

export async function assignSymbolsToDrawing({ env, project, drawing, symbolIds, operator }) {
  const cleanProject = await saveProjectHeader(env, project, operator);
  const cleanDrawing = sanitizeDrawing(drawing);
  const ids = Array.from(new Set((Array.isArray(symbolIds) ? symbolIds : []).map(sanitizeText).filter(Boolean)));

  if (!cleanDrawing.drawingNumber) {
    throw new Error('移動先の手配書Noを入力してください。');
  }
  if (!ids.length) {
    throw new Error('移動する符号を選択してください。');
  }

  const [drawingSnapshot, symbolSnapshot] = await Promise.all([
    getDocs(drawingsRef(env, cleanProject.c2)),
    getDocs(symbolsRef(env, cleanProject.c2))
  ]);
  const normalizedDrawingNumber = normalizeText(cleanDrawing.drawingNumber);
  const existingDrawing = drawingSnapshot.docs
    .map((item) => ({ id: item.id, ...item.data() }))
    .find((item) => normalizeText(item.drawingNumber) === normalizedDrawingNumber);
  const drawingId = sanitizeText(drawing.id) || existingDrawing?.id || makeDrawingId(cleanDrawing.drawingNumber);
  const drawingStatus = cleanDrawing.drawingStatus || sanitizeText(existingDrawing?.drawingStatus);
  const selectedIdSet = new Set(ids);
  const selectedEntries = symbolSnapshot.docs
    .map((item) => ({ id: item.id, ...item.data() }))
    .filter((item) => selectedIdSet.has(item.id));

  if (selectedEntries.length !== ids.length) {
    throw new Error('選択した符号の一部が見つかりません。登録済み一覧を読み込み直してください。');
  }

  const writes = [];
  writes.push((batch) => {
    batch.set(
      drawingRef(env, cleanProject.c2, drawingId),
      {
        drawingId,
        drawingNumber: cleanDrawing.drawingNumber,
        drawingStatus,
        contact: cleanProject.contact,
        updatedAt: serverTimestamp(),
        updatedBy: sanitizeText(operator)
      },
      { merge: true }
    );
  });

  selectedEntries.forEach((entry) => {
    writes.push((batch) => {
      batch.set(
        symbolRef(env, cleanProject.c2, entry.id),
        {
          c2: cleanProject.c2,
          projectName: cleanProject.projectName,
          shortName: cleanProject.shortName,
          contact: cleanProject.contact,
          drawingId,
          drawingNumber: cleanDrawing.drawingNumber,
          drawingStatus,
          updatedAt: serverTimestamp(),
          updatedBy: sanitizeText(operator)
        },
        { merge: true }
      );
    });
  });

  await commitChunked(writes);

  const summary = await recalculateProjectSummaries(env, cleanProject.c2, operator);
  const rowsSnapshot = await getDocs(query(symbolsRef(env, cleanProject.c2), where('drawingId', '==', drawingId)));

  return {
    project: {
      ...cleanProject,
      drawingCount: summary.project.drawingCount,
      symbolCount: summary.project.symbolCount
    },
    drawing: {
      id: drawingId,
      drawingNumber: cleanDrawing.drawingNumber,
      drawingStatus
    },
    drawings: summary.drawings,
    rows: rowsSnapshot.docs
      .map((item) => ({
        uiId: crypto.randomUUID(),
        docId: item.id,
        ...item.data()
      }))
      .sort(sortSymbolsByLabel)
  };
}

export async function saveProjectHeader(env, project, operator) {
  const cleanProject = sanitizeProject(project);
  if (!cleanProject.c2) {
    throw new Error('工事番号を入力してください。');
  }
  if (!cleanProject.projectName) {
    throw new Error('現場名を入力してください。');
  }

  await setDoc(
    projectRef(env, cleanProject.c2),
    {
      ...cleanProject,
      updatedAt: serverTimestamp(),
      updatedBy: sanitizeText(operator)
    },
    { merge: true }
  );

  return cleanProject;
}

async function saveDrawingBundleCore({ env, project, drawing, rows, operator, preserveExistingBlank }) {
  const cleanProject = await saveProjectHeader(env, project, operator);
  const cleanDrawing = sanitizeDrawing(drawing);

  if (!cleanDrawing.drawingNumber) {
    const drawingsSnapshot = await getDocs(drawingsRef(env, cleanProject.c2));
    return {
      project: cleanProject,
      drawing: null,
      drawings: drawingsSnapshot.docs
        .map((item) => ({ id: item.id, ...item.data() }))
        .sort((left, right) => String(left.drawingNumber || '').localeCompare(String(right.drawingNumber || ''), 'ja', { numeric: true })),
      rows: []
    };
  }

  const drawingId = sanitizeText(drawing.id) || makeDrawingId(cleanDrawing.drawingNumber);
  const nextRows = (Array.isArray(rows) ? rows : []).map(sanitizeRow).filter((row) => hasAnyRowContent(row));
  const comparableSymbols = new Set();

  nextRows.forEach((row) => {
    if (!row.symbol) {
      throw new Error('符号が空の行があります。');
    }
    const normalized = normalizeText(row.symbol);
    if (comparableSymbols.has(normalized)) {
      throw new Error(`同じ符号が重複しています: ${row.symbol}`);
    }
    comparableSymbols.add(normalized);
  });

  const existingSnapshot = await getDocs(symbolsRef(env, cleanProject.c2));
  const existingEntries = existingSnapshot.docs.map((item) => ({ id: item.id, ...item.data() }));
  const { byId, bySymbol } = indexSymbolEntries(existingEntries);
  const nextIds = new Set();
  const writes = [];

  writes.push((batch) => {
    batch.set(
      drawingRef(env, cleanProject.c2, drawingId),
      {
        drawingId,
        drawingNumber: cleanDrawing.drawingNumber,
        drawingStatus: cleanDrawing.drawingStatus,
        contact: cleanProject.contact,
        rowCount: nextRows.length,
        symbolPreview: nextRows.slice(0, SYMBOL_PREVIEW_LIMIT).map((row) => row.symbol),
        updatedAt: serverTimestamp(),
        updatedBy: sanitizeText(operator)
      },
      { merge: true }
    );
  });

  nextRows.forEach((row) => {
    const existingEntry = pickExistingEntry({ row, drawingId, byId, bySymbol });
    const docId = existingEntry?.id || row.docId || makeSymbolId(row.symbol);
    const isMove = Boolean(existingEntry) && String(existingEntry.drawingId || '') !== drawingId;
    const payload = buildSymbolPayload({
      existingEntry,
      incomingRow: row,
      cleanProject,
      cleanDrawing,
      drawingId,
      operator,
      preserveExistingBlank: Boolean(preserveExistingBlank || isMove)
    });

    nextIds.add(docId);
    writes.push((batch) => {
      batch.set(symbolRef(env, cleanProject.c2, docId), payload, { merge: true });
    });
  });

  existingEntries.forEach((entry) => {
    if (!nextIds.has(entry.id) && String(entry.drawingId || '') === drawingId) {
      writes.push((batch) => {
        batch.set(
          symbolRef(env, cleanProject.c2, entry.id),
          {
            drawingId: '',
            drawingNumber: '',
            drawingStatus: '',
            updatedAt: serverTimestamp(),
            updatedBy: sanitizeText(operator)
          },
          { merge: true }
        );
      });
    }
  });

  await commitChunked(writes);

  const summary = await recalculateProjectSummaries(env, cleanProject.c2, operator);
  const rowsSnapshot = await getDocs(query(symbolsRef(env, cleanProject.c2), where('drawingId', '==', drawingId)));

  return {
    project: {
      ...cleanProject,
      drawingCount: summary.project.drawingCount,
      symbolCount: summary.project.symbolCount
    },
    drawing: { id: drawingId, ...cleanDrawing },
    drawings: summary.drawings,
    rows: rowsSnapshot.docs
      .map((item) => ({
        uiId: crypto.randomUUID(),
        docId: item.id,
        ...item.data()
      }))
      .sort(sortSymbolsByLabel)
  };
}

export async function saveDrawingBundle({ env, project, drawing, rows, operator }) {
  return saveDrawingBundleCore({ env, project, drawing, rows, operator, preserveExistingBlank: false });
}

export async function saveDrawingSymbolBatch({ env, project, drawing, symbols, operator }) {
  const rows = (Array.isArray(symbols) ? symbols : []).map((item) => (typeof item === 'string' ? { symbol: item } : item));
  return saveDrawingBundleCore({ env, project, drawing, rows, operator, preserveExistingBlank: true });
}

export const firestoreModel = {
  projectPath: 'environments/{env}/projects/{c2}',
  drawingPath: 'environments/{env}/projects/{c2}/drawings/{drawingId}',
  symbolPath: 'environments/{env}/projects/{c2}/symbols/{symbolId}'
};
