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
    .sort((left, right) => String(left.symbol || '').localeCompare(String(right.symbol || ''), 'ja', { numeric: true }));
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

export async function saveDrawingBundle({ env, project, drawing, rows, operator }) {
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
  const nextRows = rows.map(sanitizeRow).filter((row) => hasAnyRowContent(row));
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

  const existingSnapshot = await getDocs(query(symbolsRef(env, cleanProject.c2), where('drawingId', '==', drawingId)));
  const existingIds = new Set(existingSnapshot.docs.map((item) => item.id));
  const nextIds = new Set();

  const batch = writeBatch(db);
  batch.set(
    drawingRef(env, cleanProject.c2, drawingId),
    {
      drawingId,
      drawingNumber: cleanDrawing.drawingNumber,
      drawingStatus: cleanDrawing.drawingStatus,
      contact: cleanProject.contact,
      rowCount: nextRows.length,
      symbolPreview: nextRows.slice(0, 6).map((row) => row.symbol),
      updatedAt: serverTimestamp(),
      updatedBy: sanitizeText(operator)
    },
    { merge: true }
  );

  nextRows.forEach((row) => {
    const docId = row.docId || makeSymbolId(row.symbol);
    nextIds.add(docId);
    batch.set(
      symbolRef(env, cleanProject.c2, docId),
      {
        ...row,
        c2: cleanProject.c2,
        projectName: cleanProject.projectName,
        shortName: cleanProject.shortName,
        contact: cleanProject.contact,
        drawingId,
        drawingNumber: cleanDrawing.drawingNumber,
        drawingStatus: cleanDrawing.drawingStatus,
        symbolN: normalizeText(row.symbol),
        floorN: normalizeText(row.floor),
        insideOutsideN: normalizeText(row.insideOutside),
        nameN: normalizeText(row.name),
        updatedAt: serverTimestamp(),
        updatedBy: sanitizeText(operator),
        createdAt: row.docId ? undefined : serverTimestamp(),
        createdBy: row.docId ? undefined : sanitizeText(operator)
      },
      { merge: true }
    );
  });

  existingIds.forEach((docId) => {
    if (!nextIds.has(docId)) {
      batch.delete(symbolRef(env, cleanProject.c2, docId));
    }
  });

  await batch.commit();

  const drawingsSnapshot = await getDocs(drawingsRef(env, cleanProject.c2));
  const drawings = drawingsSnapshot.docs.map((item) => ({ id: item.id, ...item.data() }));
  const drawingCount = drawings.length;
  const symbolCount = drawings.reduce((total, item) => total + Number(item.rowCount || 0), 0);

  await setDoc(
    projectRef(env, cleanProject.c2),
    {
      ...cleanProject,
      drawingCount,
      symbolCount,
      updatedAt: serverTimestamp(),
      updatedBy: sanitizeText(operator)
    },
    { merge: true }
  );

  const rowsSnapshot = await getDocs(query(symbolsRef(env, cleanProject.c2), where('drawingId', '==', drawingId)));

  return {
    project: cleanProject,
    drawing: { id: drawingId, ...cleanDrawing },
    drawings,
    rows: rowsSnapshot.docs
      .map((item) => ({
        uiId: crypto.randomUUID(),
        docId: item.id,
        ...item.data()
      }))
      .sort((left, right) => String(left.symbol || '').localeCompare(String(right.symbol || ''), 'ja', { numeric: true }))
  };
}

export const firestoreModel = {
  projectPath: 'environments/{env}/projects/{c2}',
  drawingPath: 'environments/{env}/projects/{c2}/drawings/{drawingId}',
  symbolPath: 'environments/{env}/projects/{c2}/symbols/{symbolId}'
};
