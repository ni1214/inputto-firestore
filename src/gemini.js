const GEMINI_API_BASE = 'https://generativelanguage.googleapis.com/v1beta';
const GEMINI_UPLOAD_BASE = 'https://generativelanguage.googleapis.com/upload/v1beta';
const GEMINI_API_KEY_STORAGE_KEY = 'inputto_gemini_api_key';
const DEFAULT_GEMINI_MODEL = 'gemini-2.5-flash';

const FIELD_KEYS = [
  'symbol',
  'name',
  'floor',
  'left',
  'right',
  'doubleLeft',
  'doubleRight',
  'noHand',
  'width',
  'height',
  'frameDepth',
  'dwLeft',
  'dwRight',
  'dh',
  'insideOutside',
  'labelCount',
  'labelRightCount',
  'labelDoubleCount',
  'labelNoHandCount',
  'bakeColor',
  'floorQuantity',
  'gwDensity',
  'gwThickness',
  'rwDensity',
  'rwThickness',
  'draftAssignee',
  'draftFrameAt',
  'draftDoorAt',
  'assemblyFrameCompletedAt',
  'assemblyDoorCompletedAt',
  'frameShipDate',
  'doorShipDate'
];

function hasLocalStorage() {
  return typeof window !== 'undefined' && typeof window.localStorage !== 'undefined';
}

function cleanText(value) {
  return String(value ?? '')
    .normalize('NFKC')
    .trim();
}

function cleanKey(value) {
  return cleanText(value).replace(/\s+/g, '');
}

function stripJsonFence(value) {
  const text = cleanText(value);
  if (!text.startsWith('```')) {
    return text;
  }
  return text
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/```$/i, '')
    .trim();
}

export function loadStoredGeminiApiKey() {
  if (!hasLocalStorage()) {
    return '';
  }
  return cleanKey(window.localStorage.getItem(GEMINI_API_KEY_STORAGE_KEY));
}

export function saveGeminiApiKey(apiKey) {
  const key = cleanKey(apiKey);
  if (!key) {
    throw new Error('Gemini API key is empty.');
  }
  if (!hasLocalStorage()) {
    return key;
  }
  window.localStorage.setItem(GEMINI_API_KEY_STORAGE_KEY, key);
  return key;
}

export function clearStoredGeminiApiKey() {
  if (hasLocalStorage()) {
    window.localStorage.removeItem(GEMINI_API_KEY_STORAGE_KEY);
  }
}

export function buildHandaiExtractionPrompt() {
  return [
    'あなたは日本語の手配依頼書や図面台帳を読み取る業務アシスタントです。',
    '添付されたPDFから、工事情報と手配書ごとの符号一覧を読み取り、次のJSONだけを返してください。',
    '不明な値は空文字にしてください。推測が必要な場合は、推測した値を入れたうえで errors に短く理由を入れてください。',
    '手配書No が 1-1 / 1-2 のように分割されている場合は、その表記をそのまま drawing.drawingNumber に入れてください。',
    '符号は 1 行 1 件で rows に入れてください。読み取れた列はできるだけ埋めてください。',
    '工事名の正式名と略称が両方わかるなら両方入れてください。略称はラベルに載せる短い名称です。',
    '手配書の状態は drawing.drawingStatus に入れてください。内部 / 外部 / 変更後 / 未登録 など、PDFにある表記を優先してください。',
    '返却JSONの最上位キーは project, drawing, rows, errors です。'
  ].join('\n');
}

export function buildHandaiExtractionSchema() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      project: {
        type: 'object',
        additionalProperties: false,
        properties: {
          c2: { type: 'string', description: '工事番号。' },
          projectName: { type: 'string', description: '工事の正式名称。' },
          shortName: { type: 'string', description: 'ラベル用の略称。' },
          contact: { type: 'string', description: '担当者名。' }
        },
        required: ['c2', 'projectName', 'shortName', 'contact']
      },
      drawing: {
        type: 'object',
        additionalProperties: false,
        properties: {
          drawingNumber: { type: 'string', description: '手配書No。' },
          drawingStatus: { type: 'string', description: '手配書状態。' }
        },
        required: ['drawingNumber', 'drawingStatus']
      },
      rows: {
        type: 'array',
        description: '1行1符号の一覧。',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: FIELD_KEYS.reduce((acc, key) => {
            acc[key] = { type: 'string', description: `${key} field.` };
            return acc;
          }, {}),
          required: FIELD_KEYS
        }
      },
      errors: {
        type: 'array',
        items: { type: 'string' }
      }
    },
    required: ['project', 'drawing', 'rows', 'errors']
  };
}

function normalizeExtraction(data) {
  const source = data && typeof data === 'object' ? data : {};
  const project = source.project && typeof source.project === 'object' ? source.project : {};
  const drawing = source.drawing && typeof source.drawing === 'object' ? source.drawing : {};
  const rows = Array.isArray(source.rows) ? source.rows : [];
  const errors = Array.isArray(source.errors) ? source.errors : [];

  return {
    project: {
      c2: cleanText(project.c2),
      projectName: cleanText(project.projectName),
      shortName: cleanText(project.shortName),
      contact: cleanText(project.contact)
    },
    drawing: {
      drawingNumber: cleanText(drawing.drawingNumber),
      drawingStatus: cleanText(drawing.drawingStatus)
    },
    rows: rows.map((row) => {
      const output = {};
      FIELD_KEYS.forEach((key) => {
        output[key] = cleanText(row?.[key]);
      });
      return output;
    }),
    errors: errors.map((item) => cleanText(item)).filter(Boolean)
  };
}

function getUploadHeaders(file) {
  return {
    'Content-Type': 'application/json',
    'X-Goog-Upload-Protocol': 'resumable',
    'X-Goog-Upload-Command': 'start',
    'X-Goog-Upload-Header-Content-Length': String(file.size),
    'X-Goog-Upload-Header-Content-Type': file.type || 'application/pdf'
  };
}

async function startResumableUpload(file, apiKey, fetchImpl = fetch) {
  const response = await fetchImpl(`${GEMINI_UPLOAD_BASE}/files?key=${encodeURIComponent(apiKey)}`, {
    method: 'POST',
    headers: getUploadHeaders(file),
    body: JSON.stringify({
      file: {
        display_name: file.name
      }
    })
  });

  if (!response.ok) {
    throw new Error(`Gemini upload start failed: ${response.status} ${response.statusText}`);
  }

  const uploadUrl =
    response.headers.get('x-goog-upload-url') ||
    response.headers.get('X-Goog-Upload-URL') ||
    response.headers.get('x-goog-upload-url'.toLowerCase());

  if (!uploadUrl) {
    throw new Error('Gemini upload URL was not returned.');
  }

  return uploadUrl;
}

async function finalizeResumableUpload(file, uploadUrl, fetchImpl = fetch) {
  const response = await fetchImpl(uploadUrl, {
    method: 'POST',
    headers: {
      'X-Goog-Upload-Offset': '0',
      'X-Goog-Upload-Command': 'upload, finalize'
    },
    body: await file.arrayBuffer()
  });

  if (!response.ok) {
    throw new Error(`Gemini upload finalize failed: ${response.status} ${response.statusText}`);
  }

  const payload = await response.json();
  const uploadedFile = payload?.file;
  if (!uploadedFile) {
    throw new Error('Gemini upload response did not include file metadata.');
  }
  return uploadedFile;
}

async function getFileMetadata(name, apiKey, fetchImpl = fetch) {
  const response = await fetchImpl(`${GEMINI_API_BASE}/files/${encodeURIComponent(name)}?key=${encodeURIComponent(apiKey)}`);
  if (!response.ok) {
    throw new Error(`Gemini file lookup failed: ${response.status} ${response.statusText}`);
  }
  const payload = await response.json();
  return payload?.file || payload;
}

async function waitForFileReady(fileInfo, apiKey, fetchImpl = fetch, timeoutMs = 120000, pollIntervalMs = 1500) {
  const startedAt = Date.now();
  let current = fileInfo;

  while (true) {
    const state = String(current?.state || '').toUpperCase();
    if (!state || state === 'ACTIVE') {
      return current;
    }
    if (state === 'FAILED') {
      throw new Error('Gemini file processing failed.');
    }
    if (Date.now() - startedAt > timeoutMs) {
      throw new Error('Timed out waiting for Gemini file processing.');
    }
    await new Promise((resolve) => setTimeout(resolve, pollIntervalMs));
    current = await getFileMetadata(current.name, apiKey, fetchImpl);
  }
}

async function callGenerateContent({ apiKey, model, fileInfo, prompt, fetchImpl = fetch }) {
  const response = await fetchImpl(`${GEMINI_API_BASE}/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      contents: [
        {
          role: 'user',
          parts: [
            {
              text: prompt
            },
            {
              file_data: {
                mime_type: fileInfo.mimeType || 'application/pdf',
                file_uri: fileInfo.uri
              }
            }
          ]
        }
      ],
      generationConfig: {
        responseMimeType: 'application/json',
        responseJsonSchema: buildHandaiExtractionSchema()
      }
    })
  });

  if (!response.ok) {
    throw new Error(`Gemini generation failed: ${response.status} ${response.statusText}`);
  }

  return response.json();
}

function extractResponseText(payload) {
  const candidate = payload?.candidates?.[0];
  const parts = candidate?.content?.parts || [];
  const text = parts
    .map((part) => part?.text || '')
    .join('')
    .trim();
  if (!text) {
    throw new Error('Gemini response did not include text output.');
  }
  return text;
}

export async function uploadPdfToGeminiFile(file, apiKey = loadStoredGeminiApiKey(), options = {}) {
  const resolvedKey = cleanKey(apiKey);
  if (!resolvedKey) {
    throw new Error('Gemini API key is required.');
  }
  if (!file) {
    throw new Error('PDF file is required.');
  }

  const fetchImpl = options.fetchImpl || fetch;
  const uploadUrl = await startResumableUpload(file, resolvedKey, fetchImpl);
  const uploadedFile = await finalizeResumableUpload(file, uploadUrl, fetchImpl);
  return waitForFileReady(
    uploadedFile,
    resolvedKey,
    fetchImpl,
    options.timeoutMs ?? 120000,
    options.pollIntervalMs ?? 1500
  );
}

export async function extractHandaiDataFromPdf(file, options = {}) {
  const apiKey = cleanKey(options.apiKey ?? loadStoredGeminiApiKey());
  if (!apiKey) {
    throw new Error('Gemini API key is required.');
  }

  const model = options.model || DEFAULT_GEMINI_MODEL;
  const fetchImpl = options.fetchImpl || fetch;
  const prompt = options.prompt || buildHandaiExtractionPrompt();

  const uploadedFile = await uploadPdfToGeminiFile(file, apiKey, {
    fetchImpl,
    timeoutMs: options.timeoutMs,
    pollIntervalMs: options.pollIntervalMs
  });
  const response = await callGenerateContent({
    apiKey,
    model,
    fileInfo: uploadedFile,
    prompt,
    fetchImpl
  });

  const text = stripJsonFence(extractResponseText(response));
  const parsed = normalizeExtraction(JSON.parse(text));

  return parsed;
}

export const geminiPdfExtraction = {
  storageKey: GEMINI_API_KEY_STORAGE_KEY,
  defaultModel: DEFAULT_GEMINI_MODEL,
  fieldKeys: FIELD_KEYS,
  buildHandaiExtractionPrompt,
  buildHandaiExtractionSchema,
  loadStoredGeminiApiKey,
  saveGeminiApiKey,
  clearStoredGeminiApiKey,
  uploadPdfToGeminiFile,
  extractHandaiDataFromPdf
};
