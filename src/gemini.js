import { GoogleAIBackend, Schema, getAI, getGenerativeModel } from 'firebase/ai';
import { app, firebaseConfig } from './firebase.js';

const DEFAULT_GEMINI_MODEL = 'gemini-2.5-flash';
const MAX_GENERATION_ATTEMPTS = 3;
const RETRIABLE_STATUS_CODES = new Set([429, 500, 503, 504]);
const AI_LOGIC_SETUP_URL = `https://console.firebase.google.com/project/${firebaseConfig.projectId}/ailogic/`;

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

const ai = getAI(app, { backend: new GoogleAIBackend() });

function cleanText(value) {
  return String(value ?? '')
    .normalize('NFKC')
    .trim();
}

function stripJsonFence(value) {
  const text = cleanText(value);
  if (!text.startsWith('```')) {
    return text;
  }
  return text.replace(/^```(?:json)?\s*/i, '').replace(/```$/i, '').trim();
}

function wait(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function getErrorStatus(error) {
  const directStatus = Number(error?.customErrorData?.status);
  if (Number.isFinite(directStatus) && directStatus > 0) {
    return directStatus;
  }

  const message = String(error?.message || '');
  const matched = message.match(/\[(\d{3})\s+[^\]]+\]/);
  return matched ? Number(matched[1]) : 0;
}

function isRetriableError(error) {
  return RETRIABLE_STATUS_CODES.has(getErrorStatus(error));
}

function toFriendlyError(error) {
  const code = String(error?.code || '');
  const status = getErrorStatus(error);

  if (code.includes('api-not-enabled')) {
    return new Error(`Firebase AI Logic を有効にしてください: ${AI_LOGIC_SETUP_URL}`);
  }
  if (status === 503) {
    return new Error('AI側が混み合っているため読込できませんでした。少し待ってからもう一度試してください。');
  }
  if (status === 429) {
    return new Error('AIの利用回数が集中しています。少し待ってから再実行してください。');
  }
  if (status === 403) {
    return new Error(`Firebase AI Logic の利用設定を確認してください: ${AI_LOGIC_SETUP_URL}`);
  }
  if (status === 400) {
    return new Error('AIへの送信内容に問題がありました。アプリ側を修正しましたので、もう一度試してください。');
  }
  if (error instanceof Error) {
    return error;
  }
  return new Error('PDFの解析に失敗しました。');
}

export function buildHandaiExtractionPrompt() {
  return [
    'You are reading a Japanese fabrication order PDF for doors and frames.',
    'Return only JSON that matches the requested schema.',
    'Extract project information, the current drawing or order number, the drawing status, and every symbol row.',
    'Keep each symbol as one row.',
    'If a value is missing, return an empty string.',
    'Use the drawing number exactly as written, including values like 1-1 or 1-2.',
    'Use shortName for the label-friendly site abbreviation.',
    'drawingStatus should be short text such as 外部, 内部, 共通, or 未定.',
    'Rows should include manufacturing columns like W, H, frame depth, DW, DH, inside or outside, bake color, GW or RW density and thickness, label counts, and dates whenever found.',
    'If something is uncertain, still return your best effort and note the issue in errors.',
    'Do not include markdown fences.',
    'JSON root keys must be: project, drawing, rows, errors.'
  ].join('\n');
}

function stringField(description) {
  return Schema.string({
    description,
    nullable: true
  });
}

export function buildHandaiExtractionSchema() {
  const rowProperties = FIELD_KEYS.reduce((acc, key) => {
    acc[key] = stringField(`${key} field.`);
    return acc;
  }, {});

  return Schema.object({
    properties: {
      project: Schema.object({
        properties: {
          c2: stringField('Project code.'),
          projectName: stringField('Project name.'),
          shortName: stringField('Short label-friendly project name.'),
          contact: stringField('Operator or contact name.')
        },
        optionalProperties: ['c2', 'projectName', 'shortName', 'contact']
      }),
      drawing: Schema.object({
        properties: {
          drawingNumber: stringField('Drawing or order number.'),
          drawingStatus: stringField('Drawing status like 外部 or 内部.')
        },
        optionalProperties: ['drawingNumber', 'drawingStatus']
      }),
      rows: Schema.array({
        items: Schema.object({
          properties: rowProperties,
          optionalProperties: FIELD_KEYS
        })
      }),
      errors: Schema.array({
        items: stringField('Extraction warning or note.')
      })
    },
    optionalProperties: ['project', 'drawing', 'rows', 'errors']
  });
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

async function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('PDFの読込に失敗しました。'));
    reader.onload = () => {
      const dataUrl = String(reader.result || '');
      const base64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : '';
      if (!base64) {
        reject(new Error('PDFをBase64へ変換できませんでした。'));
        return;
      }
      resolve(base64);
    };
    reader.readAsDataURL(blob);
  });
}

async function fileToInlineDataPart(file) {
  if (!file) {
    throw new Error('PDFファイルがありません。');
  }

  const mimeType = file.type || 'application/pdf';
  const fileName = String(file.name || '');
  if (mimeType !== 'application/pdf' && !/\.pdf$/i.test(fileName)) {
    throw new Error('PDFファイルを選んでください。');
  }

  return {
    inlineData: {
      mimeType: 'application/pdf',
      data: await blobToBase64(file)
    }
  };
}

async function generateWithRetry({ prompt, pdfPart, modelName }) {
  let lastError = null;

  for (let attempt = 1; attempt <= MAX_GENERATION_ATTEMPTS; attempt += 1) {
    try {
      const model = getGenerativeModel(ai, {
        model: modelName,
        generationConfig: {
          temperature: 0,
          responseMimeType: 'application/json',
          responseSchema: buildHandaiExtractionSchema()
        }
      });
      const result = await model.generateContent([prompt, pdfPart]);
      return result.response.text();
    } catch (error) {
      lastError = error;
      if (!isRetriableError(error) || attempt >= MAX_GENERATION_ATTEMPTS) {
        break;
      }
      await wait(900 * attempt);
    }
  }

  throw toFriendlyError(lastError);
}

export async function extractHandaiDataFromPdf(file, options = {}) {
  const model = cleanText(options.model || DEFAULT_GEMINI_MODEL) || DEFAULT_GEMINI_MODEL;
  const prompt = cleanText(options.prompt || buildHandaiExtractionPrompt()) || buildHandaiExtractionPrompt();
  const pdfPart = await fileToInlineDataPart(file);
  const text = stripJsonFence(await generateWithRetry({ prompt, pdfPart, modelName: model }));

  try {
    return normalizeExtraction(JSON.parse(text));
  } catch (error) {
    throw new Error(`AIの返答を読めませんでした: ${error.message}`);
  }
}

export function getAiLogicSetupUrl() {
  return AI_LOGIC_SETUP_URL;
}

export const geminiPdfExtraction = {
  defaultModel: DEFAULT_GEMINI_MODEL,
  fieldKeys: FIELD_KEYS,
  buildHandaiExtractionPrompt,
  buildHandaiExtractionSchema,
  extractHandaiDataFromPdf,
  getAiLogicSetupUrl
};
