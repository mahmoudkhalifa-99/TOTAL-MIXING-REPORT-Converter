import express from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { GoogleGenAI, Type, ThinkingLevel } from '@google/genai';
import { autoSplitCodeColumn } from './src/utils/codeSplitter.ts';
import { ensureItemTypeColumn } from './src/utils/itemClassifier.ts';
import { propagateFormulaNames } from './src/utils/formulaPropagator.ts';
import type { SheetData } from './src/types/converter.ts';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = Number(process.env.PORT) || 3000;

// Enable CORS for iframe and cross-origin environments
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

// High payload limit for PDFs and images (100mb)
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

// Initialize Google GenAI with generous timeout for multi-page PDF streaming
const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    timeout: 360000,
    headers: {
      'User-Agent': 'aistudio-build',
    },
  },
});

// JSON Schema for document conversion
const conversionSchema = {
  type: Type.OBJECT,
  properties: {
    documentTitle: {
      type: Type.STRING,
      description: 'عنوان وصفي للمستند والملف',
    },
    documentSummary: {
      type: Type.STRING,
      description: 'ملخص موجز لمحتوى الملف والبيانات المستخرجة',
    },
    detectedLanguage: {
      type: Type.STRING,
      description: 'اللغة المكتشفة مثل ar أو en أو ثنائية اللغة',
    },
    detectedType: {
      type: Type.STRING,
      description: 'نوع المستند مثل mixing_report, invoice, bank_statement, table',
    },
    sheets: {
      type: Type.ARRAY,
      description: 'أوراق العمل المستخرجة للإكسل (Sheets)',
      items: {
        type: Type.OBJECT,
        properties: {
          sheetName: {
            type: Type.STRING,
            description: 'اسم ورقة العمل في إكسل (أقل من 31 حرف)',
          },
          description: {
            type: Type.STRING,
            description: 'وصف لمحتويات ورقة العمل',
          },
          isRtl: {
            type: Type.BOOLEAN,
            description: 'صحيح إذا كان الجدول باللغة العربية، وخطأ إذا كان بالإنجليزية',
          },
          headers: {
            type: Type.ARRAY,
            description: 'عناوين الأعمدة في الجدول',
            items: { type: Type.STRING },
          },
          rows: {
            type: Type.ARRAY,
            description: 'صفوف البيانات',
            items: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
            },
          },
          totalsRow: {
            type: Type.ARRAY,
            description: 'صف المجموع أو الإجماليات اختياري',
            items: { type: Type.STRING },
          },
        },
        required: ['sheetName', 'isRtl', 'headers', 'rows'],
      },
    },
    metadata: {
      type: Type.OBJECT,
      properties: {
        pageCountEstimate: { type: Type.STRING },
        rowCountTotal: { type: Type.INTEGER },
        columnCountMax: { type: Type.INTEGER },
      },
    },
    notes: {
      type: Type.ARRAY,
      description: 'ملاحظات وتنبيهات حول جودة الاستخراج',
      items: { type: Type.STRING },
    },
  },
  required: ['documentTitle', 'sheets'],
};

// Normalize MIME types so non-standard browser types like image/jpg work cleanly with Gemini
function normalizeMimeType(mimeType: string, fileName: string, dataUrl?: string): string {
  let detected = (mimeType || '').trim().toLowerCase();

  if ((!detected || detected === 'application/octet-stream') && dataUrl) {
    const match = dataUrl.match(/^data:([^;]+);base64,/i);
    if (match && match[1]) {
      detected = match[1].trim().toLowerCase();
    }
  }

  if (detected === 'image/jpg') detected = 'image/jpeg';

  if (!detected || detected === 'application/octet-stream') {
    const lowerName = (fileName || '').toLowerCase();
    if (lowerName.endsWith('.png')) return 'image/png';
    if (lowerName.endsWith('.jpg') || lowerName.endsWith('.jpeg')) return 'image/jpeg';
    if (lowerName.endsWith('.webp')) return 'image/webp';
    return 'application/pdf';
  }

  return detected;
}

// Parse Gemini JSON safely, stripping markdown fences and recovering complete rows if truncated
function parseGeminiJson(rawText: string): { data: any; wasRecovered: boolean } {
  let cleaned = (rawText || '').trim();

  // Strip markdown fences if present
  if (cleaned.startsWith('```')) {
    cleaned = cleaned.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
  }

  try {
    return { data: JSON.parse(cleaned), wasRecovered: false };
  } catch {
    // Attempt to extract JSON object bounds
    const firstBrace = cleaned.indexOf('{');
    if (firstBrace !== -1) {
      const sub = cleaned.slice(firstBrace);
      try {
        return { data: JSON.parse(sub), wasRecovered: false };
      } catch {
        // Attempt graceful recovery by scanning backwards through recent complete row boundaries
        let searchFrom = sub.length;
        for (let step = 0; step < 8; step++) {
          const idx = sub.lastIndexOf('],', searchFrom);
          if (idx === -1) break;
          const prefix = sub.slice(0, idx + 1);
          const candidates = [prefix + ']}]}', prefix + ']}', prefix + '}]}'];
          for (const candidate of candidates) {
            try {
              const recovered = JSON.parse(candidate);
              if (recovered && Array.isArray(recovered.sheets) && recovered.sheets.length > 0) {
                return { data: recovered, wasRecovered: true };
              }
            } catch {
              // continue trying
            }
          }
          searchFrom = idx - 1;
        }
      }
    }
    throw new Error('INVALID_JSON_RESPONSE');
  }
}

// Ensure every sheet has well-formed headers and rows arrays
function sanitizeSheetData(sheet: any, index: number): SheetData {
  const safeHeaders: string[] = Array.isArray(sheet?.headers)
    ? sheet.headers.map((h: any, i: number) => {
        const str = h === null || h === undefined ? '' : String(h).trim();
        return str || `Column_${i + 1}`;
      })
    : [];

  const safeRows: string[][] = Array.isArray(sheet?.rows)
    ? sheet.rows
        .filter((row: any) => Array.isArray(row))
        .map((row: any[]) => row.map((cell: any) => (cell === null || cell === undefined ? '' : String(cell))))
    : [];

  const safeTotalsRow: string[] | undefined = Array.isArray(sheet?.totalsRow)
    ? sheet.totalsRow.map((cell: any) => (cell === null || cell === undefined ? '' : String(cell)))
    : undefined;

  return {
    sheetName:
      typeof sheet?.sheetName === 'string' && sheet.sheetName.trim()
        ? sheet.sheetName.trim()
        : `Sheet_${index + 1}`,
    description: typeof sheet?.description === 'string' ? sheet.description : '',
    isRtl: typeof sheet?.isRtl === 'boolean' ? sheet.isRtl : false,
    headers: safeHeaders,
    rows: safeRows,
    ...(safeTotalsRow && safeTotalsRow.length > 0 ? { totalsRow: safeTotalsRow } : {}),
  };
}

// Fast and resilient streaming generator with stall detection, abort cleanup, and active model fallback
async function generateWithFallback(
  filePart: any,
  textPart: any,
  systemInstruction: string,
  onProgress?: (step: string) => void
): Promise<{ rawText: string; modelUsed: string }> {
  // Active supported models from @google/genai guidelines
  const candidateModels = ['gemini-4-hydrogen', 'gemini-3.1-flash-lite'];
  let lastError: any = null;

  for (const model of candidateModels) {
    // Large multi-page PDFs with structured schema can take >90s before emitting the first JSON token
    const firstByteStallMs = 140000;
    const chunkStallMs = 90000;
    const hardMaxMs = model === 'gemini-4-hydrogen' ? 260000 : 300000;
    const maxAttempts = model === 'gemini-4-hydrogen' ? 2 : 3;

    for (let attempt = 0; attempt < maxAttempts; attempt++) {
      const abortController = new AbortController();
      try {
        if (attempt > 0) {
          const delay = 1400 * attempt + Math.floor(Math.random() * 600);
          console.log(`[Gemini] Retrying model "${model}" (attempt ${attempt + 1}/${maxAttempts}) after ${delay}ms...`);
          onProgress?.('جاري إعادة المحاولة السريعة مع نموذج الاستخراج...');
          await new Promise((r) => setTimeout(r, delay));
        }

        console.log(`[Gemini] Requesting streaming extraction with model: "${model}" (attempt ${attempt + 1}/${maxAttempts})...`);
        onProgress?.(
          model === 'gemini-4-hydrogen'
            ? 'جاري تحليل المستند بالذكاء الاصطناعي...'
            : 'جاري الاستخراج السريع لكافة صفحات المستند...'
        );
        const startTime = Date.now();

        const config: any = {
          systemInstruction,
          responseMimeType: 'application/json',
          responseSchema: conversionSchema,
          abortSignal: abortController.signal,
        };

        if (model === 'gemini-4-hydrogen') {
          config.thinkingConfig = { thinkingLevel: ThinkingLevel.LOW };
        }

        let accumulatedText = '';
        let gotFirstChunk = false;
        let lastChunkAt = Date.now();

        const streamTask = (async () => {
          const stream = await ai.models.generateContentStream({
            model,
            contents: { parts: [filePart, textPart] },
            config,
          });

          for await (const chunk of stream) {
            const piece = chunk.text || '';
            if (piece) {
              gotFirstChunk = true;
              lastChunkAt = Date.now();
              accumulatedText += piece;

              // Estimate extracted row count from array delimiters for reassuring UI feedback
              const approxRows = (accumulatedText.match(/\],/g) || []).length;
              if (approxRows > 0 && approxRows % 5 === 0) {
                onProgress?.(`جاري استخراج الجداول والتركيبات... (${approxRows} صف)`);
              } else if (approxRows === 0) {
                onProgress?.('جاري استخراج البيانات من صفحات الملف...');
              }
            }
          }
          return accumulatedText;
        })();

        // Watchdog timer that checks stall / hard limits and aborts stalled underlying requests
        let watchdogInterval: ReturnType<typeof setInterval> | undefined;
        const watchdogPromise = new Promise<never>((_, reject) => {
          watchdogInterval = setInterval(() => {
            const now = Date.now();
            const totalElapsed = now - startTime;
            const sinceLastActivity = now - (gotFirstChunk ? lastChunkAt : startTime);
            const activeStallLimit = gotFirstChunk ? chunkStallMs : firstByteStallMs;

            if (sinceLastActivity > activeStallLimit || totalElapsed > hardMaxMs) {
              abortController.abort();
              const timeoutErr: any = new Error(
                `MODEL_STALL_TIMEOUT_${Math.round(totalElapsed / 1000)}S`
              );
              timeoutErr.status = 504;
              reject(timeoutErr);
            }
          }, 1500);
        });

        const rawText = await Promise.race([streamTask, watchdogPromise]).finally(() => {
          if (watchdogInterval) clearInterval(watchdogInterval);
        });

        const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
        console.log(`[Gemini] Done streaming with model "${model}" in ${elapsed}s (length=${rawText.length})`);
        return { rawText, modelUsed: model };
      } catch (err: any) {
        abortController.abort();
        lastError = err;
        const errMsg = String(err?.message || '');
        const status = Number(err?.status || err?.code || 0);

        if (
          status === 403 ||
          errMsg.includes('PERMISSION_DENIED') ||
          errMsg.includes('API_KEY_INVALID')
        ) {
          throw err;
        }

        const isTimeout =
          status === 504 ||
          errMsg.includes('DEADLINE_EXCEEDED') ||
          errMsg.includes('ETIMEDOUT') ||
          errMsg.includes('MODEL_TIMEOUT') ||
          errMsg.includes('MODEL_STALL_TIMEOUT') ||
          errMsg.includes('504');

        console.log(`[Gemini] Model "${model}" attempt ${attempt + 1} did not complete (status=${status}):`, errMsg.slice(0, 220));

        if (isTimeout) {
          break;
        }

        const isTransient =
          status === 503 ||
          status === 429 ||
          status === 500 ||
          status === 502 ||
          errMsg.includes('503') ||
          errMsg.includes('429') ||
          errMsg.includes('high demand') ||
          errMsg.includes('UNAVAILABLE') ||
          errMsg.includes('RESOURCE_EXHAUSTED') ||
          errMsg.includes('ECONNRESET') ||
          errMsg.includes('fetch failed');

        if (isTransient && attempt < maxAttempts - 1) {
          continue;
        }

        break;
      }
    }

    console.log(`[Gemini] Trying fallback candidate model...`);
  }

  throw lastError;
}

interface ConversionPayload {
  fileBase64?: string;
  mimeType?: string;
  fileName?: string;
  customInstructions?: string;
  extractionMode?: string;
  parseNumbers?: boolean;
}

interface ConversionJob {
  id: string;
  status: 'processing' | 'done' | 'error';
  step: string;
  createdAt: number;
  result?: {
    success: true;
    data: any;
    fileName: string;
    modelUsed: string;
  };
  error?: string;
  httpStatus?: number;
}

const conversionJobs = new Map<string, ConversionJob>();

function cleanupOldJobs() {
  const cutoff = Date.now() - 30 * 60 * 1000;
  for (const [id, job] of conversionJobs.entries()) {
    if (job.createdAt < cutoff) {
      conversionJobs.delete(id);
    }
  }
}

function mapConversionError(error: any): { httpStatus: number; error: string; details: string } {
  const status = Number(error?.status || error?.code || 0);
  const rawError = `${error?.message || ''} ${status || ''}`;
  let httpStatus = Number(error?.httpStatus || 500);
  let userFriendly =
    error?.userMessage || 'حدث خطأ أثناء معالجة الملف. يرجى المحاولة مرة أخرى.';

  if (error?.userMessage) {
    return { httpStatus, error: userFriendly, details: rawError.slice(0, 400) };
  }

  if (
    status === 503 ||
    rawError.includes('503') ||
    rawError.includes('high demand') ||
    rawError.includes('UNAVAILABLE') ||
    rawError.includes('ECONNRESET') ||
    rawError.includes('fetch failed')
  ) {
    httpStatus = 503;
    userFriendly =
      'تشهد خوادم الذكاء الاصطناعي ضغطاً مؤقتاً في هذه اللحظة. يرجى النقر على زر "إعادة المحاولة الآن" للتحويل فوراً.';
  } else if (
    status === 429 ||
    rawError.includes('429') ||
    rawError.includes('RESOURCE_EXHAUSTED')
  ) {
    httpStatus = 429;
    userFriendly =
      'تم الوصول للحد الأقصى للطلبات مؤقتاً. يرجى الانتظار بضع ثوانٍ ثم الضغط على "إعادة المحاولة الآن".';
  } else if (
    status === 403 ||
    rawError.includes('403') ||
    rawError.includes('PERMISSION_DENIED') ||
    rawError.includes('API_KEY_INVALID')
  ) {
    httpStatus = 403;
    userFriendly =
      'يرجى التحقق من صحة مفتاح GEMINI_API_KEY في لوحة الإعدادات (Settings > Secrets).';
  } else if (status === 400 || rawError.includes('INVALID_ARGUMENT')) {
    httpStatus = 400;
    userFriendly =
      'تعذر قراءة الملف المرفق. تأكد من أن الملف بصيغة PDF أو صورة صالحة وغير تالفة أو محمية بكلمة مرور.';
  } else if (
    status === 504 ||
    rawError.includes('DEADLINE_EXCEEDED') ||
    rawError.includes('MODEL_TIMEOUT') ||
    rawError.includes('MODEL_STALL_TIMEOUT')
  ) {
    httpStatus = 504;
    userFriendly =
      'استغرقت معالجة الملف وقتاً طويلاً (قد يكون عدد الصفحات كبيراً). يرجى الضغط على "إعادة المحاولة الآن" أو تقسيم الملف.';
  } else if (status === 404 || rawError.includes('NOT_FOUND')) {
    httpStatus = 502;
    userFriendly = 'تعذر الوصول إلى نموذج التحويل المطلوب حالياً. يرجى المحاولة مرة أخرى.';
  }

  return {
    httpStatus,
    error: userFriendly,
    details: rawError.slice(0, 400),
  };
}

async function runPdfConversionPipeline(
  payload: ConversionPayload,
  onProgress?: (step: string) => void
) {
  const {
    fileBase64,
    mimeType = 'application/pdf',
    fileName = 'document.pdf',
    customInstructions = '',
    extractionMode = 'auto',
    parseNumbers = true,
  } = payload;

  if (!fileBase64) {
    const err: any = new Error('NO_FILE');
    err.httpStatus = 400;
    err.userMessage = 'لم يتم توفير ملف للتحويل. يرجى اختيار ملف PDF أو صورة جدول.';
    throw err;
  }

  if (!process.env.GEMINI_API_KEY) {
    const err: any = new Error('MISSING_KEY');
    err.httpStatus = 500;
    err.userMessage =
      'مفتاح GEMINI_API_KEY غير موجود في إعدادات الخادم. يرجى التأكد من ضبط المفتاح في لوحة التحكم (Settings > Secrets).';
    throw err;
  }

  const normalizedMimeType = normalizeMimeType(mimeType, fileName, fileBase64);
  const allowedMimeTypes = ['application/pdf', 'image/png', 'image/jpeg', 'image/webp'];
  if (!allowedMimeTypes.includes(normalizedMimeType)) {
    const err: any = new Error('UNSUPPORTED_MIME');
    err.httpStatus = 400;
    err.userMessage = `نوع الملف (${normalizedMimeType}) غير مدعوم. يرجى رفع ملف PDF أو صورة PNG/JPEG/WEBP.`;
    throw err;
  }

  const cleanBase64 = fileBase64.replace(/^data:[^;]+;base64,/, '');

  const systemInstruction = `أنت خبير فائق الدقة والاحترافية في استخراج الجداول من ملفات الـ PDF والصور وتحويلها إلى جداول Excel متكاملة.
تعليمات شمولية حاسمة:
1. قاعدة الشمولية التامة (100% Complete Extraction - Zero Missing Formulas):
   - يجب إلزامياً استخراج كافة الصفحات وكافة الجداول والتركيبات الموجودة في ملف الـ PDF من أول صفحة حتى آخر صفحة دون أي حذف أو اختصار أو تجزئة.
   - إذا كان المستند يحتوي على 5 أو 10 أو 20 أو 50 تركيبة موزعة على صفحات متعددة (مثل تقارير الخلطات والتشغيل Total Mixing Report / Consumption List):
     يجب استخراج كل تركيبة وكل خامة وصنف تابع لها حتى آخر بند في الملف.
   - ممنوع التوقف في منتصف الملف، وممنوع الاكتفاء بالصفحة الأولى أو ببعض التركيبات فقط، بل يجب استخراج التقرير كاملاً بنسبة 100%.
   - ضع جميع التركيبات وخاماتها كاملة في ورقة عمل رئيسية شاملة موحدة (Sheet Name: 'Mixing_Report') دون تجزئة ناقصة.
2. قاعدة جوهرية وإلزامية لاسم التركيبة ورقم الـ Supledger أمام كل صنف:
   في تقارير الخلطات والإنتاج (مثل ConsumptionByFinishedProductList)، تظهر التركيبات كترويسة قسم بهذا الشكل:
   'FN 19-9 A/2 Code111   (FN 19-9 A/2 Code111) 11336'
   'FN 19-9 A/7           (FN 19-9 A/7) 11308'
   'GR 28-6 A/8 (25KG)    (GR 28-6 A/8 (25KG)) 1504526'
   'Tilapia 13-10 S (30%) A/3   (Tilapia 13-10 S (30%) A/3) 908140'
   وتحتها أصناف وخامات التشغيل التابعة لها.
   - أضف عموداً أول باسم 'Finished_Product' لاسم التركيبة (شاملاً أي وصف مثل '(25KG)' أو 'Code111' بدون الرقم الأخير)، وكرّر اسم التركيبة أمام كل صنف وخامة في نفس الصف!
   - أضف عموداً باسم 'Supledger' واستخرج فيه إلزامياً الرقم التسلسلي/التشغيلي الذي يظهر في نهاية ترويسة التركيبة (مثل: 11336 أو 11308 أو 1504526 أو 908140) أمام كل صنف تابع لتلك التركيبة. ممنوع وضع '25KG' أو 'Mono' في عمود Supledger.
3. قاعدة فصل الأكواد: إذا كان اسم الصنف أو المكون يحتوي على كود رقمي تعريفي (مثل 'Anti ox ( 1472106 )' أو '(Bile Acid 200g) 907948' أو 'Corn (11342)'), افصل الكود دائماً في عمود مستقل بذاته يسمى 'Ingredient_Code'، واجعل اسم الصنف خالصاً في عمود 'Ingredient_Name'.
4. قاعدة عمود 'نوع الصنف' (إلزامية للخلطات والأصناف):
   أضف دائماً عموداً اسمه 'نوع الصنف' بجانب اسم التركيبة بحيث تتحدد قيمته لكل صف بناءً على اسم التركيبة:
   - إذا كان اسم التركيبة يحتوي على كلمة 'Tilapia' -> نوع الصنف: 'سمك'
   - إذا كان اسم التركيبة يحتوي على كلمة 'Duck' -> نوع الصنف: 'بط'
   - إذا كان اسم التركيبة يحتوي على أي من الرموز 'GR' أو 'FN' أو 'ST' أو 'SS' -> نوع الصنف: 'تسمين'
   - غير ذلك (أي تركيبة أخرى) -> نوع الصنف: 'اليف'
5. التمييز اللغوي واتجاه الجدول:
   - إذا كان المستند باللغة الإنجليزية (مثل تقارير الخلط والمصانع Mixing Reports)، اجعل isRtl = false.
   - إذا كان باللغة العربية، اجعل isRtl = true.
6. حافظ على سلامة الأرقام، النسب، التواريخ، والأكواد دون أي حذف أو اختصار.
7. لا تتأخر أو تضيف أي نصوص زائدة خارج الـ JSON المطلوب.`;

  const isMixingReport =
    fileName.toLowerCase().includes('mixing') ||
    fileName.toLowerCase().includes('consumption') ||
    fileName.toLowerCase().includes('خلط') ||
    fileName.toLowerCase().includes('علف') ||
    extractionMode === 'mixing_reports';

  const userPromptText = `استخرج جميع جداول وتركيبات هذا المستند (${fileName}) بالكامل بنسبة 100% دون حذف أو إسقاط أي تركيبة أو خامة من كافة صفحات الملف.
${isMixingReport ? 'تنبيه خاص لتقارير الخلطات: المستند يحتوي على عدة تركيبات متتالية موزعة على الصفحات. تأكد إلزامياً من المرور على جميع الصفحات واستخراج 100% من كافة التركيبات، وكتابة اسم التركيبة في عمود Finished_Product، واستخراج الرقم التسلسلي/التشغيلي (Supledger مثل 11336 أو 11308 أو 1504526) الظاهر بجانب اسم التركيبة في عمود مستقل باسم Supledger أمام كل صف تابع لها.' : ''}
تأكد من فصل الأكواد والأرقام التعريفية للأصناف في عمود خاص يسمى Ingredient_Code.
تأكد من إضافة وتعبئة عمود 'نوع الصنف' وفقاً للقواعد:
- اسم التركيبة فيه Tilapia -> 'سمك'
- اسم التركيبة فيه Duck -> 'بط'
- اسم التركيبة فيه GR أو FN أو ST أو SS -> 'تسمين'
- غير ذلك -> 'اليف'
نوع الاستخراج: ${extractionMode}
${customInstructions ? `ملاحظات إضافية: ${customInstructions}` : ''}
${parseNumbers ? 'تأكد من نظافة الأرقام لتناسب حسابات إكسل.' : ''}`;

  const filePart = {
    inlineData: {
      mimeType: normalizedMimeType,
      data: cleanBase64,
    },
  };

  const textPart = {
    text: userPromptText,
  };

  const { rawText, modelUsed } = await generateWithFallback(
    filePart,
    textPart,
    systemInstruction,
    onProgress
  );

  onProgress?.('جاري تنظيم الصفوف وفصل الأكواد وتصنيف الأنواع...');

  if (!rawText.trim()) {
    const err: any = new Error('EMPTY_MODEL_TEXT');
    err.httpStatus = 502;
    err.userMessage =
      'لم يُرجع نموذج الذكاء الاصطناعي بيانات لهذا الملف. تأكد من وضوح الملف وأعد المحاولة.';
    throw err;
  }

  let parsedData: any;
  let wasRecovered = false;
  try {
    const parsedResult = parseGeminiJson(rawText);
    parsedData = parsedResult.data;
    wasRecovered = parsedResult.wasRecovered;
  } catch {
    console.error('Failed to parse Gemini JSON output (length=' + rawText.length + '):', rawText.slice(0, 500));
    const err: any = new Error('INVALID_MODEL_JSON');
    err.httpStatus = 502;
    err.userMessage =
      'تعذر قراءة استجابة الذكاء الاصطناعي (قد يكون المستند ضخماً جداً). يرجى تقسيم الملف إلى أجزاء أصغر أو إعادة المحاولة.';
    throw err;
  }

  if (!parsedData || !Array.isArray(parsedData.sheets) || parsedData.sheets.length === 0) {
    const err: any = new Error('NO_SHEETS_FOUND');
    err.httpStatus = 422;
    err.userMessage = 'لم يتم العثور على جداول بيانات واضحة داخل هذا الملف.';
    throw err;
  }

  if (wasRecovered) {
    parsedData.notes = Array.isArray(parsedData.notes) ? parsedData.notes : [];
    parsedData.notes.unshift(
      'تنبيه: الملف كبير جداً؛ تم استرجاع الصفوف المكتملة بنجاح. يُفضّل تقسيم الملفات الضخمة لمزيد من الدقة.'
    );
  }

  parsedData.sheets = parsedData.sheets.map((rawSheet: any, idx: number) => {
    let currentSheet: SheetData = sanitizeSheetData(rawSheet, idx);
    try {
      const formulaRes = propagateFormulaNames(currentSheet, parsedData);
      currentSheet = formulaRes.sheet;

      const splitRes = autoSplitCodeColumn(currentSheet);
      if (splitRes) {
        currentSheet = splitRes.newSheet;
      }

      const classifiedRes = ensureItemTypeColumn(currentSheet, parsedData);
      currentSheet = classifiedRes.sheet;
    } catch (postErr) {
      console.warn(`[PostProcess] Non-fatal warning on sheet ${idx}:`, postErr);
    }
    return currentSheet;
  });

  return {
    success: true as const,
    data: parsedData,
    fileName,
    modelUsed,
  };
}

// Asynchronous start endpoint (immune to reverse-proxy idle request timeouts)
app.post('/api/convert-pdf/start', (req, res) => {
  cleanupOldJobs();
  const body = req.body || {};

  if (!body.fileBase64) {
    return res.status(400).json({ error: 'لم يتم توفير ملف للتحويل. يرجى اختيار ملف PDF أو صورة جدول.' });
  }
  if (!process.env.GEMINI_API_KEY) {
    return res.status(500).json({
      error: 'مفتاح GEMINI_API_KEY غير موجود في إعدادات الخادم. يرجى التأكد من ضبط المفتاح في لوحة التحكم (Settings > Secrets).',
    });
  }

  const jobId = `job_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
  const job: ConversionJob = {
    id: jobId,
    status: 'processing',
    step: 'جاري قراءة وتجهيز المستند للتحويل...',
    createdAt: Date.now(),
  };
  conversionJobs.set(jobId, job);

  // Fire background processing without holding the HTTP socket open
  void (async () => {
    try {
      const result = await runPdfConversionPipeline(body, (step) => {
        job.step = step;
      });
      job.result = result;
      job.status = 'done';
      job.step = 'اكتمل استخراج البيانات بنجاح!';
    } catch (error: any) {
      console.error(`[Job ${jobId}] Conversion error:`, error);
      const mapped = mapConversionError(error);
      job.status = 'error';
      job.httpStatus = mapped.httpStatus;
      job.error = mapped.error;
    }
  })();

  return res.status(202).json({ jobId, status: 'processing', step: job.step });
});

// Status polling endpoint
app.get('/api/convert-pdf/status/:jobId', (req, res) => {
  const job = conversionJobs.get(req.params.jobId);
  if (!job) {
    return res.status(404).json({ error: 'انتهت صلاحية جلسة التحويل أو لم يتم العثور عليها. يرجى إعادة المحاولة.' });
  }

  if (job.status === 'processing') {
    return res.json({ status: 'processing', step: job.step });
  }

  if (job.status === 'error') {
    return res.json({
      status: 'error',
      httpStatus: job.httpStatus || 500,
      error: job.error || 'حدث خطأ أثناء معالجة الملف. يرجى المحاولة مرة أخرى.',
    });
  }

  return res.json({
    status: 'done',
    step: job.step,
    result: job.result,
  });
});

// Direct synchronous conversion endpoint (kept for direct API compatibility)
app.post('/api/convert-pdf', async (req, res) => {
  try {
    const result = await runPdfConversionPipeline(req.body || {});
    return res.json(result);
  } catch (error: any) {
    console.error('Error during PDF conversion:', error);
    const mapped = mapConversionError(error);
    return res.status(mapped.httpStatus).json({
      error: mapped.error,
      details: mapped.details,
    });
  }
});

// Python FastAPI (Camelot + pdfplumber) optional backend bridge
const PYTHON_PDF_API_URL = process.env.PYTHON_PDF_API_URL || 'http://127.0.0.1:8000';

app.get('/api/python-pdf/health', async (_req, res) => {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 2500);
    const response = await fetch(`${PYTHON_PDF_API_URL}/health`, {
      signal: controller.signal,
    }).finally(() => clearTimeout(timer));

    if (!response.ok) {
      return res.json({ available: false, camelot: false, pdfplumber: false });
    }
    const data = await response.json();
    const isPdfApi =
      typeof data?.camelot === 'boolean' && typeof data?.pdfplumber === 'boolean';
    if (!isPdfApi) {
      return res.json({ available: false, camelot: false, pdfplumber: false });
    }
    return res.json({ available: true, ...data });
  } catch {
    return res.json({ available: false, camelot: false, pdfplumber: false });
  }
});

app.post('/api/python-pdf/extract', async (req, res) => {
  try {
    const {
      fileBase64,
      fileName = 'document.pdf',
      pages = 'all',
      flavor = 'auto',
      mergeMultipage = true,
    } = req.body || {};

    if (!fileBase64 || typeof fileBase64 !== 'string') {
      return res.status(400).json({ error: 'لم يتم إرسال بيانات ملف PDF.' });
    }

    const cleanB64 = fileBase64.replace(/^data:[^;]+;base64,/, '');
    const buffer = Buffer.from(cleanB64, 'base64');

    const formData = new FormData();
    formData.append('file', new Blob([buffer], { type: 'application/pdf' }), fileName);
    formData.append('pages', String(pages || 'all'));
    formData.append('flavor', String(flavor || 'auto'));
    formData.append('merge_multipage', mergeMultipage ? 'true' : 'false');

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 120000);
    const response = await fetch(`${PYTHON_PDF_API_URL}/extract`, {
      method: 'POST',
      body: formData,
      signal: controller.signal,
    }).finally(() => clearTimeout(timer));

    const jsonData = await response.json().catch(() => null);
    if (!response.ok) {
      return res.status(response.status).json({
        error: jsonData?.detail || 'فشل استخراج الجداول عبر خادم Python FastAPI.',
      });
    }

    return res.json(jsonData);
  } catch {
    return res.status(503).json({
      error:
        'خدمة Python FastAPI (Camelot + pdfplumber) غير مشغلة حالياً على المنفذ 8000. يمكنك استخدام المحرك المحلي المدمج أو تشغيل الخدمة عبر backend/pdf_extractor_api.py.',
    });
  }
});

// Global Express error handler to ensure JSON response on all errors (e.g. payload too large)
app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
  console.error('[Express Error Handler]', err?.message || err);
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Content-Type', 'application/json; charset=utf-8');

  if (err?.type === 'entity.too.large' || err?.status === 413) {
    return res.status(413).json({
      error: 'حجم الملف كبير جداً بالنسبة للخادم. يرجى اختيار ملف أصغر حجماً (أقل من 20 ميجابايت) أو تقليل عدد الصفحات.',
    });
  }

  if (err instanceof SyntaxError && 'body' in err) {
    return res.status(400).json({
      error: 'حدث خطأ في قراءة بيانات الملف المرفق. يرجى اختيار الملف مرة أخرى والمحاولة مجدداً.',
    });
  }

  return res.status(err?.status || 500).json({
    error: err?.message || 'حدث خطأ غير متوقع في الخادم أثناء معالجة الطلب.',
  });
});

// Ensure unmatched /api/* endpoints always return JSON 404 rather than Vite/static HTML
app.all('/api/*', (_req, res) => {
  return res.status(404).json({ error: 'مسار API غير موجود' });
});

// Start dev Vite or serve static assets
async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: {
        middlewareMode: true,
        hmr: process.env.DISABLE_HMR !== 'true',
      },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.resolve(__dirname, 'dist')));
    app.get('*', (_req, res) => {
      res.sendFile(path.resolve(__dirname, 'dist', 'index.html'));
    });
  }

  const server = app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running at http://0.0.0.0:${PORT}`);
  });
  server.setTimeout(360000); // 6 minutes timeout for processing
  server.keepAliveTimeout = 65000;
  server.headersTimeout = 66000;
}

startServer();
