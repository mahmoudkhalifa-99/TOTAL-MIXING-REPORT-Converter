import type * as pdfjsLib from 'pdfjs-dist';
import { createWorker } from 'tesseract.js';
import { ExtractedTable, OcrLanguageOption } from './types';
import { buildTableFromTokens, PositionedToken } from './tableExtractor';

/**
 * Parses Tesseract TSV output into top-down inverted PositionedToken items
 * compatible with the coordinate-based column alignment engine.
 */
function parseTsvToTokens(tsv: string, canvasHeight: number): {
  tokens: PositionedToken[];
  meanConfidence: number;
} {
  const tokens: PositionedToken[] = [];
  const confidences: number[] = [];

  const lines = (tsv || '').split(/\r?\n/);
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) continue;
    const cols = line.split('\t');
    if (cols.length < 12) continue;

    const level = Number(cols[0]);
    const left = Number(cols[6]);
    const top = Number(cols[7]);
    const width = Number(cols[8]);
    const height = Number(cols[9]);
    const conf = Number(cols[10]);
    const text = cols.slice(11).join('\t').trim();

    if (level !== 5 || !text) continue;
    if (Number.isFinite(conf) && conf >= 0) {
      confidences.push(conf);
    }

    // Invert canvas Y (top-down) into PDF bottom-up Y so alignment helper works identically
    const yBottomUp = canvasHeight - (top + height * 0.5);
    tokens.push({
      text,
      xStart: left,
      xEnd: left + Math.max(4, width),
      y: yBottomUp,
      height: Math.max(8, height),
    });
  }

  const meanConfidence =
    confidences.length > 0
      ? confidences.reduce((sum, c) => sum + c, 0) / confidences.length
      : 0;

  return { tokens, meanConfidence };
}

/**
 * Runs optional open-source Tesseract.js OCR (Arabic + English) on a rendered PDF page
 * and extracts structured tables from word bounding boxes without guessing values.
 */
export async function extractTablesWithOcr(
  pdfDoc: pdfjsLib.PDFDocumentProxy,
  pageNumbers: number[],
  lang: OcrLanguageOption,
  tableIndexOffset: number,
  onProgress?: (percent: number, stage: string) => void
): Promise<ExtractedTable[]> {
  if (pageNumbers.length === 0) return [];

  onProgress?.(5, `جاري تهيئة محرك OCR للغة (${lang})...`);
  const worker = await createWorker(lang);

  const tables: ExtractedTable[] = [];

  try {
    for (let idx = 0; idx < pageNumbers.length; idx++) {
      const pageNum = pageNumbers[idx];
      const basePct = Math.round((idx / pageNumbers.length) * 90);
      onProgress?.(basePct, `جاري تشغيل OCR على الصفحة ${pageNum} (${idx + 1}/${pageNumbers.length})...`);

      const page = await pdfDoc.getPage(pageNum);
      const viewport = page.getViewport({ scale: 2.0 });

      const canvas = document.createElement('canvas');
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      const ctx = canvas.getContext('2d');
      if (!ctx) {
        throw new Error('تعذر تهيئة Canvas لعرض الصفحة لمحرك OCR.');
      }

      await page.render({ canvasContext: ctx, viewport, canvas }).promise;

      const result = await worker.recognize(canvas, undefined, { tsv: true, text: true });
      const tsvStr = result?.data?.tsv || '';
      const { tokens, meanConfidence } = parseTsvToTokens(tsvStr, canvas.height);

      if (tokens.length > 0) {
        const pageTables = buildTableFromTokens(
          tokens,
          pageNum,
          tableIndexOffset + tables.length,
          'tesseract-ocr',
          meanConfidence
        );

        pageTables.forEach((tbl) => {
          if (meanConfidence > 0 && meanConfidence < 75) {
            tbl.warnings.unshift(
              `ثقة التعرف الضوئي OCR منخفضة نسبياً (${Math.round(meanConfidence)}%) في الصفحة ${pageNum} — يُرجى مراجعة الأرقام والنصوص المستخرجة.`
            );
          }
          tables.push(tbl);
        });
      } else if (result?.data?.text?.trim()) {
        // Fallback line split if TSV returned empty
        const textLines = result.data.text
          .split(/\r?\n/)
          .map((l) => l.trim())
          .filter(Boolean);
        const rows = textLines.map((l) => l.split(/\s{2,}|\t/).map((c) => c.trim()));
        const maxCols = Math.max(...rows.map((r) => r.length), 1);
        const padded = rows.map((r) => [...r, ...new Array(maxCols - r.length).fill('')]);
        tables.push({
          id: `ocr_p${pageNum}_${tables.length + 1}`,
          name: `جدول OCR (صفحة ${pageNum})`,
          pages: [pageNum],
          isRtl: lang.includes('ara'),
          headers: padded[0] || ['النص'],
          rows: padded.slice(1),
          engine: 'tesseract-ocr',
          accuracyScore: Math.round(result.data.confidence || 60),
          emptyCellRatio: 0,
          warnings: [
            'تم استخراج النص من الصفحة الممسوحة عبر OCR؛ يُنصح بمراجعة محاذاة الأعمدة قبل التصدير.',
          ],
        });
      }
    }

    onProgress?.(100, 'اكتمل التعرف الضوئي OCR بنجاح.');
    return tables;
  } finally {
    await worker.terminate();
  }
}
