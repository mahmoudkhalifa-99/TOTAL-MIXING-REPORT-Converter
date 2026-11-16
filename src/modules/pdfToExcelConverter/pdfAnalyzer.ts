import * as pdfjsLib from 'pdfjs-dist';
import pdfjsWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { PageInspection, PageKind, PdfDocumentInspection } from './types';

if (typeof pdfjsWorkerUrl === 'string') {
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorkerUrl;
}

/**
 * Inspects every page of a PDF locally using pdfjs-dist to classify whether
 * pages contain selectable vector text, scanned images, or both.
 */
export async function inspectPdfDocument(
  file: File,
  onProgress?: (percent: number, stage: string) => void
): Promise<{
  inspection: PdfDocumentInspection;
  pdfDoc: pdfjsLib.PDFDocumentProxy;
}> {
  onProgress?.(5, 'جاري فتح ملف الـ PDF وفحص الصفحات...');
  const buffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    verbosity: 0,
  });
  const pdfDoc = await loadingTask.promise;
  const totalPages = pdfDoc.numPages;

  if (totalPages === 0) {
    throw new Error('ملف الـ PDF لا يحتوي على أي صفحات قابلة للقراءة.');
  }

  const pages: PageInspection[] = [];
  let textPagesCount = 0;
  let scannedPagesCount = 0;
  let mixedPagesCount = 0;

  for (let p = 1; p <= totalPages; p++) {
    const pct = Math.min(95, Math.round((p / totalPages) * 90));
    onProgress?.(pct, `جاري فحص الصفحة ${p} من ${totalPages} (نصي / ممسوح ضوئياً)...`);

    const page = await pdfDoc.getPage(p);
    const viewport = page.getViewport({ scale: 1.0 });
    const textContent = await page.getTextContent();

    let textItemCount = 0;
    let charCount = 0;
    for (const item of textContent.items as any[]) {
      if (item && typeof item.str === 'string') {
        const cleaned = item.str.trim();
        if (cleaned) {
          textItemCount++;
          charCount += cleaned.length;
        }
      }
    }

    let imageCount = 0;
    try {
      const opList = await page.getOperatorList();
      const fnArray = opList.fnArray || [];
      for (const fn of fnArray) {
        if (
          fn === pdfjsLib.OPS.paintImageXObject ||
          fn === pdfjsLib.OPS.paintInlineImageXObject ||
          fn === pdfjsLib.OPS.paintImageMaskXObject
        ) {
          imageCount++;
        }
      }
    } catch {
      // Operator list optional on protected streams
    }

    let kind: PageKind = 'text';
    if (charCount < 15) {
      kind = 'scanned';
      scannedPagesCount++;
    } else if (imageCount > 0 && charCount < 140) {
      kind = 'mixed';
      mixedPagesCount++;
    } else {
      kind = 'text';
      textPagesCount++;
    }

    pages.push({
      pageNumber: p,
      kind,
      textItemCount,
      charCount,
      imageCount,
      width: Math.round(viewport.width),
      height: Math.round(viewport.height),
    });
  }

  let documentKind: 'text' | 'scanned' | 'mixed' = 'text';
  if (scannedPagesCount === totalPages) {
    documentKind = 'scanned';
  } else if (scannedPagesCount > 0 || mixedPagesCount > 0) {
    documentKind = 'mixed';
  }

  onProgress?.(100, 'اكتمل فحص بنية الملف.');

  return {
    inspection: {
      fileName: file.name,
      fileSize: file.size,
      totalPages,
      documentKind,
      textPagesCount,
      scannedPagesCount,
      mixedPagesCount,
      pages,
    },
    pdfDoc,
  };
}

/**
 * Parses a human-friendly page range string such as "1-3, 5, 8-10" into sorted unique page numbers.
 */
export function parsePageRangeSpec(spec: string, totalPages: number): number[] {
  const trimmed = (spec || '').trim().toLowerCase();
  if (!trimmed || trimmed === 'all' || trimmed === 'الكل') {
    return Array.from({ length: totalPages }, (_, i) => i + 1);
  }

  const chosen = new Set<number>();
  const parts = trimmed.split(/[,،\s]+/);
  for (const part of parts) {
    if (!part) continue;
    if (part.includes('-')) {
      const [rawStart, rawEnd] = part.split('-', 2);
      const s = parseInt(rawStart.trim(), 10);
      const e = parseInt(rawEnd.trim(), 10);
      if (!isNaN(s) && !isNaN(e)) {
        const lo = Math.max(1, Math.min(s, e));
        const hi = Math.min(totalPages, Math.max(s, e));
        for (let p = lo; p <= hi; p++) chosen.add(p);
      }
    } else {
      const p = parseInt(part, 10);
      if (!isNaN(p) && p >= 1 && p <= totalPages) {
        chosen.add(p);
      }
    }
  }

  return Array.from(chosen).sort((a, b) => a - b);
}
