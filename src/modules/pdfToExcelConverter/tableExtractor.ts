import type * as pdfjsLib from 'pdfjs-dist';
import { ExtractedTable } from './types';

export interface PositionedToken {
  text: string;
  xStart: number;
  xEnd: number;
  y: number;
  height: number;
}

export interface AlignedRow {
  y: number;
  height: number;
  chunks: { text: string; xStart: number; xEnd: number; xCenter: number }[];
}

interface ColumnCluster {
  xStart: number;
  xEnd: number;
  xCenter: number;
}

const ARABIC_RE = /[\u0600-\u06FF]/g;
const LATIN_RE = /[A-Za-z]/g;

function normalizeTokenText(raw: string): string {
  // Preserve exact numbers, dates, leading zeros, decimals, commas, currency signs verbatim
  return (raw || '')
    .normalize('NFKC')
    .replace(/\u200f|\u200e|\ufeff/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function isPureNumericOrFinancial(text: string): boolean {
  const s = (text || '').trim();
  if (!s) return false;
  return (
    /^[+-]?[\d,]+(?:\.\d+)?%?$/.test(s) ||
    /^\([+-]?[\d,]+(?:\.\d+)?%?\)$/.test(s) ||
    /^[$€£¥]?\s*[\d,.]+\s*[$€£¥]?$/.test(s) ||
    /^[-—]+$/.test(s)
  );
}

function hasDecimalPart(text: string): boolean {
  const s = (text || '').trim();
  return /^[+-]?[\d,]+\.\d+$/.test(s) || /^\([+-]?[\d,]+\.\d+\)$/.test(s);
}

/**
 * Groups raw positioned spans from a single PDF page (or OCR TSV tokens) into
 * visual horizontal rows, keeping distinct columns separate (never merging across column gaps).
 */
export function groupTokensIntoAlignedRows(tokens: PositionedToken[]): AlignedRow[] {
  if (tokens.length === 0) return [];

  const heights = tokens
    .map((t) => t.height)
    .filter((h) => h > 2 && h < 60)
    .sort((a, b) => a - b);
  const medianH = heights.length > 0 ? heights[Math.floor(heights.length / 2)] : 10;
  const yTol = Math.max(3, medianH * 0.45);

  // PDF text tokens already use bottom-up Y (higher Y = top of page); sort Y descending
  const sorted = [...tokens].sort((a, b) => b.y - a.y || a.xStart - b.xStart);

  const rowBuckets: { y: number; heights: number[]; items: PositionedToken[] }[] = [];
  for (const tok of sorted) {
    let target: (typeof rowBuckets)[number] | null = null;
    let bestDiff = Infinity;
    for (let i = rowBuckets.length - 1; i >= Math.max(0, rowBuckets.length - 8); i--) {
      const diff = Math.abs(rowBuckets[i].y - tok.y);
      if (diff <= yTol && diff < bestDiff) {
        bestDiff = diff;
        target = rowBuckets[i];
      }
    }
    if (target) {
      target.items.push(tok);
      target.heights.push(tok.height);
      const count = target.items.length;
      target.y = (target.y * (count - 1) + tok.y) / count;
    } else {
      rowBuckets.push({ y: tok.y, heights: [tok.height], items: [tok] });
    }
  }

  const rows: AlignedRow[] = [];

  for (const bucket of rowBuckets) {
    const lineH =
      bucket.heights.reduce((s, v) => s + v, 0) / Math.max(1, bucket.heights.length);
    const items = bucket.items.sort((a, b) => a.xStart - b.xStart);

    const chunks: AlignedRow['chunks'] = [];
    for (const item of items) {
      const text = normalizeTokenText(item.text);
      if (!text) continue;

      const xStart = item.xStart;
      const xEnd =
        item.xEnd > item.xStart ? item.xEnd : item.xStart + Math.max(4, lineH * 0.4);

      if (chunks.length === 0) {
        chunks.push({ text, xStart, xEnd, xCenter: (xStart + xEnd) / 2 });
        continue;
      }

      const prev = chunks[chunks.length - 1];
      const gap = xStart - prev.xEnd;
      const bothDecimals = hasDecimalPart(prev.text) && hasDecimalPart(text);
      const bothNumbers =
        isPureNumericOrFinancial(prev.text) && isPureNumericOrFinancial(text);
      const isSplitNumFragment =
        (/[,\.]$/.test(prev.text) || /^[\.,]\d+$/.test(text) || /^%$/.test(text)) &&
        gap <= Math.max(2.5, lineH * 0.25);
      const mergeThreshold = isSplitNumFragment
        ? Math.max(2.5, lineH * 0.25)
        : bothDecimals
        ? -0.5
        : bothNumbers
        ? 0.5
        : isPureNumericOrFinancial(prev.text) !== isPureNumericOrFinancial(text)
        ? Math.max(3.5, lineH * 0.45)
        : Math.max(6.5, lineH * 0.85);

      if (gap <= mergeThreshold) {
        const needsSpace = gap > 1.0 && !prev.text.endsWith(' ') && !text.startsWith(' ');
        prev.text = `${prev.text}${needsSpace ? ' ' : ''}${text}`.trim();
        prev.xEnd = Math.max(prev.xEnd, xEnd);
        prev.xCenter = (prev.xStart + prev.xEnd) / 2;
      } else {
        chunks.push({ text, xStart, xEnd, xCenter: (xStart + xEnd) / 2 });
      }
    }

    if (chunks.length > 0) {
      rows.push({ y: bucket.y, height: lineH, chunks });
    }
  }

  return rows;
}

/**
 * Splits a page's visual rows into distinct table segments when large vertical gaps
 * or single-cell section titles separate multiple tables on the same page.
 */
function segmentPageRowsIntoTableBlocks(rows: AlignedRow[]): AlignedRow[][] {
  if (rows.length === 0) return [];

  // Keep multi-cell rows (>= 2 chunks) and single-cell rows closely adjacent to them
  const blocks: AlignedRow[][] = [];
  let current: AlignedRow[] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const prev = current.length > 0 ? current[current.length - 1] : null;
    const medianH = Math.max(8, row.height);

    // Split into a new table block if there is a large vertical gap (> 3.8x line height)
    if (prev && prev.y - row.y > medianH * 3.8 && current.length >= 2) {
      const multiCount = current.filter((r) => r.chunks.length >= 2).length;
      if (multiCount >= 2) {
        blocks.push(current);
        current = [];
      }
    }

    current.push(row);
  }

  if (current.length > 0) {
    blocks.push(current);
  }

  // Filter out pure single-column prose blocks unless no multi-column tables were detected at all
  const tabularBlocks = blocks.filter(
    (b) => b.filter((r) => r.chunks.length >= 2).length >= 2
  );
  if (tabularBlocks.length > 0) {
    return tabularBlocks;
  }

  return blocks.filter((b) => b.length >= 1);
}

/**
 * Clusters column X intervals across a table block while enforcing the same-row
 * separation invariant (two chunks on the same row can NEVER merge into one column).
 */
function clusterBlockColumns(block: AlignedRow[]): ColumnCluster[] {
  const clusters: ColumnCluster[] = [];

  // Seed clusters from the earliest dense row (>= 2 chunks, preferably a non-numeric header row)
  const seedRow =
    block.find(
      (r) =>
        r.chunks.length >= 2 &&
        r.chunks.filter((c) => !isPureNumericOrFinancial(c.text)).length >= 2
    ) || block.reduce((best, r) => (r.chunks.length > best.chunks.length ? r : best), block[0]);

  if (seedRow) {
    for (const c of seedRow.chunks) {
      clusters.push({ xStart: c.xStart, xEnd: c.xEnd, xCenter: c.xCenter });
    }
  }

  for (const row of block) {
    const usedInRow = new Set<number>();

    for (const chunk of row.chunks) {
      let bestIdx = -1;
      let bestScore = -Infinity;

      const isNum = isPureNumericOrFinancial(chunk.text);
      const pad = isNum ? 1.5 : 4;

      for (let i = 0; i < clusters.length; i++) {
        if (usedInRow.has(i)) continue;
        const cl = clusters[i];
        const overlap =
          Math.min(cl.xEnd + pad, chunk.xEnd) - Math.max(cl.xStart - pad, chunk.xStart);
        const dist = Math.abs(cl.xCenter - chunk.xCenter);
        const endDist = Math.abs(cl.xEnd - chunk.xEnd);
        if (overlap > 0) {
          const score = overlap * 3 - (isNum ? endDist * 0.35 + dist * 0.15 : dist * 0.25);
          if (score > bestScore) {
            bestScore = score;
            bestIdx = i;
          }
        } else if (isNum ? endDist <= 6 && dist <= 14 : dist <= 18) {
          const score = 10 - (isNum ? endDist : dist);
          if (score > bestScore) {
            bestScore = score;
            bestIdx = i;
          }
        }
      }

      if (bestIdx === -1) {
        clusters.push({
          xStart: chunk.xStart,
          xEnd: chunk.xEnd,
          xCenter: chunk.xCenter,
        });
        usedInRow.add(clusters.length - 1);
      } else {
        const cl = clusters[bestIdx];
        cl.xStart = Math.min(cl.xStart, chunk.xStart);
        cl.xEnd = Math.max(cl.xEnd, chunk.xEnd);
        cl.xCenter = (cl.xStart + cl.xEnd) / 2;
        usedInRow.add(bestIdx);
      }
    }
  }

  return clusters.sort((a, b) => a.xCenter - b.xCenter);
}

/**
 * Converts a table block + column clusters into a rectangular grid preserving
 * every empty cell and keeping distinct columns separate.
 */
export function buildTableFromTokens(
  tokens: PositionedToken[],
  pageNumber: number,
  tableIndexOffset: number,
  engine: ExtractedTable['engine'],
  baseConfidence?: number
): ExtractedTable[] {
  const alignedRows = groupTokensIntoAlignedRows(tokens);
  const blocks = segmentPageRowsIntoTableBlocks(alignedRows);
  const results: ExtractedTable[] = [];

  for (let bIdx = 0; bIdx < blocks.length; bIdx++) {
    const block = blocks[bIdx];
    if (block.length === 0) continue;

    // Trim leading/trailing page-number-only lines like "Page 1 of 5"
    const cleanedBlock = block.filter((row) => {
      const joined = row.chunks.map((c) => c.text).join(' ').trim();
      if (!joined) return false;
      if (/^(?:page|p\.?|صفحة)\s*\d+(?:\s*(?:of|\/|من)\s*\d+)?$/i.test(joined)) {
        return false;
      }
      return true;
    });

    if (cleanedBlock.length === 0) continue;

    const clusters = clusterBlockColumns(cleanedBlock);
    if (clusters.length === 0) continue;

    const gridRows: string[][] = cleanedBlock.map((row) => {
      const cells = new Array(clusters.length).fill('');
      const usedCols = new Set<number>();

      for (const chunk of row.chunks) {
        let bestIdx = -1;
        let bestScore = -Infinity;

        const isNum = isPureNumericOrFinancial(chunk.text);
        const pad = isNum ? 1.5 : 4;

        for (let i = 0; i < clusters.length; i++) {
          if (usedCols.has(i)) continue;
          const cl = clusters[i];
          const overlap =
            Math.min(cl.xEnd + pad, chunk.xEnd) - Math.max(cl.xStart - pad, chunk.xStart);
          const dist = Math.abs(cl.xCenter - chunk.xCenter);
          const endDist = Math.abs(cl.xEnd - chunk.xEnd);
          const score =
            overlap > 0
              ? overlap * 3 - (isNum ? endDist * 0.35 + dist * 0.15 : dist * 0.2)
              : -(isNum ? endDist * 0.6 + dist * 0.4 : dist);
          if (score > bestScore) {
            bestScore = score;
            bestIdx = i;
          }
        }

        if (bestIdx !== -1) {
          usedCols.add(bestIdx);
          // Never merge data from different columns; only place into distinct free slot
          cells[bestIdx] = chunk.text;
        }
      }

      return cells;
    });

    // Detect RTL dominance from text characters
    const joinedSample = gridRows.slice(0, 40).flat().join(' ');
    const arabicCount = (joinedSample.match(ARABIC_RE) || []).length;
    const latinCount = (joinedSample.match(LATIN_RE) || []).length;
    const isRtl = arabicCount > latinCount;

    const orientedRows = isRtl ? gridRows.map((r) => [...r].reverse()) : gridRows;

    // Find best header row index (first row with >= 2 non-numeric cells, defaulting to row 0)
    let headerIdx = orientedRows.findIndex(
      (r) =>
        r.filter((c) => c.trim() !== '').length >= Math.min(2, clusters.length) &&
        r.filter((c) => c.trim() !== '' && !isPureNumericOrFinancial(c)).length >=
          Math.min(2, clusters.length)
    );
    if (headerIdx === -1) headerIdx = 0;

    // Keep any rows BEFORE headerIdx too if they contain valid data, OR slice cleanly
    const headers =
      orientedRows.length > 1
        ? orientedRows[headerIdx].map((h, i) => h || `عمود_${i + 1}`)
        : clusters.map((_, i) => `عمود_${i + 1}`);

    const dataRows =
      orientedRows.length > 1
        ? [...orientedRows.slice(0, headerIdx), ...orientedRows.slice(headerIdx + 1)]
        : orientedRows;

    // Preserve ALL rows & columns — never remove rows because of empty cells!
    const diagnostics = evaluateTableDiagnostics(headers, dataRows, baseConfidence);

    results.push({
      id: `tbl_p${pageNumber}_${tableIndexOffset + results.length + 1}`,
      name:
        blocks.length > 1
          ? `جدول ${tableIndexOffset + results.length + 1} (صفحة ${pageNumber})`
          : `جدول صفحة ${pageNumber}`,
      pages: [pageNumber],
      isRtl,
      headers,
      rows: dataRows,
      engine,
      accuracyScore: diagnostics.accuracyScore,
      emptyCellRatio: diagnostics.emptyCellRatio,
      warnings: diagnostics.warnings,
    });
  }

  return results;
}

export function evaluateTableDiagnostics(
  headers: string[],
  rows: string[][],
  baseConfidence?: number
): { accuracyScore: number; emptyCellRatio: number; warnings: string[] } {
  const colCount = Math.max(headers.length, 1);
  if (rows.length === 0) {
    return {
      accuracyScore: baseConfidence ? Math.round(baseConfidence * 0.7) : 55,
      emptyCellRatio: 1,
      warnings: ['تم استخراج صف العناوين فقط دون صفوف بيانات إضافية.'],
    };
  }

  let emptyCells = 0;
  let totalCells = rows.length * colCount;
  let sparseRows = 0;

  for (const r of rows) {
    let populated = 0;
    for (let c = 0; c < colCount; c++) {
      if ((r[c] ?? '').trim() === '') {
        emptyCells++;
      } else {
        populated++;
      }
    }
    if (populated < Math.max(1, Math.floor(colCount * 0.4))) {
      sparseRows++;
    }
  }

  const emptyCellRatio = emptyCells / Math.max(1, totalCells);
  const alignmentConsistency = Math.max(0, 100 - (sparseRows / rows.length) * 65);
  const colBonus = colCount >= 2 ? 100 : 65;

  let score = alignmentConsistency * 0.75 + colBonus * 0.25;
  if (typeof baseConfidence === 'number') {
    score = score * 0.55 + baseConfidence * 0.45;
  }
  const accuracyScore = Math.max(0, Math.min(100, Math.round(score)));

  const warnings: string[] = [];
  if (accuracyScore < 80) {
    warnings.push(
      `تم رصد انخفاض محتمل في دقة محاذاة هذا الجدول (${accuracyScore}%)؛ يُنصح بمراجعة الخلايا قبل التصدير.`
    );
  }
  if (colCount === 1) {
    warnings.push(
      'تم اكتشاف عمود واحد فقط؛ قد يكون الجدول غير محدد الحدود أو يحتاج إلى خادم Camelot/pdfplumber أو تعديل يدوي.'
    );
  }
  if (emptyCellRatio > 0.45) {
    warnings.push(
      `نسبة الخلايا الفارغة ${Math.round(emptyCellRatio * 100)}% (تم الحفاظ على كافة الصفوف والأعمدة دون حذف).`
    );
  }

  return { accuracyScore, emptyCellRatio: Number(emptyCellRatio.toFixed(3)), warnings };
}

function splitSpanTokens(tok: PositionedToken): PositionedToken[] {
  const hasWideGap = /(\t|\s{2,})/.test(tok.text);
  const hasDecimalNextToNum =
    /(?:[+-]?[\d,]+\.\d+\s+[+-]?[\d,(]+|[+-]?[\d,.)]+\s+[+-]?[\d,]+\.\d+)/.test(tok.text);

  if (!hasWideGap && !hasDecimalNextToNum) return [tok];
  const totalLen = tok.text.length;
  if (totalLen === 0) return [];
  const totalWidth = tok.xEnd > tok.xStart ? tok.xEnd - tok.xStart : totalLen * Math.max(4, tok.height * 0.42);
  const charW = totalWidth / totalLen;

  const out: PositionedToken[] = [];
  const wideRegex = /\S+(?: \S+)*/g;
  let match: RegExpExecArray | null;
  while ((match = wideRegex.exec(tok.text)) !== null) {
    const baseIdx = match.index;
    const segment = match[0];
    if (/(?:[+-]?[\d,]+\.\d+\s+[+-]?[\d,(]+|[+-]?[\d,.)]+\s+[+-]?[\d,]+\.\d+)/.test(segment)) {
      const subParts = segment.split(
        /(?<=[+-]?[\d,]+\.\d+\)?)\s+(?=[+-]?\(?[\d,]+(?:\.\d+)?\)?%?)|(?<=^|\s)(?<=[+-]?\(?[\d,]+(?:\.\d+)?\)?)\s+(?=[+-]?[\d,]+\.\d+)/
      );
      let cursor = 0;
      for (const sub of subParts) {
        if (!sub) continue;
        const foundAt = segment.indexOf(sub, cursor);
        const relIdx = foundAt !== -1 ? foundAt : cursor;
        cursor = relIdx + sub.length;
        const sX = tok.xStart + (baseIdx + relIdx) * charW;
        out.push({
          text: sub,
          xStart: sX,
          xEnd: sX + sub.length * charW,
          y: tok.y,
          height: tok.height,
        });
      }
    } else {
      const sX = tok.xStart + baseIdx * charW;
      out.push({
        text: segment,
        xStart: sX,
        xEnd: sX + segment.length * charW,
        y: tok.y,
        height: tok.height,
      });
    }
  }
  return out.length > 0 ? out : [tok];
}

/**
 * Extracts tables directly from a selectable-text PDF page using pdfjs-dist coordinates.
 */
export async function extractTablesFromPdfPage(
  pdfDoc: pdfjsLib.PDFDocumentProxy,
  pageNumber: number,
  tableIndexOffset: number
): Promise<ExtractedTable[]> {
  const pdfjsRuntime = await import('pdfjs-dist');
  const page = await pdfDoc.getPage(pageNumber);
  const viewport = page.getViewport({ scale: 1 });
  const content = await page.getTextContent();

  const tokens: PositionedToken[] = [];
  for (const item of content.items as any[]) {
    if (!item || typeof item.str !== 'string') continue;
    const text = item.str.trim();
    if (!text) continue;
    const tr = Array.isArray(item.transform) ? item.transform : [1, 0, 0, 10, 0, 0];
    const tx = pdfjsRuntime.Util.transform(viewport.transform, tr);

    // Skip vertical margin watermarks (90-degree rotated text relative to viewport)
    if (Math.abs(tx[1]) > Math.abs(tx[0]) * 1.5 && Math.abs(tx[1]) > 0.5) {
      continue;
    }

    const h =
      Math.hypot(tx[2] || 0, tx[3] || 0) ||
      Math.hypot(tx[0] || 0, tx[1] || 0) ||
      item.height ||
      10;
    const xStart = tx[4] || 0;
    const w = item.width > 0 ? Math.abs(item.width) : text.length * Math.max(4, h * 0.42);
    // Keep bottom-up Y convention (higher Y = top of page) consistent with OCR TSV & groupTokensIntoAlignedRows
    const bottomUpY = viewport.height - (tx[5] || 0);

    const rawTok: PositionedToken = {
      text: item.str,
      xStart,
      xEnd: xStart + w,
      y: bottomUpY,
      height: h,
    };
    tokens.push(...splitSpanTokens(rawTok));
  }

  if (tokens.length === 0) return [];
  return buildTableFromTokens(tokens, pageNumber, tableIndexOffset, 'pdfjs-coords');
}

/**
 * Merges consecutive tables across pages when they share the same column count
 * (and matching headers, skipping duplicate repeated page headers cleanly).
 */
export function mergeMultiPageExtractedTables(tables: ExtractedTable[]): ExtractedTable[] {
  if (tables.length <= 1) return tables;

  const merged: ExtractedTable[] = [];

  for (const tbl of tables) {
    if (merged.length === 0) {
      merged.push({ ...tbl, pages: [...tbl.pages], rows: tbl.rows.map((r) => [...r]) });
      continue;
    }

    const prev = merged[merged.length - 1];
    const sameColumnCount = prev.headers.length === tbl.headers.length;
    const nextPage = tbl.pages[0] <= prev.pages[prev.pages.length - 1] + 1;
    const normPrev = prev.headers.map((h) => h.toLowerCase().trim()).join('|');
    const normCurr = tbl.headers.map((h) => h.toLowerCase().trim()).join('|');
    const sameHeaders = normPrev === normCurr;

    if (sameColumnCount && nextPage && (sameHeaders || prev.headers.length >= 3)) {
      if (!sameHeaders) {
        // First row of next page was actually data continuation, preserve it!
        prev.rows.push([...tbl.headers]);
      }
      for (const r of tbl.rows) {
        prev.rows.push([...r]);
      }
      for (const p of tbl.pages) {
        if (!prev.pages.includes(p)) prev.pages.push(p);
      }
      prev.name =
        prev.pages.length > 1
          ? `جدول ممتد (صفحات ${prev.pages[0]}-${prev.pages[prev.pages.length - 1]})`
          : prev.name;

      const updatedDiag = evaluateTableDiagnostics(
        prev.headers,
        prev.rows,
        Math.round((prev.accuracyScore + tbl.accuracyScore) / 2)
      );
      prev.accuracyScore = updatedDiag.accuracyScore;
      prev.emptyCellRatio = updatedDiag.emptyCellRatio;
      prev.warnings = Array.from(new Set([...prev.warnings, ...updatedDiag.warnings]));
    } else {
      merged.push({ ...tbl, pages: [...tbl.pages], rows: tbl.rows.map((r) => [...r]) });
    }
  }

  return merged;
}
