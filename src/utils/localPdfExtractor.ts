import * as pdfjsLib from 'pdfjs-dist';
import pdfjsWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import * as XLSX from 'xlsx';
import { ConversionOptions, ExtractedDocument, SheetData } from '../types/converter';
import {
  autoSplitCodeColumn,
  consolidateDuplicateCodeColumns,
  extractCodeFromValue,
} from './codeSplitter';
import {
  extractSupledgerAndCleanFormula,
  propagateFormulaNames,
} from './formulaPropagator';
import { ensureItemTypeColumn } from './itemClassifier';

if (typeof pdfjsWorkerUrl === 'string') {
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorkerUrl;
}

interface RawSpan {
  str: string;
  x: number;
  y: number;
  width: number;
  height: number;
  pageNum: number;
}

interface CellChunk {
  text: string;
  xStart: number;
  xEnd: number;
  xCenter: number;
}

interface PdfLine {
  pageNum: number;
  y: number;
  height: number;
  chunks: CellChunk[];
  rawText: string;
}

interface ColumnCluster {
  xStart: number;
  xEnd: number;
  xCenter: number;
  sumCenter: number;
  count: number;
}

const ARABIC_INDIC_MAP: Record<string, string> = {
  '٠': '0',
  '١': '1',
  '٢': '2',
  '٣': '3',
  '٤': '4',
  '٥': '5',
  '٦': '6',
  '٧': '7',
  '٨': '8',
  '٩': '9',
  '٫': '.',
  '٬': ',',
};

function normalizeText(input: string, parseNumbers: boolean): string {
  let out = (input || '')
    .normalize('NFKC')
    .replace(/\u200f|\u200e|\ufeff/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  if (parseNumbers && out) {
    out = out.replace(/[٠-٩٫٬]/g, (ch) => ARABIC_INDIC_MAP[ch] ?? ch);
  }
  return out;
}

function isPureNumericLike(text: string): boolean {
  const s = (text || '').trim();
  if (!s) return false;
  return (
    /^[+-]?[\d,]+(?:\.\d+)?%?$/.test(s) ||
    /^\([+-]?[\d,]+(?:\.\d+)?%?\)$/.test(s) ||
    /^[-—*/]+$/.test(s)
  );
}

function hasDecimalWeight(text: string): boolean {
  const s = (text || '').trim();
  return /^[+-]?[\d,]+\.\d+$/.test(s) || /^\([+-]?[\d,]+\.\d+\)$/.test(s);
}

function isPageChromeLine(line: PdfLine): boolean {
  const s = line.rawText.trim();
  if (!s) return true;
  if (/^(?:page|p\.?|صفحة)\s*\d+(?:\s*(?:of|\/|من)\s*\d+)?$/i.test(s)) return true;

  const hasDecimal = line.chunks.some((c) => hasDecimalWeight(c.text));
  if (!hasDecimal) {
    const isPageToken = (t: string) =>
      /^(?:page|p\.?|صفحة)\s*\d+(?:\s*(?:of|\/|من)\s*\d+)?$/i.test(t.trim());
    const isDateTimeToken = (t: string) =>
      /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AP]M)?)?$/i.test(
        t.trim()
      );
    const isPrintedByToken = (t: string) =>
      /^(?:printed\s*(?:on|by|at)|run\s*date|report\s*date|تاريخ\s*الطباعة)/i.test(t.trim());

    if (line.chunks.some((c) => isPageToken(c.text))) {
      return true;
    }
    if (
      line.chunks.length > 0 &&
      line.chunks.every(
        (c) => isDateTimeToken(c.text) || isPageToken(c.text) || isPrintedByToken(c.text)
      )
    ) {
      return true;
    }
  }

  if (/^(?:printed\s*(?:on|by|at)|run\s*date|report\s*date|تاريخ\s*الطباعة)/i.test(s)) {
    return true;
  }
  return false;
}

const HEADER_KEYWORDS = [
  'finished',
  'product',
  'formula',
  'ident',
  'code',
  'ingredient',
  'material',
  'item',
  'name',
  'description',
  'desc',
  'target',
  'actual',
  'diff',
  'variance',
  'qty',
  'quantity',
  'weight',
  'unit',
  'price',
  'amount',
  'total',
  'date',
  'supledger',
  'subledger',
  'batch',
  'اسم',
  'إسم',
  'كود',
  'رمز',
  'الصنف',
  'الخامة',
  'المكون',
  'التركيبة',
  'الخلطة',
  'الكمية',
  'الوزن',
  'المستهدف',
  'الفعلي',
  'الفرق',
  'السعر',
  'الاجمالي',
  'الإجمالي',
  'المبلغ',
  'الوحدة',
  'التاريخ',
  'البيان',
  'الوصف',
  'م',
];

function scoreHeaderLine(line: PdfLine): number {
  if (line.chunks.length < 2) return -10;
  let score = 0;
  let numericCells = 0;

  for (const chunk of line.chunks) {
    const lower = chunk.text.toLowerCase();
    if (isPureNumericLike(chunk.text)) {
      numericCells++;
      score -= 3;
      continue;
    }
    for (const kw of HEADER_KEYWORDS) {
      if (lower === kw || lower.includes(kw)) {
        score += 3;
        break;
      }
    }
  }

  if (numericCells === line.chunks.length) return -10;
  if (numericCells > 0) score -= numericCells * 2;
  if (line.chunks.length >= 3) score += 2;
  return score;
}

function normalizeHeaderForCompare(line: PdfLine): string {
  return line.chunks
    .map((c) => c.text.toLowerCase().replace(/[^a-z0-9\u0600-\u06ff]+/g, ''))
    .filter(Boolean)
    .join('|');
}

/**
 * A wrapped continuation of an ITEM name (e.g. "Broiler 1.5 kg" after "Mineral Premix"),
 * not a formula banner: ends with a bare weight/unit and has no formula code pattern.
 */
function looksLikeItemNameFragment(text: string): boolean {
  const t = (text || '').trim();
  if (!t) return false;
  const endsWithUnit = /(?:^|\s)\d+(?:[.,]\d+)?\s*(?:kg|kgs|كجم|كيلو|gm|g|gram|ton|طن|%)\s*$/i.test(t);
  const hasFormulaCode =
    /\d+\s*-\s*\d+/.test(t) ||
    /Code\s*\d+/i.test(t) || /(?:^|[\s\-_/(])(FN|GR|ST|SS)\s*\d+/i.test(t);
  return endsWithUnit && !hasFormulaCode;
}

/**
 * Identifies Crystal Reports / ERP Formula Section Header lines (e.g. "FN 19-9 A/7 (FN 19-9 A/7) 11308")
 * WITHOUT ever swallowing actual ingredient rows (even if ingredient names contain "Starter", "Grower", "Duck", "Pet", etc.).
 */
function isLikelyFormulaGroupLine(line: PdfLine): boolean {
  const joined = line.chunks.map((c) => c.text).join(' ').trim();
  if (!joined) return false;
  if (/إجمالي|الاجمالي|المجموع|\b(?:sub\s*-?\s*total|subtotal|grand\s*total|formula\s*total|total)\b/i.test(joined)) {
    return false;
  }

  const nonNumericChunks = line.chunks.filter((c) => !isPureNumericLike(c.text));
  const nonNumericText = nonNumericChunks.map((c) => c.text).join(' ').trim();
  const decimalNumbers = line.chunks.filter((c) => hasDecimalWeight(c.text));

  // CRITICAL: Any row containing decimal weights (Target / Actual kg) is a DATA or TOTAL row,
  // NEVER a formula banner (even if ingredient name repeats inside parentheses like "DDG (DDG ) 11350").
  if (decimalNumbers.length >= 1) {
    return false;
  }

  // Wrapped item-name fragment (e.g. "Broiler 1.5 kg") is part of an ingredient name, never a formula banner
  if (looksLikeItemNameFragment(nonNumericText) || looksLikeItemNameFragment(joined)) {
    return false;
  }

  // Reject raw ingredient names (such as "Duck Premix GROWING", "Rejected Pets", "Fish premix", "DDG (DDG ) 11350")
  const looksLikeIngredientName =
    /\b(?:premix|permix|permex|rejected|digest|powder|meal|bran|acid|chloride|choline|vitamin|mineral|methionin|lysine|threonine|limestone|monocalcium|mono calcium|salt|water|oil|soya|corn|ddg|wheat|screenings|mycofix|anti\s*ox|bicarbonate|phytase|toxin)\b/i.test(
      nonNumericText
    );
  if (looksLikeIngredientName && extractSupledgerAndCleanFormula(joined).supledger === '') {
    return false;
  }

  // Explicit Crystal Reports duplicated parenthetical formula group banner:
  // e.g. "FN 19-9 A/7 (FN 19-9 A/7) 11308" or "GR 28-6 A/8 (25KG) (GR 28-6 A/8 (25KG)) 1504526"
  const hasDuplicatedBanner = /^(.+?)\s*\(\s*\1\s*\)\s*(?:\d{4,})?$/i.test(nonNumericText);
  if (hasDuplicatedBanner && !looksLikeIngredientName) {
    return true;
  }

  // If the row has NO decimal weights and <= 3 chunks matching formula naming:
  // Supports both 2-chunk ("Tilapia 13-10 S (30%) | Tilapia 13-10 S (30%) A/3- 908140")
  // and 3-chunk banners when supledger splits into a separate chunk.
  if (nonNumericChunks.length <= 3 && line.chunks.length <= 3 && !looksLikeIngredientName) {
    // A banner must carry a real formula code (Code111 / FN-GR-ST-SS 19-9 / "13-10" style numbers).
    // Generic words alone (broiler, starter, grower, pet...) are NOT enough: they appear in wrapped item names.
    const hasFormulaCodePattern =
      /Code\s*\d+/i.test(nonNumericText) ||
      /(?:^|[\s\-_/(])(FN|GR|ST|SS)\s*\d+(?:[\s\-_/]+\d+)?/i.test(nonNumericText) ||
      /\d+\s*-\s*\d+/.test(nonNumericText);

    if (hasFormulaCodePattern) {
      return true;
    }
  }

  return false;
}

/**
 * A plain group title that carries NO formula code, e.g. "Cleaning Aqua".
 * It is only trusted when it comes right after the table header or a group subtotal
 * (see `expectBanner` in the mixing loop), because wrapped item names look the same in isolation.
 */
function isPlainGroupTitleCandidate(line: PdfLine): boolean {
  const joined = line.chunks.map((c) => c.text).join(' ').trim();
  if (!joined || joined.length < 3 || !/[A-Za-z\u0600-\u06FF]/.test(joined)) return false;
  if (line.chunks.some((c) => hasDecimalWeight(c.text))) return false;
  if (line.chunks.length > 3) return false;
  if (/إجمالي|الاجمالي|المجموع|\b(?:sub\s*-?\s*total|subtotal|grand\s*total|total)\b/i.test(joined)) return false;
  // page footer / header chrome such as "23/11/2026 08:03:46 Page 5 of 5"
  if (/\bpage\b|صفحة|\d{1,2}:\d{2}(?::\d{2})?|\d{1,2}\/\d{1,2}\/\d{2,4}/i.test(joined)) return false;
  if (looksLikeItemNameFragment(joined)) return false;
  // an item code in the line means it is an ingredient row, not a title
  if (/\b\d{4,}\b/.test(joined)) return false;
  return true;
}

function isLikelySubtotalRow(cells: string[]): boolean {
  const nonEmpty = cells.map((c) => c.trim()).filter(Boolean);
  if (nonEmpty.length === 0) return true;

  const firstText = nonEmpty[0].toLowerCase();
  if (
    /^(?:إجمالي|الاجمالي|المجموع|مجموع|sub\s*-?\s*total|subtotal|grand\s*total|formula\s*total|sum\s+of)\b/i.test(
      firstText
    ) ||
    /^total(?:\s*[:\-]|\s*\[[^\]]+\]|\s+(?:for|of|formula|mixing|production|consumption|fn|gr|st|ss|duck|tilapia|broiler|pet|item|weight|qty|kg|ton)\b|$)/i.test(
      firstText
    )
  ) {
    return true;
  }

  // Check if any cell is a total label or ends with "Total"
  const hasTotalLabelCell = nonEmpty.some(
    (c) =>
      /^(?:total|sub\s*-?\s*total|subtotal|grand\s*total|sum|formula\s*total|إجمالي|الاجمالي|المجموع)(?:\s*[:\-]|\s*\[[^\]]+\])?$/i.test(
        c
      ) ||
      /\b(?:sub\s*-?\s*total|subtotal|grand\s*total|formula\s*total|group\s*total)\b/i.test(c) ||
      /\s+total\s*[:\-]?$/i.test(c)
  );
  if (hasTotalLabelCell) return true;

  // Row with a single non-numeric text cell starting with "Total ..." plus numeric totals
  const textCells = nonEmpty.filter((c) => !isPureNumericLike(c));
  if (textCells.length === 0) {
    return true;
  }
  if (textCells.length === 1 && /^total\b/i.test(textCells[0])) {
    return true;
  }

  return false;
}

function splitMultiSpaceSpan(sp: RawSpan): RawSpan[] {
  // Split single PDF spans on wide gaps (>=2 spaces or tab) AND separate space-delimited decimal weights
  // so adjacent table columns like "Corn  5,036.40 5,036.40" or "94,284.60 94,284.60 0.00" never fuse.
  const hasWideGap = /(\t|\s{2,})/.test(sp.str);
  const hasDecimalNextToNum =
    /(?:[+-]?[\d,]+\.\d+\s+[+-]?[\d,(]+|[+-]?[\d,.)]+\s+[+-]?[\d,]+\.\d+)/.test(sp.str);

  if (!hasWideGap && !hasDecimalNextToNum) {
    return [sp];
  }

  const totalLen = sp.str.length;
  if (totalLen === 0) return [];
  const totalWidth = sp.width > 0 ? sp.width : totalLen * Math.max(4, sp.height * 0.45);
  const charWidth = totalWidth / totalLen;

  // First slice on wide gaps (\t or 2+ spaces), then refine each slice on decimal-to-number boundaries
  const parts: RawSpan[] = [];
  const wideRegex = /\S+(?: \S+)*/g;
  let match: RegExpExecArray | null;
  while ((match = wideRegex.exec(sp.str)) !== null) {
    const baseOffset = match.index;
    const segment = match[0];

    if (/(?:[+-]?[\d,]+\.\d+\s+[+-]?[\d,(]+|[+-]?[\d,.)]+\s+[+-]?[\d,]+\.\d+)/.test(segment)) {
      // Split between adjacent numeric/decimal tokens while preserving multi-word names like "Soya 46%" or "Hi-phos GT ( 20000 )"
      const subParts = segment.split(
        /(?<=[+-]?[\d,]+\.\d+\)?)\s+(?=[+-]?\(?[\d,]+(?:\.\d+)?\)?%?)|(?<=^|\s)(?<=[+-]?\(?[\d,]+(?:\.\d+)?\)?)\s+(?=[+-]?[\d,]+\.\d+)/
      );
      let cursor = 0;
      for (const sub of subParts) {
        if (!sub) continue;
        const foundAt = segment.indexOf(sub, cursor);
        const relIdx = foundAt !== -1 ? foundAt : cursor;
        cursor = relIdx + sub.length;
        parts.push({
          str: sub,
          x: sp.x + (baseOffset + relIdx) * charWidth,
          y: sp.y,
          width: sub.length * charWidth,
          height: sp.height,
          pageNum: sp.pageNum,
        });
      }
    } else {
      parts.push({
        str: segment,
        x: sp.x + baseOffset * charWidth,
        y: sp.y,
        width: segment.length * charWidth,
        height: sp.height,
        pageNum: sp.pageNum,
      });
    }
  }
  return parts.length > 0 ? parts : [sp];
}

function groupSpansIntoLines(rawSpans: RawSpan[], parseNumbers: boolean): PdfLine[] {
  const spans = rawSpans.flatMap(splitMultiSpaceSpan);
  if (spans.length === 0) return [];

  const heights = spans
    .map((s) => s.height)
    .filter((h) => h > 2 && h < 40)
    .sort((a, b) => a - b);
  const medianHeight = heights.length > 0 ? heights[Math.floor(heights.length / 2)] : 10;
  const yTol = Math.max(2.5, medianHeight * 0.48);

  // Viewport coordinates have y=0 at TOP increasing downward -> sort top-to-bottom ascending y
  const sortedByY = [...spans].sort((a, b) => a.y - b.y || a.x - b.x);

  const lineGroups: { y: number; heights: number[]; spans: RawSpan[] }[] = [];
  for (const sp of sortedByY) {
    let matchedGroup: (typeof lineGroups)[number] | null = null;
    let bestDiff = Infinity;
    for (let i = lineGroups.length - 1; i >= Math.max(0, lineGroups.length - 8); i--) {
      const diff = Math.abs(lineGroups[i].y - sp.y);
      if (diff <= yTol && diff < bestDiff) {
        bestDiff = diff;
        matchedGroup = lineGroups[i];
      }
    }
    if (matchedGroup) {
      matchedGroup.spans.push(sp);
      matchedGroup.heights.push(sp.height);
      const count = matchedGroup.spans.length;
      matchedGroup.y = (matchedGroup.y * (count - 1) + sp.y) / count;
    } else {
      lineGroups.push({ y: sp.y, heights: [sp.height], spans: [sp] });
    }
  }

  const lines: PdfLine[] = [];

  for (const grp of lineGroups) {
    const rowSpans = grp.spans.sort((a, b) => a.x - b.x);
    const lineHeight =
      grp.heights.reduce((s, v) => s + v, 0) / Math.max(1, grp.heights.length);

    const chunks: CellChunk[] = [];
    for (const sp of rowSpans) {
      const text = normalizeText(sp.str, parseNumbers);
      if (!text) continue;

      const spWidth =
        sp.width > 0 ? sp.width : text.length * Math.max(3.5, lineHeight * 0.42);
      const spStart = sp.x;
      const spEnd = sp.x + spWidth;

      if (chunks.length === 0) {
        chunks.push({
          text,
          xStart: spStart,
          xEnd: spEnd,
          xCenter: (spStart + spEnd) / 2,
        });
        continue;
      }

      const prev = chunks[chunks.length - 1];
      const gap = spStart - prev.xEnd;

      // Keep adjacent numbers and Code+Text columns strictly distinct while merging split numeric fragments or normal words
      const prevIsDec = hasDecimalWeight(prev.text);
      const currIsDec = hasDecimalWeight(text);
      const prevIsNum = isPureNumericLike(prev.text);
      const currIsNum = isPureNumericLike(text);
      const isSplitNumericFragment =
        (/[,\.]$/.test(prev.text) || /^[\.,]\d+$/.test(text) || /^%$/.test(text)) &&
        gap <= Math.max(2.5, lineHeight * 0.25);
      const mergeThreshold = isSplitNumericFragment
        ? Math.max(2.5, lineHeight * 0.25)
        : prevIsDec && currIsDec
        ? -0.5
        : prevIsNum && currIsNum
        ? 0.5
        : prevIsNum !== currIsNum
        ? Math.max(3.5, lineHeight * 0.45)
        : Math.max(9.0, lineHeight * 1.0);

      if (gap <= mergeThreshold) {
        const needsSpace = gap > 1.0 && !prev.text.endsWith(' ') && !text.startsWith(' ');
        prev.text = `${prev.text}${needsSpace ? ' ' : ''}${text}`.trim();
        prev.xEnd = Math.max(prev.xEnd, spEnd);
        prev.xCenter = (prev.xStart + prev.xEnd) / 2;
      } else {
        chunks.push({
          text,
          xStart: spStart,
          xEnd: spEnd,
          xCenter: (spStart + spEnd) / 2,
        });
      }
    }

    if (chunks.length > 0) {
      const rawText = chunks.map((c) => c.text).join(' ');
      lines.push({
        pageNum: rowSpans[0].pageNum,
        y: grp.y,
        height: lineHeight,
        chunks,
        rawText,
      });
    }
  }

  return lines;
}

/**
 * Clusters columns primarily from dense data rows so wide header cells or stray text never collapse
 * two distinct side-by-side data columns into one. Enforces the same-row co-occurrence invariant.
 */
function buildColumnClusters(
  dataLines: PdfLine[],
  headerLine: PdfLine | null
): ColumnCluster[] {
  const clusters: ColumnCluster[] = [];

  const denseRows = dataLines.filter((l) => l.chunks.length >= 2);
  const seedLines = denseRows.length > 0 ? denseRows : headerLine ? [headerLine] : dataLines;

  // Sort lines by descending chunk count so the fullest rows seed clean column centers first
  const sortedSeedLines = [...seedLines].sort((a, b) => b.chunks.length - a.chunks.length);

  const findClusterForChunk = (
    chunk: CellChunk,
    usedClusters: Set<number>
  ): number => {
    let bestIdx = -1;
    let bestScore = -Infinity;

    const chunkIsNum = isPureNumericLike(chunk.text);
    const pad = chunkIsNum ? 1.5 : 4;

    for (let i = 0; i < clusters.length; i++) {
      if (usedClusters.has(i)) continue;
      const cl = clusters[i];
      const overlap =
        Math.min(cl.xEnd + pad, chunk.xEnd) - Math.max(cl.xStart - pad, chunk.xStart);
      const centerDist = Math.abs(cl.xCenter - chunk.xCenter);
      const startDist = Math.abs(cl.xStart - chunk.xStart);
      const endDist = Math.abs(cl.xEnd - chunk.xEnd);
      const minEdgeDist = chunkIsNum
        ? Math.min(centerDist, endDist)
        : Math.min(centerDist, startDist, endDist);

      const closeEnough = chunkIsNum
        ? overlap > 0 || (endDist <= 6 && centerDist <= 14)
        : overlap > 0 || minEdgeDist <= 14;

      if (closeEnough) {
        const score = chunkIsNum
          ? overlap * 2.5 - endDist * 0.45 - centerDist * 0.25
          : overlap * 2 - centerDist * 0.35 - Math.min(startDist, endDist) * 0.25;
        if (score > bestScore) {
          bestScore = score;
          bestIdx = i;
        }
      }
    }

    return bestIdx;
  };

  for (const line of sortedSeedLines) {
    const usedClusters = new Set<number>();

    for (const chunk of line.chunks) {
      const idx = findClusterForChunk(chunk, usedClusters);

      if (idx === -1) {
        clusters.push({
          xStart: chunk.xStart,
          xEnd: chunk.xEnd,
          xCenter: chunk.xCenter,
          sumCenter: chunk.xCenter,
          count: 1,
        });
        usedClusters.add(clusters.length - 1);
      } else {
        const cl = clusters[idx];
        // Update gently without letting one huge outlier string balloon bounds across neighbors
        const chunkWidth = chunk.xEnd - chunk.xStart;
        const clWidth = cl.xEnd - cl.xStart;
        if (chunkWidth <= Math.max(clWidth * 2.0, 140)) {
          cl.xStart = Math.min(cl.xStart, chunk.xStart);
          cl.xEnd = Math.max(cl.xEnd, chunk.xEnd);
        }
        cl.count += 1;
        cl.sumCenter += chunk.xCenter;
        cl.xCenter = cl.sumCenter / cl.count;
        usedClusters.add(idx);
      }
    }
  }

  return clusters.sort((a, b) => a.xCenter - b.xCenter);
}

function assignLineToColumns(
  line: PdfLine,
  clusters: ColumnCluster[],
  isRtl: boolean
): string[] {
  const row = new Array(clusters.length).fill('');
  const used = new Set<number>();

  for (const chunk of line.chunks) {
    let bestIdx = -1;
    let bestScore = -Infinity;
    const chunkIsNum = isPureNumericLike(chunk.text);
    const pad = chunkIsNum ? 1.5 : 4;

    for (let i = 0; i < clusters.length; i++) {
      const cl = clusters[i];
      const overlap =
        Math.min(cl.xEnd + pad, chunk.xEnd) - Math.max(cl.xStart - pad, chunk.xStart);
      const centerDist = Math.abs(cl.xCenter - chunk.xCenter);
      const startDist = Math.abs(cl.xStart - chunk.xStart);
      const endDist = Math.abs(cl.xEnd - chunk.xEnd);
      const minDist = chunkIsNum
        ? Math.min(centerDist, endDist)
        : Math.min(centerDist, startDist, endDist);

      let score =
        overlap > 0
          ? overlap * 3 - (chunkIsNum ? endDist * 0.35 + centerDist * 0.15 : minDist * 0.25)
          : -(chunkIsNum ? endDist * 0.6 + centerDist * 0.4 : minDist);
      if (used.has(i)) score -= 600;
      if (score > bestScore) {
        bestScore = score;
        bestIdx = i;
      }
    }

    if (bestIdx !== -1) {
      used.add(bestIdx);
      row[bestIdx] = row[bestIdx] ? `${row[bestIdx]} ${chunk.text}` : chunk.text;
    }
  }

  return isRtl ? row.reverse() : row;
}

function looksLikeMixingReport(lines: PdfLine[], fileName: string): boolean {
  const sampleText = lines
    .slice(0, 100)
    .map((l) => l.rawText)
    .join(' ')
    .toLowerCase();
  const fLower = fileName.toLowerCase();

  if (
    fLower.includes('mixing') ||
    fLower.includes('consumption') ||
    fLower.includes('خلط') ||
    fLower.includes('علف') ||
    sampleText.includes('consumptionbyfinishedproduct') ||
    sampleText.includes('total mixing') ||
    sampleText.includes('mixing report')
  ) {
    return true;
  }

  const hasTargetActual =
    (sampleText.includes('target') && sampleText.includes('actual')) ||
    (sampleText.includes('المستهدف') && sampleText.includes('الفعلي'));
  const hasFormulaCodes =
    /(?:^|[\s(])(fn|gr|st|ss)\s*\d+/i.test(sampleText) ||
    sampleText.includes('tilapia') ||
    sampleText.includes('duck');

  return hasTargetActual || hasFormulaCodes;
}

function normalizeMixingHeaders(headers: string[]): string[] {
  return headers.map((h, i) => {
    const norm = (h || '')
      .toLowerCase()
      .replace(/[()._]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (norm === 'ident' || norm === 'code' || norm === 'item code' || norm === 'mat code' || norm === 'ingredient code') {
      return 'Ingredient_Code';
    }
    if (
      norm === 'name' ||
      norm === 'ingredient' ||
      norm === 'ingredient name' ||
      norm === 'material' ||
      norm === 'material name' ||
      norm === 'raw material' ||
      norm === 'item name' ||
      norm === 'description'
    ) {
      return 'Ingredient_Name';
    }
    if (/^(?:target|target kg|target qty|planned|plan qty|std qty)$/.test(norm)) {
      return 'Target_kg';
    }
    if (/^(?:actual|actual kg|actual qty|act qty|consumed|consumed kg|used|used qty|dispensed)$/.test(norm)) {
      return 'Actual_kg';
    }
    return h || `Column_${i + 1}`;
  });
}

function isPureJdeCodeToken(val: string): boolean {
  const s = (val || '').trim();
  if (!s) return false;
  if (/^[1-9]0{3,}$/.test(s)) return false;
  return /^[0-9]{4,}$/.test(s);
}

/**
 * Repairs PDF layouts where the Crystal Reports column header causes Short Ingredient Names
 * (e.g. "Corn", "Mono calcium", "Poultry meal feathers") to land under Ingredient_Code while
 * "(corn 7.5%) 11342" or bare JDE codes ("808819", "752042") land under Ingredient_Name.
 */
function repairSwappedCodeAndNameColumns(sheet: SheetData): SheetData {
  if (!sheet?.headers || !sheet?.rows || sheet.rows.length === 0) {
    return sheet;
  }

  const normHeaders = sheet.headers.map((h) => (h ?? '').toString().trim().toLowerCase());
  let codeCol = normHeaders.findIndex(
    (h) =>
      h === 'ingredient_code' ||
      h === 'ingredient code' ||
      h === 'كود الصنف' ||
      h === 'كود_الصنف' ||
      h === 'ident' ||
      h === 'code'
  );
  let nameCol = normHeaders.findIndex(
    (h) =>
      h === 'ingredient_name' ||
      h === 'ingredient name' ||
      h === 'اسم الصنف' ||
      h === 'إسم الصنف' ||
      h === 'اسم المكون' ||
      h === 'description'
  );

  // Fallback: if left column is unlabeled ("Column_1") holding short names and right column holds desc+codes
  if (codeCol === -1 && nameCol > 0) {
    const prevNorm = normHeaders[nameCol - 1] || '';
    if (prevNorm.startsWith('column_') || prevNorm === 'item' || prevNorm === 'material') {
      codeCol = nameCol - 1;
    }
  }

  if (codeCol === -1 || nameCol === -1 || codeCol === nameCol) {
    return sheet;
  }

  let codeTotal = 0;
  let codeLettersCount = 0;
  let codePureNumCount = 0;
  let nameTotal = 0;
  let nameExtractableCount = 0;

  for (const row of sheet.rows) {
    const cVal = (row[codeCol] ?? '').toString().trim();
    const nVal = (row[nameCol] ?? '').toString().trim();

    if (cVal && cVal !== '-' && cVal !== '—') {
      codeTotal++;
      if (isPureJdeCodeToken(cVal)) {
        codePureNumCount++;
      } else if (/[A-Za-z\u0600-\u06FF]/.test(cVal) && !extractCodeFromValue(cVal)) {
        codeLettersCount++;
      }
    }

    if (nVal && nVal !== '-' && nVal !== '—') {
      nameTotal++;
      if (isPureJdeCodeToken(nVal) || extractCodeFromValue(nVal) !== null) {
        nameExtractableCount++;
      }
    }
  }

  const looksSwapped =
    codeTotal > 0 &&
    nameTotal > 0 &&
    codeLettersCount / codeTotal >= 0.6 &&
    codePureNumCount / codeTotal < 0.3 &&
    nameExtractableCount / nameTotal >= 0.4;

  if (!looksSwapped) {
    return sheet;
  }

  const nextHeaders = [...sheet.headers];
  nextHeaders[codeCol] = 'Ingredient_Code';
  nextHeaders[nameCol] = 'Ingredient_Name';

  const nextRows = sheet.rows.map((row) => {
    const nextRow = [...row];
    const shortName = (nextRow[codeCol] ?? '').toString().trim();
    const descOrCode = (nextRow[nameCol] ?? '').toString().trim();

    let extractedCode = '';
    let fallbackCleanName = descOrCode;

    if (isPureJdeCodeToken(descOrCode)) {
      extractedCode = descOrCode;
      fallbackCleanName = '';
    } else {
      const parsed = extractCodeFromValue(descOrCode);
      if (parsed) {
        extractedCode = parsed.code;
        fallbackCleanName = parsed.cleanText;
      } else if (descOrCode.startsWith('(') && descOrCode.endsWith(')')) {
        fallbackCleanName = descOrCode.slice(1, -1).trim();
      }
    }

    nextRow[codeCol] = extractedCode;
    nextRow[nameCol] = shortName || fallbackCleanName;
    return nextRow;
  });

  return {
    ...sheet,
    headers: nextHeaders,
    rows: nextRows,
  };
}

function postProcessSheet(sheet: SheetData, doc: Partial<ExtractedDocument>): SheetData {
  let current = consolidateDuplicateCodeColumns(sheet);

  try {
    const formulaRes = propagateFormulaNames(current, doc);
    current = formulaRes.sheet;
  } catch {
    // ignore non-fatal
  }

  try {
    current = repairSwappedCodeAndNameColumns(current);
  } catch {
    // ignore non-fatal
  }

  try {
    const splitRes = autoSplitCodeColumn(current);
    if (splitRes) {
      current = splitRes.newSheet;
    }
  } catch {
    // ignore non-fatal
  }

  try {
    current = consolidateDuplicateCodeColumns(current);
    const classifiedRes = ensureItemTypeColumn(current, doc);
    current = classifiedRes.sheet;
  } catch {
    // ignore non-fatal
  }

  return current;
}

export async function extractDocumentFromPdfLocal(
  file: File,
  options: ConversionOptions,
  onProgress?: (step: string) => void
): Promise<ExtractedDocument> {
  onProgress?.('جاري قراءة صفحات ملف الـ PDF محلياً...');

  const buffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    verbosity: 0,
  });

  const pdf = await loadingTask.promise;
  const numPages = pdf.numPages;

  if (numPages === 0) {
    throw new Error('ملف الـ PDF فارغ ولا يحتوي على صفحات.');
  }

  const allPagesLines: PdfLine[][] = [];
  let totalTextChars = 0;

  for (let p = 1; p <= numPages; p++) {
    onProgress?.(`جاري استخراج الجداول والنصوص من الصفحة (${p} / ${numPages})...`);
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();

    const spans: RawSpan[] = [];
    for (const item of content.items as any[]) {
      if (!item || typeof item.str !== 'string') continue;
      const trimmed = item.str.trim();
      if (!trimmed) continue;
      totalTextChars += trimmed.length;

      const rawTr = Array.isArray(item.transform) ? item.transform : [1, 0, 0, 10, 0, 0];
      // Transform into visual upright viewport coordinates (handles /Rotate 90/180/270 & CropBox)
      const tx = pdfjsLib.Util.transform(viewport.transform, rawTr);

      // Skip vertical margin watermarks (90-degree rotated text relative to viewport)
      if (Math.abs(tx[1]) > Math.abs(tx[0]) * 1.5 && Math.abs(tx[1]) > 0.5) {
        continue;
      }

      const fontHeight =
        Math.hypot(tx[2] || 0, tx[3] || 0) ||
        Math.hypot(tx[0] || 0, tx[1] || 0) ||
        item.height ||
        10;

      spans.push({
        str: item.str,
        x: tx[4] || 0,
        y: tx[5] || 0,
        width: Math.abs(item.width || 0),
        height: fontHeight,
        pageNum: p,
      });
    }

    const pageLines = groupSpansIntoLines(spans, options.parseNumbers);
    allPagesLines.push(pageLines);
  }

  if (totalTextChars < 5) {
    throw new Error(
      'هذا الملف عبارة عن صورة ممسوحة ضوئياً بدون طبقة نصية. التحويل المباشر بدون API يتطلب ملف PDF نصي مُصدَّر من النظام أو ملف Excel/CSV.'
    );
  }

  onProgress?.('جاري تحليل الأعمدة وتنسيق الصفوف...');

  const flatLines: PdfLine[] = allPagesLines.flat().filter((l) => !isPageChromeLine(l));
  if (flatLines.length === 0) {
    throw new Error('لم يتم العثور على نصوص أو جداول واضحة داخل ملف الـ PDF.');
  }

  // Safely detect ONLY true repeating page chrome (must have no decimal quantities, no 4+ digit codes, and appear in top 3 or bottom 2 lines of >= 2 pages)
  const repeatingChrome = new Set<string>();
  if (allPagesLines.length >= 2) {
    const topLineCounts = new Map<string, number>();
    const bottomLineCounts = new Map<string, number>();

    const isSafeChromeCandidate = (l: PdfLine) => {
      const t = l.rawText.trim();
      if (!t) return false;
      if (l.chunks.some((c) => hasDecimalWeight(c.text) || /^\d{4,}$/.test(c.text.trim()))) {
        return false;
      }
      if (isLikelyFormulaGroupLine(l)) return false;
      return true;
    };

    for (const pLines of allPagesLines) {
      const seenTop = new Set<string>();
      for (const l of pLines.slice(0, 3)) {
        if (isSafeChromeCandidate(l)) {
          seenTop.add(l.rawText.trim());
        }
      }
      seenTop.forEach((t) => topLineCounts.set(t, (topLineCounts.get(t) || 0) + 1));

      const seenBottom = new Set<string>();
      for (const l of pLines.slice(-2)) {
        if (isSafeChromeCandidate(l)) {
          seenBottom.add(l.rawText.trim());
        }
      }
      seenBottom.forEach((b) => bottomLineCounts.set(b, (bottomLineCounts.get(b) || 0) + 1));
    }

    topLineCounts.forEach((cnt, t) => {
      if (cnt >= 2) repeatingChrome.add(t);
    });
    bottomLineCounts.forEach((cnt, b) => {
      if (cnt >= 2) repeatingChrome.add(b);
    });
  }

  // Search early lines of Page 1 (top 40 lines) for the true column header row so we never skip early data rows
  const page1Lines = (allPagesLines[0] || []).filter((l) => !isPageChromeLine(l));
  const searchPool = page1Lines.slice(0, Math.min(page1Lines.length, 40));

  let headerLine: PdfLine | null = null;
  let headerIdxInFlat = -1;
  let bestHeaderScore = 2;

  searchPool.forEach((line, idx) => {
    let score = scoreHeaderLine(line);
    const nextLine = searchPool[idx + 1];
    if (score > 0 && nextLine && nextLine.chunks.length >= 2 && nextLine.chunks.some((c) => isPureNumericLike(c.text))) {
      score += 3;
    }
    if (score > bestHeaderScore) {
      bestHeaderScore = score;
      headerLine = line;
    }
  });

  if (headerLine) {
    headerIdxInFlat = flatLines.indexOf(headerLine);
  } else {
    // Only treat earliest line as header if it contains non-numeric text-only multi-column headers
    const candidateIdx = searchPool.findIndex(
      (l) =>
        l.chunks.length >= 2 &&
        !/\d/.test(l.rawText) &&
        l.chunks.every((c) => !isPureNumericLike(c.text)) &&
        !isLikelyFormulaGroupLine(l)
    );
    if (candidateIdx !== -1 && candidateIdx <= 5) {
      headerLine = searchPool[candidateIdx];
      headerIdxInFlat = flatLines.indexOf(headerLine);
    }
  }

  // Document title from lines above the detected header on Page 1
  const titleLines =
    headerIdxInFlat > 0
      ? flatLines
          .slice(0, headerIdxInFlat)
          .filter((l) => l.pageNum === 1 && !isPageChromeLine(l))
          .map((l) => l.rawText)
      : [];
  const documentTitle =
    titleLines.find((t) => t.length >= 4 && t.length <= 120) ||
    file.name.replace(/\.[^/.]+$/, '');

  const headerNormSig = headerLine ? normalizeHeaderForCompare(headerLine) : '';

  const isMixingMode =
    options.extractionMode === 'mixing_reports' || looksLikeMixingReport(flatLines, file.name);

  // Detect RTL vs LTR
  const fullJoined = flatLines.slice(0, 120).map((l) => l.rawText).join(' ');
  const arabicCount = (fullJoined.match(/[\u0600-\u06FF]/g) || []).length;
  const latinCount = (fullJoined.match(/[A-Za-z]/g) || []).length;
  const isRtl = options.autoDetectRTL ? arabicCount > latinCount * 1.15 : false;

  // Keep ALL body lines after the header row (or from start if no header found), skipping ONLY repeated headers/chrome
  const startIdx = headerIdxInFlat !== -1 ? headerIdxInFlat + 1 : 0;
  const bodyLines: PdfLine[] = [];
  for (let i = startIdx; i < flatLines.length; i++) {
    const line = flatLines[i];
    const trimmed = line.rawText.trim();
    if (!trimmed) continue;
    if (repeatingChrome.has(trimmed)) continue;
    if (headerNormSig && normalizeHeaderForCompare(line) === headerNormSig) continue;
    bodyLines.push(line);
  }

  // Cluster columns from actual data rows (excluding formula group banners)
  const gridCandidateLines = isMixingMode
    ? bodyLines.filter((l) => !isLikelyFormulaGroupLine(l))
    : bodyLines;

  const clusters = buildColumnClusters(
    gridCandidateLines.length > 0 ? gridCandidateLines : bodyLines,
    headerLine
  );

  if (clusters.length === 0) {
    throw new Error('لم يتم العثور على أعمدة جداول قابلة للاستخراج في هذا الملف.');
  }

  let headers = headerLine
    ? assignLineToColumns(headerLine, clusters, isRtl).map(
        (h, i) => h || `Column_${i + 1}`
      )
    : clusters.map((_, i) => `Column_${i + 1}`);

  if (isMixingMode) {
    headers = normalizeMixingHeaders(headers);
  }

  const rawRows: string[][] = [];
  let grandTotalsRow: string[] | undefined;

  if (isMixingMode) {
    let currentFormula = '';
    let currentSupledger = '';
    let anySupledger = false;
    let justSawFormulaHeader = false;
    let formulaHeaderPageNum = -1;
    // true at the start of the table and right after a group subtotal: the next text-only line is a group title
    let expectBanner = true;
    const rowFormulas: string[] = [];
    const rowSupledgers: string[] = [];

    for (const line of bodyLines) {
      const isCodedBanner = isLikelyFormulaGroupLine(line);
      const isPlainTitle = !isCodedBanner && expectBanner && isPlainGroupTitleCandidate(line);
      if (isCodedBanner || isPlainTitle) {
        expectBanner = false;
        const joined = line.chunks.map((c) => c.text).join(' ').trim();
        const parsed = extractSupledgerAndCleanFormula(joined);
        if (parsed.cleanName) {
          currentFormula = parsed.cleanName;
          currentSupledger = parsed.supledger || '';
          if (currentSupledger) anySupledger = true;
          justSawFormulaHeader = true;
          formulaHeaderPageNum = line.pageNum;
        }
        continue;
      }

      const aligned = assignLineToColumns(line, clusters, isRtl);

      // Skip subtotal / group footer summary lines so ingredient weights are never double-counted
      if (isLikelySubtotalRow(aligned)) {
        justSawFormulaHeader = false;
        // Only a real group subtotal (numbers on the line) opens a new group; blank/decorative lines do not
        const subtotalJoined = aligned.join(' ');
        expectBanner =
          aligned.some((c) => hasDecimalWeight(c)) &&
          !/total\s*\[|grand total|الإجمالي/i.test(subtotalJoined);
        const joinedLower = aligned.join(' ').toLowerCase();
        if (
          joinedLower.includes('grand total') ||
          joinedLower.includes('الإجمالي الكلي') ||
          joinedLower.includes('الإجمالي العام') ||
          /\btotal\s*\[\s*(?:kg|كجم|ton|طن)\s*\]/i.test(joinedLower)
        ) {
          grandTotalsRow = aligned;
        }
        continue;
      }

      const nonEmptyCells = aligned.map((c) => c.trim()).filter(Boolean);
      if (nonEmptyCells.length === 0) continue;

      // Skip purely decorative separator lines (e.g. "-----", "=====")
      if (nonEmptyCells.every((c) => /^[-_=*•.|/\\]+$/.test(c))) {
        continue;
      }

      // Fold single-column wrapped continuation lines cleanly instead of creating blank orphan rows:
      // - Right after a formula banner (e.g. wrapped "A/3"): fold into currentFormula if not already present
      // - Right after an ingredient row (e.g. "Poultry meal" -> "feathers", "Vitamin Startar" -> "DSM"): fold into previous row's name
      const hasAnyDecimal = aligned.some((c) => hasDecimalWeight(c));
      const nonEmptyColIndices = aligned
        .map((c, idx) => (c.trim() ? idx : -1))
        .filter((idx) => idx !== -1);

      if (!hasAnyDecimal && nonEmptyColIndices.length === 1 && nonEmptyColIndices[0] <= 1) {
        const wrapCol = nonEmptyColIndices[0];
        const wrapText = aligned[wrapCol].trim();
        const isPure4PlusDigits = /^\d{4,}$/.test(wrapText.replace(/,/g, ''));
        if (
          !isPure4PlusDigits &&
          wrapText.length <= 35 &&
          !/^(?:total|page|صفحة)/i.test(wrapText)
        ) {
          if (
            justSawFormulaHeader &&
            currentFormula &&
            line.pageNum === formulaHeaderPageNum
          ) {
            if (!currentFormula.toLowerCase().includes(wrapText.toLowerCase())) {
              currentFormula = `${currentFormula} ${wrapText}`.trim();
            }
            continue;
          }
          if (rawRows.length > 0) {
            const prevRow = rawRows[rawRows.length - 1];
            const targetCol = prevRow[wrapCol]?.trim()
              ? wrapCol
              : prevRow[0]?.trim()
              ? 0
              : wrapCol;
            if (prevRow[targetCol] !== undefined) {
              prevRow[targetCol] = `${prevRow[targetCol]} ${wrapText}`.trim();
            }
            continue;
          }
        }
      }

      justSawFormulaHeader = false;
      expectBanner = false;
      rawRows.push(aligned);
      rowFormulas.push(currentFormula);
      rowSupledgers.push(currentSupledger);
    }

    const nonEmptyColIndices = headers
      .map((_, cIdx) => cIdx)
      .filter(
        (cIdx) =>
          (headers[cIdx] && !headers[cIdx].startsWith('Column_')) ||
          rawRows.some((r) => (r[cIdx] || '').trim() !== '')
      );

    headers = nonEmptyColIndices.map((idx) => headers[idx]);
    const compactedRows = rawRows.map((r) => nonEmptyColIndices.map((idx) => r[idx] || ''));
    if (grandTotalsRow) {
      grandTotalsRow = nonEmptyColIndices.map((idx) => grandTotalsRow?.[idx] || '');
    }

    const hasFormulaAlready = headers.some((h) => {
      const n = h.toLowerCase();
      return n.includes('finished') || n.includes('formula') || n.includes('تركيبة');
    });

    if (!hasFormulaAlready && rowFormulas.some(Boolean)) {
      headers = ['Finished_Product', ...(anySupledger ? ['Supledger'] : []), ...headers];
      compactedRows.forEach((r, idx) => {
        if (anySupledger) {
          r.unshift(rowSupledgers[idx] || '');
        }
        r.unshift(rowFormulas[idx] || documentTitle);
      });
      if (grandTotalsRow) {
        if (anySupledger) grandTotalsRow.unshift('');
        grandTotalsRow.unshift('الإجمالي العام');
      }
    }

    rawRows.length = 0;
    compactedRows.forEach((r) => rawRows.push(r));
  } else {
    for (const line of bodyLines) {
      const aligned = assignLineToColumns(line, clusters, isRtl);
      const nonEmptyCount = aligned.filter((c) => c.trim() !== '').length;
      if (nonEmptyCount === 0) continue;

      const joinedLower = aligned.join(' ').toLowerCase();
      if (
        (joinedLower.includes('grand total') ||
          joinedLower.includes('الإجمالي الكلي') ||
          joinedLower.includes('الإجمالي العام')) &&
        !grandTotalsRow
      ) {
        grandTotalsRow = aligned;
        continue;
      }

      rawRows.push(aligned);
    }

    const activeColIndices = headers
      .map((_, cIdx) => cIdx)
      .filter(
        (cIdx) =>
          !headers[cIdx].startsWith('Column_') ||
          rawRows.some((r) => (r[cIdx] || '').trim() !== '')
      );
    headers = activeColIndices.map((idx) => headers[idx]);
    const cleaned = rawRows.map((r) => activeColIndices.map((idx) => r[idx] || ''));
    rawRows.length = 0;
    cleaned.forEach((r) => rawRows.push(r));
    if (grandTotalsRow) {
      grandTotalsRow = activeColIndices.map((idx) => grandTotalsRow?.[idx] || '');
    }
  }

  if (rawRows.length === 0) {
    throw new Error('تعذر العثور على صفوف بيانات منظمة داخل ملف الـ PDF.');
  }

  onProgress?.('جاري فصل الأكواد وتصنيف الأصناف وتجهيز صفحات الكنترول...');

  const docDraft: ExtractedDocument = {
    documentTitle,
    documentSummary: `تم استخراج ${rawRows.length} صف من ${numPages} صفحة محلياً بدون الحاجة إلى API أو ذكاء اصطناعي.`,
    detectedLanguage: arabicCount > latinCount ? 'العربية' : 'English',
    detectedType: isMixingMode ? 'mixing_report' : 'table',
    sheets: [],
    metadata: {
      pageCountEstimate: `${numPages} صفحة`,
      rowCountTotal: rawRows.length,
      columnCountMax: headers.length,
    },
    notes: [
      `تم التحويل محلياً مباشرة داخل المتصفح (${numPages} صفحة، ${rawRows.length} صف) بدون استهلاك أي API.`,
    ],
  };

  const baseSheet: SheetData = {
    sheetName: isMixingMode ? 'Mixing_Report' : 'Extracted_Table',
    description: `جدول مستخرج محلياً من ${file.name}`,
    isRtl,
    headers,
    rows: rawRows,
    ...(grandTotalsRow ? { totalsRow: grandTotalsRow } : {}),
  };

  const processedSheet = postProcessSheet(baseSheet, docDraft);
  docDraft.sheets = [processedSheet];
  docDraft.metadata = {
    pageCountEstimate: `${numPages} صفحة`,
    rowCountTotal: processedSheet.rows.length,
    columnCountMax: processedSheet.headers.length,
  };

  return docDraft;
}

/**
 * Detects and parses the Crystal Reports "ConsumptionByFinishedProductList" layout when it is
 * exported straight to Excel (merged cells, formula banner rows, Ident / Name / Target / Actual columns).
 * Produces the SAME raw table shape the PDF path builds, so the rest of the pipeline
 * (code split, classification, Sadat control report, pivot ...) behaves identically for Excel and PDF.
 */
function parseCrystalConsumptionMatrix(
  matrix: any[][],
  fileName: string
): { sheet: SheetData; reportDate: string } | null {
  const cellText = (v: any) => (v === null || v === undefined ? '' : String(v).replace(/\s+/g, ' ').trim());
  const isNum = (v: any) =>
    typeof v === 'number' ? Number.isFinite(v) : /^[+-]?[\d,]+(?:\.\d+)?$/.test(cellText(v));
  const toNum = (v: any) => (typeof v === 'number' ? v : parseFloat(cellText(v).replace(/,/g, '')));

  // 1. Find the header row: contains "Ident" and "Name" plus Target/Actual labels
  let headerIdx = -1;
  for (let r = 0; r < Math.min(60, matrix.length); r++) {
    const texts = (matrix[r] || []).map(cellText).map((t) => t.toLowerCase());
    if (
      texts.some((t) => t === 'ident') &&
      texts.some((t) => t === 'name') &&
      texts.some((t) => /^(target|actual)\b/.test(t))
    ) {
      headerIdx = r;
      break;
    }
  }
  if (headerIdx === -1) return null;

  const headerTexts = (matrix[headerIdx] || []).map(cellText);
  const hasTarget = headerTexts.some((t) => /^target/i.test(t));
  const hasActual = headerTexts.some((t) => /^actual/i.test(t));

  // 2. Report date: prefer the "To:" date, then any date in the first rows, then the printed footer date
  let reportDate = '';
  const topText = matrix.slice(0, headerIdx).flat().map(cellText).filter(Boolean);
  const toLine = topText.find((t) => /^to\s*:/i.test(t));
  const dateRe = /(\d{1,2}\/\d{1,2}\/\d{2,4})/;
  if (toLine) reportDate = (toLine.match(dateRe) || [])[1] || '';
  if (!reportDate) {
    for (const t of topText) {
      const m = t.match(dateRe);
      if (m) {
        reportDate = m[1];
        break;
      }
    }
  }
  if (!reportDate) {
    for (let r = matrix.length - 1; r >= Math.max(0, matrix.length - 10); r--) {
      const m = (matrix[r] || []).map(cellText).join(' ').match(dateRe);
      if (m) {
        reportDate = m[1];
        break;
      }
    }
  }

  // 3. Walk the body
  const rows: (string | number)[][] = [];
  let totalsRow: (string | number)[] | undefined;
  let currentFormula = '';
  let currentSupledger = '';
  let anySupledger = false;

  for (let r = headerIdx + 1; r < matrix.length; r++) {
    const row = matrix[r] || [];
    // Non-empty cells in column order (merged-cell exports shift columns, so order matters, not position)
    const cells = row
      .map((c) => ({ text: cellText(c), isNumber: typeof c === 'number' }))
      .filter((c) => c.text !== '');
    if (cells.length === 0) continue;

    const first = cells[0];
    // Page footer: "16/11/2026 08:12:39" / "Page 1 of 1"
    if (
      cells.every((c) => /^\d{1,2}\/\d{1,2}\/\d{2,4}(\s+[\d:]+)?$|^page\s+\d+/i.test(c.text))
    ) {
      continue;
    }

    // Grand total: "Total [kg]" followed by numeric Target / Actual
    if (!first.isNumber && /^total\b/i.test(first.text)) {
      const w = cells.slice(1).filter((c) => isNum(c.text)).map((c) => toNum(c.text));
      if (w.length > 0) {
        const vals = hasTarget && hasActual && w.length >= 2 ? [w[0], w[1]] : [w[w.length - 1]];
        totalsRow = ['الإجمالي العام', '', '', '', ...vals.map((v) => String(v))];
      }
      continue;
    }

    // Per-formula subtotal rows start directly with a number (no item name)
    if (first.isNumber) continue;

    const shortName = first.text;
    const longName = cells[1] ? cells[1].text : '';
    // Weights = numeric-typed cells after the name cells (a bare "1287242" code is a STRING cell, so it is never mistaken for a weight)
    const weights = cells.slice(2).filter((c) => c.isNumber || isNum(c.text)).map((c) => toNum(c.text));

    // Formula banner: short name + "(name) 12345" with no weights
    if (weights.length === 0) {
      const fromLong = extractSupledgerAndCleanFormula(longName || shortName);
      const fromShort = extractSupledgerAndCleanFormula(shortName);
      currentFormula = fromLong.cleanName || fromShort.cleanName || shortName;
      currentSupledger = fromLong.supledger || fromShort.supledger || '';
      if (currentSupledger) anySupledger = true;
      continue;
    }

    // Ingredient row
    let target: number | '' = '';
    let actual: number | '' = '';
    if (weights.length >= 2) {
      target = weights[0];
      actual = weights[1];
    } else if (hasActual && !hasTarget) {
      actual = weights[0];
    } else {
      target = weights[0];
    }
    // Split "(corn 7.5%) 11342" / "1287242" / "Lysine HCL 13556" into a JDE code, and keep the short name as the item name
    let itemCode = '';
    if (/^\d{4,}$/.test(longName)) {
      itemCode = longName;
    } else {
      const parsed = extractCodeFromValue(longName);
      if (parsed) itemCode = parsed.code;
    }
    const itemName = shortName || (extractCodeFromValue(longName)?.cleanText ?? longName.replace(/^\(|\)$/g, ''));

    const out: (string | number)[] = [
      currentFormula || fileName.replace(/\.[^/.]+$/, ''),
      currentSupledger,
      itemCode,
      itemName,
    ];
    if (hasTarget) out.push(target === '' ? '' : String(target));
    if (hasActual) out.push(actual === '' ? '' : String(actual));
    rows.push(out);
  }

  if (rows.length === 0) return null;

  // Same canonical header names the PDF path produces, so every downstream step treats Excel and PDF identically
  const headers = ['Finished_Product', 'Supledger', 'Ingredient_Code', 'Ingredient_Name'];
  if (hasTarget) headers.push('Target [kg]');
  if (hasActual) headers.push('Actual [kg]');

  const sheet: SheetData = {
    sheetName: 'Mixing_Report',
    description: `جدول مستخرج محلياً من ${fileName}`,
    isRtl: false,
    headers,
    rows,
    ...(totalsRow ? { totalsRow } : {}),
  };
  return { sheet, reportDate };
}

export async function extractDocumentFromSpreadsheetLocal(
  file: File,
  options: ConversionOptions,
  onProgress?: (step: string) => void
): Promise<ExtractedDocument> {
  onProgress?.('جاري قراءة ملف البيانات محلياً...');
  const buffer = await file.arrayBuffer();
  const isCsv = file.name.toLowerCase().endsWith('.csv');

  let wb: XLSX.WorkBook;
  if (isCsv) {
    // Decode UTF-8 explicitly so Google Sheets CSV exports without BOM preserve Arabic text cleanly
    const csvText = new TextDecoder('utf-8').decode(buffer).replace(/^\ufeff/, '');
    wb = XLSX.read(csvText, { type: 'string' });
  } else {
    wb = XLSX.read(buffer, { type: 'array' });
  }

  if (!wb.SheetNames || wb.SheetNames.length === 0) {
    throw new Error('الملف لا يحتوي على أي أوراق عمل.');
  }

  let detectedReportDate = '';
  const docDraft: ExtractedDocument = {
    documentTitle: file.name.replace(/\.[^/.]+$/, ''),
    documentSummary: `تم استيراد ${wb.SheetNames.length} ورقة عمل محلياً بدون API.`,
    detectedLanguage: 'Auto',
    detectedType: 'spreadsheet',
    sheets: [],
    notes: ['تم استيراد البيانات محلياً مباشرة بدون الحاجة إلى API.'],
  };

  const sheets: SheetData[] = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rawMatrix: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Crystal Reports "ConsumptionByFinishedProductList" exported to Excel -> same pipeline as the PDF
    const crystal = parseCrystalConsumptionMatrix(rawMatrix, file.name);
    if (crystal) {
      if (crystal.reportDate && !detectedReportDate) detectedReportDate = crystal.reportDate;
      docDraft.detectedType = 'mixing_report';
      sheets.push(postProcessSheet(crystal.sheet, docDraft));
      continue;
    }
    const nonEmptyRows = rawMatrix.filter(
      (r) => Array.isArray(r) && r.some((cell) => String(cell ?? '').trim() !== '')
    );
    if (nonEmptyRows.length < 2) continue;

    // Extract embedded report date if present in early header/dashboard rows (e.g. "تاريخ التقرير: 16/11/2026")
    for (let r = 0; r < Math.min(5, nonEmptyRows.length); r++) {
      for (let c = 0; c < nonEmptyRows[r].length; c++) {
        const val = String(nonEmptyRows[r][c] ?? '').trim();
        const m = val.match(/\b(\d{1,2}\/\d{1,2}\/\d{2,4})\b/);
        if (m && !detectedReportDate) {
          detectedReportDate = m[1];
        }
      }
    }

    // Detect Google Sheets Mixing-Dump layout: Columns A, B, C hold [Formula, Ingredient, Qty]
    // while right-side columns hold pivot summaries / totals (often leaving Row 0 cols A-C blank).
    const mixingTriples: string[][] = [];
    for (let r = 0; r < nonEmptyRows.length; r++) {
      const row = nonEmptyRows[r];
      const col0 = normalizeText(String(row[0] ?? ''), options.parseNumbers);
      const col1 = normalizeText(String(row[1] ?? ''), options.parseNumbers);
      const col2 = normalizeText(String(row[2] ?? ''), options.parseNumbers);
      const num2 = Number(col2.replace(/,/g, ''));
      if (
        col0 &&
        col1 &&
        col2 &&
        !Number.isNaN(num2) &&
        !/^(?:finished|formula|تركيبة|الصنف|ingredient|الاجمالي|الإجمالي|grand total)/i.test(col0) &&
        !/^(?:ingredient|الصنف|اسم الصنف|الاجمالي|الإجمالي|grand total)/i.test(col1)
      ) {
        mixingTriples.push([col0, col1, col2]);
      }
    }

    const row0HasBlankLeadCols =
      String(nonEmptyRows[0][0] ?? '').trim() === '' &&
      String(nonEmptyRows[0][1] ?? '').trim() === '' &&
      String(nonEmptyRows[0][2] ?? '').trim() === '';

    if (mixingTriples.length >= 5 && (row0HasBlankLeadCols || nonEmptyRows[0].length >= 6)) {
      const mixingSheet: SheetData = {
        sheetName: 'Mixing_Report',
        description: `بيانات الخلطات المستوردة من ${file.name} (${mixingTriples.length} سطر)`,
        isRtl: false,
        headers: ['Finished_Product', 'Ingredient_Name', 'Actual_kg'],
        rows: mixingTriples,
      };
      docDraft.detectedType = 'mixing_report';
      sheets.push(postProcessSheet(mixingSheet, docDraft));
    }

    const headers = nonEmptyRows[0].map((c, idx) => {
      const s = normalizeText(String(c ?? ''), options.parseNumbers);
      return s || `Column_${idx + 1}`;
    });

    const rows = nonEmptyRows.slice(1).map((r) =>
      headers.map((_, idx) => normalizeText(String(r[idx] ?? ''), options.parseNumbers))
    );

    const sampleText = `${headers.join(' ')} ${rows.slice(0, 10).flat().join(' ')}`;
    const arabicCount = (sampleText.match(/[\u0600-\u06FF]/g) || []).length;
    const latinCount = (sampleText.match(/[A-Za-z]/g) || []).length;
    const isRtl = options.autoDetectRTL ? arabicCount > latinCount : false;

    const rawSheetObj: SheetData = {
      sheetName:
        mixingTriples.length >= 5 && (row0HasBlankLeadCols || nonEmptyRows[0].length >= 6)
          ? 'Original_Sheet'
          : sheetName.slice(0, 31),
      description: `ورقة عمل مستوردة من ${file.name}`,
      isRtl,
      headers,
      rows,
    };

    sheets.push(
      rawSheetObj.sheetName === 'Original_Sheet'
        ? rawSheetObj
        : postProcessSheet(rawSheetObj, docDraft)
    );
  }

  if (detectedReportDate) {
    docDraft.documentTitle = `${docDraft.documentTitle} - ${detectedReportDate}`;
  }

  if (sheets.length === 0) {
    throw new Error('لم يتم العثور على صفوف بيانات صالحة داخل الملف.');
  }

  docDraft.sheets = sheets;
  docDraft.metadata = {
    pageCountEstimate: `${sheets.length} ورقة`,
    rowCountTotal: sheets[0].rows.length,
    columnCountMax: Math.max(...sheets.map((s) => s.headers.length)),
  };

  return docDraft;
}
