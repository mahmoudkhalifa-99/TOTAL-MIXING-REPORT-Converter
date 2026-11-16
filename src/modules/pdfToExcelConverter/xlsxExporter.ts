import * as XLSX from 'xlsx';
import { ExportSheetStrategy, ExtractedTable } from './types';

function sanitizeSheetName(raw: string, usedNames: Set<string>): string {
  const cleaned = (raw || 'Sheet')
    .replace(/[\\/*?:\[\]]/g, '_')
    .trim()
    .slice(0, 31) || 'Sheet';

  if (!usedNames.has(cleaned)) {
    usedNames.add(cleaned);
    return cleaned;
  }

  let idx = 2;
  while (idx < 999) {
    const suffix = `_${idx}`;
    const candidate = `${cleaned.slice(0, 31 - suffix.length)}${suffix}`;
    if (!usedNames.has(candidate)) {
      usedNames.add(candidate);
      return candidate;
    }
    idx++;
  }
  return `Sheet_${Date.now() % 10000}`;
}

/**
 * Converts pure simple integer/decimal strings to numeric cells ONLY when
 * exact numeric round-trip keeps the original digits unchanged (preserves leading zeros,
 * trailing decimal formatting like 12.50, dates, commas, and financial formatting).
 */
function preserveCellValue(raw: string): string | number {
  const val = raw ?? '';
  const trimmed = val.trim();
  if (!trimmed) return '';

  // Keep leading zeros like "00142" as text
  if (/^0\d+$/.test(trimmed)) return trimmed;

  // Keep formatted decimals with trailing zeros like "12.50" or comma-formatted "1,250.00" as exact text
  // OR store clean numbers if exact round-trip matches identically
  if (/^-?\d+(?:\.\d+)?$/.test(trimmed)) {
    const num = Number(trimmed);
    if (Number.isFinite(num) && String(num) === trimmed) {
      return num;
    }
  }
  return val;
}

function buildWorksheetFromGrid(
  headers: string[],
  rows: string[][],
  isRtl: boolean
): XLSX.WorkSheet {
  const aoa: (string | number)[][] = [];
  if (headers.length > 0) {
    aoa.push(headers.map((h) => h ?? ''));
  }
  for (const row of rows) {
    const fullRow = new Array(Math.max(headers.length, row.length)).fill('');
    for (let c = 0; c < fullRow.length; c++) {
      fullRow[c] = preserveCellValue(row[c] ?? '');
    }
    aoa.push(fullRow);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Set RTL sheet view in Excel when needed
  if (isRtl) {
    ws['!views'] = [{ RTL: true }];
  }

  // Auto-fit column widths gently based on content length
  const colCount = Math.max(headers.length, ...rows.map((r) => r.length), 1);
  ws['!cols'] = Array.from({ length: colCount }, (_, colIdx) => {
    let maxLen = (headers[colIdx] || '').length;
    for (let r = 0; r < Math.min(rows.length, 200); r++) {
      const len = (rows[r]?.[colIdx] || '').length;
      if (len > maxLen) maxLen = len;
    }
    return { wch: Math.min(Math.max(maxLen + 4, 12), 55) };
  });

  return ws;
}

/**
 * Exports extracted tables to an .xlsx workbook via SheetJS (xlsx) without altering
 * numeric/financial strings or dropping empty cells/rows.
 */
export function exportTablesToXlsx(
  tables: ExtractedTable[],
  fileName: string,
  strategy: ExportSheetStrategy
): void {
  if (!tables || tables.length === 0) {
    throw new Error('لا توجد جداول مستخرجة للتصدير.');
  }

  const wb = XLSX.utils.book_new();
  const usedNames = new Set<string>();

  if (strategy === 'separate_sheets_per_table') {
    tables.forEach((tbl, idx) => {
      const sheetName = sanitizeSheetName(tbl.name || `جدول_${idx + 1}`, usedNames);
      const ws = buildWorksheetFromGrid(tbl.headers, tbl.rows, tbl.isRtl);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
  } else if (strategy === 'separate_sheets_per_page') {
    const byPage = new Map<number, ExtractedTable[]>();
    for (const tbl of tables) {
      const pageKey = tbl.pages[0] ?? 1;
      const list = byPage.get(pageKey) || [];
      list.push(tbl);
      byPage.set(pageKey, list);
    }

    const sortedPages = Array.from(byPage.keys()).sort((a, b) => a - b);
    for (const pNum of sortedPages) {
      const pageTables = byPage.get(pNum) || [];
      const aoa: (string | number)[][] = [];
      let anyRtl = false;

      pageTables.forEach((tbl, tIdx) => {
        if (tbl.isRtl) anyRtl = true;
        if (tIdx > 0) aoa.push([]);
        if (pageTables.length > 1) {
          aoa.push([tbl.name]);
        }
        if (tbl.headers.length > 0) {
          aoa.push(tbl.headers);
        }
        tbl.rows.forEach((r) => aoa.push(r.map((cell) => preserveCellValue(cell))));
      });

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      if (anyRtl) ws['!views'] = [{ RTL: true }];
      const sheetName = sanitizeSheetName(`صفحة_${pNum}`, usedNames);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    }
  } else {
    // single_merged_sheet
    const aoa: (string | number)[][] = [];
    let anyRtl = false;
    tables.forEach((tbl, tIdx) => {
      if (tbl.isRtl) anyRtl = true;
      if (tIdx > 0) aoa.push([]);
      if (tables.length > 1) {
        aoa.push([`${tbl.name} (صفحات: ${tbl.pages.join(', ')})`]);
      }
      if (tbl.headers.length > 0) {
        aoa.push(tbl.headers);
      }
      tbl.rows.forEach((r) => aoa.push(r.map((cell) => preserveCellValue(cell))));
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (anyRtl) ws['!views'] = [{ RTL: true }];
    const sheetName = sanitizeSheetName('كافة_الجداول', usedNames);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  const baseName = (fileName || 'converted').replace(/\.[^/.]+$/, '');
  XLSX.writeFile(wb, `${baseName}_PDF_to_Excel.xlsx`);
}
