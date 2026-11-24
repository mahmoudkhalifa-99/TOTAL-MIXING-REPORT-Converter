import * as XLSX from 'xlsx';
import { SheetData } from '../types/converter';
import {
  calculateControlPivot,
  buildControlPivotSheetData,
  loadStoredPivotColumnOrder,
  PIVOT_COLUMN_DEFS,
} from './controlPivotCalculator';

/**
 * Checks if a string looks like a standard number (excluding phone numbers, dates, IDs)
 */
function tryParseNumber(val: any): any {
  if (typeof val === 'number') return val;
  if (typeof val !== 'string') return val;
  const trimmed = val.trim();
  if (trimmed === '') return '';

  // Avoid converting leading zeros like "00123" or phone numbers like "+966..." or "055..."
  if (/^0[0-9]+$/.test(trimmed)) return trimmed;
  if (/^\+?[0-9\s-]{10,}$/.test(trimmed)) return trimmed;

  // Clean currency symbols and commas (e.g., "$1,250.50" or "1,250.50 ريال")
  const cleanNumeric = trimmed
    .replace(/[$,€£¥]/g, '')
    .replace(/\s*(SAR|EGP|AED|USD|EUR|ريال|جنيه|درهم|دولار)\s*/gi, '')
    .replace(/,/g, '')
    .trim();

  if (/^-?\d+(\.\d+)?$/.test(cleanNumeric)) {
    const num = parseFloat(cleanNumeric);
    if (!isNaN(num)) return num;
  }

  return val;
}

/**
 * Exports multiple sheets into a native .xlsx Excel workbook
 */
export function exportToExcel(
  sheets: SheetData[],
  fileName: string = 'converted_document.xlsx',
  parseNumbers: boolean = true
) {
  const wb = XLSX.utils.book_new();

  sheets.forEach((sheet, index) => {
    // Prepare table array with headers and rows
    const tableData: any[][] = [];

    // Header row
    tableData.push(sheet.headers);

    // Data rows
    sheet.rows.forEach((row) => {
      const processedRow = row.map((cell) => (parseNumbers ? tryParseNumber(cell) : cell));
      tableData.push(processedRow);
    });

    // Optional totals row
    if (sheet.totalsRow && sheet.totalsRow.length > 0) {
      const processedTotal = sheet.totalsRow.map((cell) =>
        parseNumbers ? tryParseNumber(cell) : cell
      );
      tableData.push(processedTotal);
    }

    const ws = XLSX.utils.aoa_to_sheet(tableData);

    // Calculate column widths
    const colWidths = sheet.headers.map((h, colIndex) => {
      let maxLen = (h || '').toString().length;
      sheet.rows.forEach((row) => {
        const cell = row[colIndex];
        if (cell !== undefined && cell !== null) {
          const len = cell.toString().length;
          if (len > maxLen) maxLen = len;
        }
      });
      return { wch: Math.min(Math.max(maxLen + 3, 12), 45) };
    });
    ws['!cols'] = colWidths;

    // RTL configuration for Excel
    if (sheet.isRtl) {
      ws['!views'] = [{ rightToLeft: true }];
    }

    // Sanitize sheet name (Excel limits: max 31 chars, no : \ / ? * [ ], and must be unique)
    let baseName = (sheet.sheetName || `Sheet${index + 1}`)
      .replace(/[:\\/?*\[\]]/g, '_')
      .trim()
      .substring(0, 31);
    if (!baseName) baseName = `Sheet${index + 1}`;

    let safeName = baseName;
    let suffixNum = 2;
    while (wb.SheetNames.includes(safeName)) {
      const suffix = `_${suffixNum++}`;
      safeName = baseName.substring(0, 31 - suffix.length) + suffix;
    }

    XLSX.utils.book_append_sheet(wb, ws, safeName);
  });

  // If any sheet has "نوع الصنف" column, append a dedicated Summary sheet for the 4 categories
  try {
    const sheetWithType = sheets.find(
      (s) => s.headers && s.headers.some((h) => /نوع الصنف|item type|category/i.test(h))
    );

    if (sheetWithType) {
      const summaryRows: any[][] = [];
      summaryRows.push(['ملخص إجماليات الأنواع الأربعة (سمك • تسمين • بط • أليف) والإجمالي الكلي']);
      summaryRows.push(['تاريخ الاستخراج:', new Date().toLocaleDateString('ar-EG'), 'المصنف:', fileName]);
      summaryRows.push([]);
      summaryRows.push([
        'نوع الصنف (Category)',
        'الوصف والرموز',
        'عدد البنود (Count)',
        'إجمالي المستهدف (Target)',
        'إجمالي الفعلي (Actual)',
        'الفرق (Variance)',
        'نسبة المساهمة الفعلي (%)',
      ]);

      // Calculate totals for each category
      const categories = ['سمك', 'تسمين', 'بط', 'اليف', 'تنظيف'] as const;
      const typeCol = sheetWithType.headers.findIndex((h) => /نوع الصنف|item type|category/i.test(h));
      const targetCol = sheetWithType.headers.findIndex((h) => /target|المستهدف/i.test(h));
      const actualCol = sheetWithType.headers.findIndex((h) => /actual|الفعلي/i.test(h));

      let grandCount = 0;
      let grandTarget = 0;
      let grandActual = 0;

      const catMeta: Record<string, { label: string; desc: string }> = {
        'سمك': { label: '🐟 سمك (Fish)', desc: 'الأصناف التي تحتوي تركيبتها على Tilapia' },
        'تسمين': { label: '🌾 تسمين (Broiler)', desc: 'الأصناف التي تحتوي تركيبتها على GR, FN, ST, SS' },
        'بط': { label: '🦆 بط (Duck)', desc: 'الأصناف التي تحتوي تركيبتها على Duck' },
        'اليف': { label: '🐾 أليف (Pets)', desc: 'كافة الأصناف والخلطات الأخرى' },
        'تنظيف': { label: '🧹 Cleaning (تنظيف)', desc: 'مجموعات التنظيف - غير مشمولة بالإجمالي الكلي' },
      };

      const catSums: Record<string, { count: number; target: number; actual: number }> = {
        'سمك': { count: 0, target: 0, actual: 0 },
        'تسمين': { count: 0, target: 0, actual: 0 },
        'بط': { count: 0, target: 0, actual: 0 },
        'اليف': { count: 0, target: 0, actual: 0 },
        'تنظيف': { count: 0, target: 0, actual: 0 },
      };

      sheetWithType.rows.forEach((row) => {
        let cat = 'اليف';
        if (typeCol !== -1) {
          const val = (row[typeCol] ?? '').toString().trim();
          if (val === 'سمك' || val === 'تسمين' || val === 'بط' || val === 'اليف' || val === 'تنظيف') {
            cat = val;
          }
        }
        const tVal = targetCol !== -1 ? parseFloat(String(row[targetCol] || '0').replace(/,/g, '')) || 0 : 0;
        const aVal = actualCol !== -1 ? parseFloat(String(row[actualCol] || '0').replace(/,/g, '')) || 0 : 0;

        catSums[cat].count++;
        catSums[cat].target += tVal;
        catSums[cat].actual += aVal;

        // Cleaning is reported on its own line and is not part of the grand total
        if (cat !== 'تنظيف') {
          grandCount++;
          grandTarget += tVal;
          grandActual += aVal;
        }
      });

      categories.forEach((catKey) => {
        const item = catSums[catKey];
        const meta = catMeta[catKey];
        const variance = item.actual - item.target;
        const pct = grandActual > 0 ? (item.actual / grandActual) * 100 : 0;

        summaryRows.push([
          meta.label,
          meta.desc,
          item.count,
          Number(item.target.toFixed(3)),
          Number(item.actual.toFixed(3)),
          Number(variance.toFixed(3)),
          `${pct.toFixed(2)}%`,
        ]);
      });

      // Grand Total Row
      const grandVariance = grandActual - grandTarget;
      summaryRows.push([]);
      summaryRows.push([
        '⭐ الإجمالي الكلي للأربعة (Grand Total)',
        'مجموع كافة الأصناف الـ 4 معاً',
        grandCount,
        Number(grandTarget.toFixed(3)),
        Number(grandActual.toFixed(3)),
        Number(grandVariance.toFixed(3)),
        '100.00%',
      ]);

      const summaryWs = XLSX.utils.aoa_to_sheet(summaryRows);
      summaryWs['!cols'] = [
        { wch: 25 },
        { wch: 38 },
        { wch: 16 },
        { wch: 22 },
        { wch: 22 },
        { wch: 18 },
        { wch: 22 },
      ];
      summaryWs['!views'] = [{ rightToLeft: true }];
      if (!wb.SheetNames.includes('ملخص_إجماليات_الأربعة')) {
        XLSX.utils.book_append_sheet(wb, summaryWs, 'ملخص_إجماليات_الأربعة');
      }

      // Also append the Control Aggregation Pivot Sheet (كود الصنف | إسم الصنف | تسمين | منصرف سمك | منصرف بط | حيوانات أليفة | الاجمالى)
      if (!wb.SheetNames.includes('تجميع_الكنترول')) {
        const joinedHeaders = (sheetWithType.headers || []).join(' ').toLowerCase();
        const isKg =
          joinedHeaders.includes('kg') ||
          joinedHeaders.includes('كجم') ||
          (!joinedHeaders.includes('ton') && !joinedHeaders.includes('طن'));
        const divisor = isKg ? 1000 : 1;
        const decimals = isKg ? 6 : 3;
        const pivot = calculateControlPivot(sheetWithType, 'actual', divisor);
        if (pivot.rows.length > 0) {
          const pivotColOrder = loadStoredPivotColumnOrder();
          const pivotSheet = buildControlPivotSheetData(pivot, decimals, pivotColOrder);
          const pivotAoa: any[][] = [pivotSheet.headers, ...pivotSheet.rows];
          if (pivotSheet.totalsRow) {
            pivotAoa.push(pivotSheet.totalsRow);
          }
          const pivotWs = XLSX.utils.aoa_to_sheet(pivotAoa);
          pivotWs['!cols'] = pivotColOrder.map((colKey) => ({
            wch: PIVOT_COLUMN_DEFS[colKey].width,
          }));
          pivotWs['!views'] = [{ rightToLeft: true }];
          XLSX.utils.book_append_sheet(wb, pivotWs, 'تجميع_الكنترول');
        }
      }
    }
  } catch (err) {
    console.warn('Could not append category summary sheet:', err);
  }

  // Ensure fileName ends with .xlsx
  const safeFileName = fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`;
  XLSX.writeFile(wb, safeFileName);
}

/**
 * Exports a single sheet to CSV
 */
export function exportToCsv(sheet: SheetData, fileName: string = 'table.csv') {
  const tableData: any[][] = [];
  tableData.push(sheet.headers);
  sheet.rows.forEach((row) => tableData.push(row));
  if (sheet.totalsRow) {
    tableData.push(sheet.totalsRow);
  }

  const ws = XLSX.utils.aoa_to_sheet(tableData);
  const csv = XLSX.utils.sheet_to_csv(ws);

  // Add UTF-8 BOM so Excel opens Arabic correctly in CSV
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName.endsWith('.csv') ? fileName : `${fileName}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Copies table data as TSV (tab-separated) directly to system clipboard
 */
export async function copyTableToClipboard(sheet: SheetData): Promise<boolean> {
  const tableData: any[][] = [];
  tableData.push(sheet.headers);
  sheet.rows.forEach((row) => tableData.push(row));
  if (sheet.totalsRow) {
    tableData.push(sheet.totalsRow);
  }

  const tsv = tableData.map((row) => row.join('\t')).join('\n');
  try {
    await navigator.clipboard.writeText(tsv);
    return true;
  } catch (err) {
    console.error('Failed to copy to clipboard', err);
    return false;
  }
}
