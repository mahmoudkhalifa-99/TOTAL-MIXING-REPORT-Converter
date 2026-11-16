import { SheetData } from '../types/converter';
import { findItemTypeColumnIndex, findFormulaColumnIndex } from './itemClassifier';

/**
 * Converts Arabic-Indic numerals (٠١٢٣٤٥٦٧٨٩) to standard Latin digits (0123456789)
 * and cleans commas, spaces, currency symbols, and units.
 */
export function parseCleanNumber(val: any): number {
  if (typeof val === 'number') {
    return isNaN(val) ? 0 : val;
  }
  if (val === null || val === undefined) return 0;

  let str = val.toString().trim();
  if (!str) return 0;

  // Replace Arabic-Indic digits
  const arabicIndicDigits = ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩'];
  for (let i = 0; i < 10; i++) {
    str = str.replaceAll(arabicIndicDigits[i], i.toString());
  }

  // Remove common currency symbols and textual labels
  str = str
    .replace(/[$,€£¥]/g, '')
    .replace(/\s*(SAR|EGP|AED|USD|EUR|ر\.س|ريال|جنيه|درهم|دولار|طن|كجم|طن\/ساعة|ton|tons|kg)\s*/gi, '')
    .replace(/,/g, '')
    .trim();

  // Match floating point or integer
  const match = str.match(/^-?\d+(\.\d+)?/);
  if (match) {
    const num = parseFloat(match[0]);
    return isNaN(num) ? 0 : num;
  }

  return 0;
}

export interface NumericColumnInfo {
  index: number;
  name: string;
  isLikelyQuantity: boolean;
  isLikelyAmount: boolean;
  sampleTotal: number;
}

/**
 * Detects all columns in the sheet that contain numerical data.
 */
export function detectNumericColumns(sheet: SheetData): NumericColumnInfo[] {
  if (!sheet || !sheet.headers || !sheet.rows || sheet.rows.length === 0) {
    return [];
  }

  const results: NumericColumnInfo[] = [];

  sheet.headers.forEach((header, colIndex) => {
    const h = (header || '').toLowerCase().trim();

    // Skip sequence index columns like "م", "#", "id", "index", "no"
    if (h === 'م' || h === '#' || h === 'no' || h === 'id' || h === 'index' || h === 'ت') {
      return;
    }

    // Skip "نوع الصنف" column
    if (h.includes('نوع الصنف') || h.includes('نوع_الصنف') || h.includes('item type')) {
      return;
    }

    // Count numeric cells in this column
    let numericCount = 0;
    let sum = 0;

    sheet.rows.forEach((row) => {
      const cell = row[colIndex];
      if (cell !== undefined && cell !== null && cell !== '') {
        const parsed = parseCleanNumber(cell);
        if (!isNaN(parsed) && parsed !== 0) {
          numericCount++;
          sum += parsed;
        } else if (cell === 0 || cell === '0') {
          numericCount++;
        }
      }
    });

    const isNumericRatio = numericCount / sheet.rows.length;

    // If at least 40% of rows have valid numbers, consider it a numeric column
    if (isNumericRatio >= 0.4 || numericCount >= 2) {
      const isLikelyQuantity =
        /إنتاج|انتاج|كمية|وزن|طن|كجم|ton|qty|quantity|weight|volume|batch/i.test(h);
      const isLikelyAmount =
        /إجمالي|اجمالي|مبلغ|سعر|قيمة|ر\.س|ريال|sar|amount|total|price|cost|value/i.test(h);

      results.push({
        index: colIndex,
        name: header,
        isLikelyQuantity,
        isLikelyAmount,
        sampleTotal: sum,
      });
    }
  });

  return results;
}

export interface CategorySummaryItem {
  category: string;
  label: string;
  icon: string;
  badgeClass: string;
  cardBgClass: string;
  borderColor: string;
  textColor: string;
  barColor: string;
  count: number;
  totalQuantity: number;
  totalAmount: number;
  percentage: number;
  sampleItems: string[];
}

export interface SheetControlTotalResult {
  sheetName: string;
  grandTotalQuantity: number;
  grandTotalAmount: number;
  totalRows: number;
  quantityColumnIndex: number;
  quantityColumnName: string;
  amountColumnIndex: number;
  amountColumnName: string;
  categories: CategorySummaryItem[];
  numericColumns: NumericColumnInfo[];
  formulaColumnIndex: number;
  itemTypeColumnIndex: number;
  crossCheckSum: number;
  isCrossCheckBalanced: boolean;
}

/**
 * Calculates the comprehensive Control Total and the breakdown for each item type.
 */
export function calculateControlTotals(
  sheet: SheetData,
  specifiedQuantityColIndex?: number,
  specifiedAmountColIndex?: number
): SheetControlTotalResult | null {
  if (!sheet || !sheet.headers || !sheet.rows || sheet.rows.length === 0) {
    return null;
  }

  const numericCols = detectNumericColumns(sheet);
  const itemTypeColIndex = findItemTypeColumnIndex(sheet.headers);
  const formulaColIndex = findFormulaColumnIndex(sheet.headers);

  // Determine Quantity / Primary Control column
  let qtyColIndex = -1;
  if (
    specifiedQuantityColIndex !== undefined &&
    specifiedQuantityColIndex >= 0 &&
    specifiedQuantityColIndex < sheet.headers.length
  ) {
    qtyColIndex = specifiedQuantityColIndex;
  } else {
    // Look for explicit quantity/production column first
    const qtyCandidate = numericCols.find((c) => c.isLikelyQuantity);
    if (qtyCandidate) {
      qtyColIndex = qtyCandidate.index;
    } else if (numericCols.length > 0) {
      // Default to first numeric column
      qtyColIndex = numericCols[0].index;
    }
  }

  // Determine Amount / Secondary Financial column
  let amtColIndex = -1;
  if (
    specifiedAmountColIndex !== undefined &&
    specifiedAmountColIndex >= 0 &&
    specifiedAmountColIndex < sheet.headers.length
  ) {
    amtColIndex = specifiedAmountColIndex;
  } else {
    // Look for amount column different from quantity column
    const amtCandidate = numericCols.find(
      (c) => c.isLikelyAmount && c.index !== qtyColIndex
    );
    if (amtCandidate) {
      amtColIndex = amtCandidate.index;
    } else {
      const secondCandidate = numericCols.find((c) => c.index !== qtyColIndex);
      if (secondCandidate) {
        amtColIndex = secondCandidate.index;
      }
    }
  }

  const qtyColName =
    qtyColIndex !== -1 ? sheet.headers[qtyColIndex] : 'الكمية / الإنتاج';
  const amtColName =
    amtColIndex !== -1 ? sheet.headers[amtColIndex] : 'المبلغ / الإجمالي';

  // Structure category tracking
  const categoryMap: Record<
    string,
    {
      count: number;
      totalQuantity: number;
      totalAmount: number;
      sampleItems: string[];
    }
  > = {
    تسمين: { count: 0, totalQuantity: 0, totalAmount: 0, sampleItems: [] },
    سمك: { count: 0, totalQuantity: 0, totalAmount: 0, sampleItems: [] },
    بط: { count: 0, totalQuantity: 0, totalAmount: 0, sampleItems: [] },
    اليف: { count: 0, totalQuantity: 0, totalAmount: 0, sampleItems: [] },
  };

  let grandTotalQuantity = 0;
  let grandTotalAmount = 0;

  sheet.rows.forEach((row) => {
    // Read raw category
    let cat = 'اليف';
    if (itemTypeColIndex !== -1 && row[itemTypeColIndex] !== undefined) {
      const val = (row[itemTypeColIndex] ?? '').toString().trim();
      if (val === 'تسمين' || val === 'سمك' || val === 'بط' || val === 'اليف') {
        cat = val;
      } else if (val) {
        cat = val;
      }
    }

    if (!categoryMap[cat]) {
      categoryMap[cat] = { count: 0, totalQuantity: 0, totalAmount: 0, sampleItems: [] };
    }

    const rowQty = qtyColIndex !== -1 ? parseCleanNumber(row[qtyColIndex]) : 0;
    const rowAmt = amtColIndex !== -1 ? parseCleanNumber(row[amtColIndex]) : 0;

    categoryMap[cat].count += 1;
    categoryMap[cat].totalQuantity += rowQty;
    categoryMap[cat].totalAmount += rowAmt;

    // Collect formula name sample
    if (formulaColIndex !== -1 && row[formulaColIndex]) {
      const formulaName = (row[formulaColIndex] ?? '').toString().trim();
      if (formulaName && !categoryMap[cat].sampleItems.includes(formulaName)) {
        categoryMap[cat].sampleItems.push(formulaName);
      }
    }

    grandTotalQuantity += rowQty;
    grandTotalAmount += rowAmt;
  });

  // Calculate percentages and styling metadata
  const standardCategories = ['تسمين', 'سمك', 'بط', 'اليف'];
  const allCategoryKeys = Object.keys(categoryMap);

  // Ensure standard categories come in a sensible order
  const sortedKeys = [
    ...standardCategories,
    ...allCategoryKeys.filter((k) => !standardCategories.includes(k)),
  ];

  const categoryConfigs: Record<
    string,
    {
      label: string;
      icon: string;
      badgeClass: string;
      cardBgClass: string;
      borderColor: string;
      textColor: string;
      barColor: string;
    }
  > = {
    تسمين: {
      label: 'تسمين (GR - FN - ST - SS)',
      icon: '🌾',
      badgeClass: 'bg-amber-100 text-amber-900 border-amber-300',
      cardBgClass: 'bg-amber-50/60 hover:bg-amber-50',
      borderColor: 'border-amber-200',
      textColor: 'text-amber-950',
      barColor: 'bg-amber-500',
    },
    سمك: {
      label: 'سمك (Tilapia)',
      icon: '🐟',
      badgeClass: 'bg-sky-100 text-sky-900 border-sky-300',
      cardBgClass: 'bg-sky-50/60 hover:bg-sky-50',
      borderColor: 'border-sky-200',
      textColor: 'text-sky-950',
      barColor: 'bg-sky-500',
    },
    بط: {
      label: 'بط (Duck)',
      icon: '🦆',
      badgeClass: 'bg-teal-100 text-teal-900 border-teal-300',
      cardBgClass: 'bg-teal-50/60 hover:bg-teal-50',
      borderColor: 'border-teal-200',
      textColor: 'text-teal-950',
      barColor: 'bg-teal-500',
    },
    اليف: {
      label: 'أليف (أصناف أخرى)',
      icon: '🐾',
      badgeClass: 'bg-purple-100 text-purple-900 border-purple-300',
      cardBgClass: 'bg-purple-50/60 hover:bg-purple-50',
      borderColor: 'border-purple-200',
      textColor: 'text-purple-950',
      barColor: 'bg-purple-500',
    },
  };

  let crossCheckSum = 0;
  const categories: CategorySummaryItem[] = [];

  sortedKeys.forEach((key) => {
    const data = categoryMap[key];
    if (!data) return;

    crossCheckSum += data.totalQuantity;

    const percentage =
      grandTotalQuantity > 0
        ? (data.totalQuantity / grandTotalQuantity) * 100
        : 0;

    const config = categoryConfigs[key] || {
      label: key,
      icon: '📦',
      badgeClass: 'bg-slate-100 text-slate-800 border-slate-300',
      cardBgClass: 'bg-slate-50',
      borderColor: 'border-slate-200',
      textColor: 'text-slate-900',
      barColor: 'bg-slate-500',
    };

    categories.push({
      category: key,
      label: config.label,
      icon: config.icon,
      badgeClass: config.badgeClass,
      cardBgClass: config.cardBgClass,
      borderColor: config.borderColor,
      textColor: config.textColor,
      barColor: config.barColor,
      count: data.count,
      totalQuantity: data.totalQuantity,
      totalAmount: data.totalAmount,
      percentage: Number(percentage.toFixed(2)),
      sampleItems: data.sampleItems,
    });
  });

  const isCrossCheckBalanced =
    Math.abs(crossCheckSum - grandTotalQuantity) < 0.001;

  return {
    sheetName: sheet.sheetName || 'ورقة العمل',
    grandTotalQuantity: Number(grandTotalQuantity.toFixed(2)),
    grandTotalAmount: Number(grandTotalAmount.toFixed(2)),
    totalRows: sheet.rows.length,
    quantityColumnIndex: qtyColIndex,
    quantityColumnName: qtyColName,
    amountColumnIndex: amtColIndex,
    amountColumnName: amtColName,
    categories,
    numericColumns: numericCols,
    formulaColumnIndex: formulaColIndex,
    itemTypeColumnIndex: itemTypeColIndex,
    crossCheckSum: Number(crossCheckSum.toFixed(2)),
    isCrossCheckBalanced,
  };
}

/**
 * Generates formatted textual reconciliation report for WhatsApp or email sharing
 */
export function generateReconciliationTextReport(
  summary: SheetControlTotalResult,
  pdfControlTarget?: number
): string {
  const dateStr = new Date().toLocaleDateString('ar-EG', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });

  let report = `📊 *تقرير مطابقة الكنترول وإجمالي أنواع الأصناف*\n`;
  report += `📅 التاريخ: ${dateStr}\n`;
  report += `📑 ورقة العمل: ${summary.sheetName} (${summary.totalRows} بند / تركيبة)\n`;
  report += `-----------------------------------------\n`;
  report += `🎯 *إجمالي الكنترول المستخرج (${summary.quantityColumnName}):* ${summary.grandTotalQuantity.toLocaleString('ar-EG', { minimumFractionDigits: 2 })}\n`;

  if (summary.amountColumnIndex !== -1 && summary.grandTotalAmount > 0) {
    report += `💰 *إجمالي القيمة المالية (${summary.amountColumnName}):* ${summary.grandTotalAmount.toLocaleString('ar-EG', { minimumFractionDigits: 2 })} ر.س\n`;
  }

  if (pdfControlTarget !== undefined && !isNaN(pdfControlTarget) && pdfControlTarget > 0) {
    const diff = summary.grandTotalQuantity - pdfControlTarget;
    report += `\n🔍 *نتائج المطابقة مع ملف الـ PDF:*\n`;
    report += `• كنترول ملف الـ PDF الأصلي: ${pdfControlTarget.toLocaleString('ar-EG', { minimumFractionDigits: 2 })}\n`;
    report += `• كنترول الجدول المحسوب: ${summary.grandTotalQuantity.toLocaleString('ar-EG', { minimumFractionDigits: 2 })}\n`;

    if (Math.abs(diff) < 0.001) {
      report += `• حالة المطابقة: ✅ متطابق تماماً 100% (فرق المطابقة: 0.00)\n`;
    } else if (diff > 0) {
      report += `• حالة المطابقة: ⚠️ زيادة بمقدار (+${diff.toFixed(2)})\n`;
    } else {
      report += `• حالة المطابقة: ⚠️ عجز بمقدار (${diff.toFixed(2)})\n`;
    }
  }

  report += `\n📦 *تفاصيل إجمالي كل نوع صنف:*\n`;
  summary.categories.forEach((cat) => {
    report += `${cat.icon} *${cat.category}:*\n`;
    report += `   - إجمالي الكمية: ${cat.totalQuantity.toLocaleString('ar-EG', { minimumFractionDigits: 2 })} (${cat.percentage}%)\n`;
    report += `   - عدد التركيبات: ${cat.count}\n`;
    if (cat.totalAmount > 0) {
      report += `   - إجمالي القيمة: ${cat.totalAmount.toLocaleString('ar-EG', { minimumFractionDigits: 2 })} ر.س\n`;
    }
  });

  report += `-----------------------------------------\n`;
  report += `✅ التحقق الرياضي: مجموع الأنواع = ${summary.crossCheckSum.toLocaleString('ar-EG', { minimumFractionDigits: 2 })} (${summary.isCrossCheckBalanced ? 'متوازن 100%' : 'تنبيه في التوازن'})\n`;
  report += `تم الإنشاء عبر محول PDF إلى Excel الذكي.`;

  return report;
}

/**
 * Creates an additional reconciliation Excel sheet to be bundled in the downloaded Excel workbook
 */
export function createReconciliationSheet(
  summary: SheetControlTotalResult,
  pdfControlTarget?: number
): SheetData {
  const headers = [
    'نوع الصنف',
    'البيان والتفصيل',
    `إجمالي ${summary.quantityColumnName}`,
    'النسبة المئوية (%)',
    'عدد التركيبات / البنود',
    `إجمالي ${summary.amountColumnName}`,
    'ملاحظات المطابقة',
  ];

  const rows: (string | number)[][] = summary.categories.map((cat) => [
    cat.category,
    cat.label,
    cat.totalQuantity,
    `${cat.percentage}%`,
    cat.count,
    cat.totalAmount > 0 ? cat.totalAmount : '',
    cat.sampleItems.slice(0, 3).join('، '),
  ]);

  const diff =
    pdfControlTarget !== undefined && !isNaN(pdfControlTarget)
      ? summary.grandTotalQuantity - pdfControlTarget
      : 0;

  const matchStatus =
    pdfControlTarget !== undefined && !isNaN(pdfControlTarget)
      ? Math.abs(diff) < 0.001
        ? 'متطابق 100% مع الـ PDF'
        : `يوجد فرق (${diff > 0 ? '+' : ''}${diff.toFixed(2)})`
      : 'تم الحساب آلياً';

  const totalsRow: (string | number)[] = [
    'إجمالي الكنترول العام',
    'مجموع كافة الأصناف',
    summary.grandTotalQuantity,
    '100.00%',
    summary.totalRows,
    summary.grandTotalAmount > 0 ? summary.grandTotalAmount : '',
    matchStatus,
  ];

  return {
    sheetName: 'مطابقة الكنترول والأنواع',
    description: 'تقرير مطابقة إجمالي الكنترول وتفاصيل كل نوع صنف مع ملف الـ PDF',
    isRtl: true,
    headers,
    rows,
    totalsRow,
  };
}
