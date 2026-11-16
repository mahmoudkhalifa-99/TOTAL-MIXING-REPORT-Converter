import { SheetData } from '../types/converter';
import { findItemTypeColumnIndex, findFormulaColumnIndex, ItemType } from './itemClassifier';

export interface CategorySummaryItem {
  category: ItemType;
  label: string;
  emoji: string;
  color: {
    bg: string;
    border: string;
    text: string;
    badgeBg: string;
    ring: string;
    progress: string;
  };
  count: number;
  percentOfTotalCount: number;
  targetSum: number;
  actualSum: number;
  variance: number;
  variancePct: number;
  actualPercentOfGrand: number;
  targetPercentOfGrand: number;
  hasTargetData: boolean;
  hasActualData: boolean;
  sampleItems: Array<{
    product: string;
    ingredient: string;
    code: string;
    target: number | string;
    actual: number | string;
  }>;
}

export interface OverallTotalsSummary {
  hasItemTypeColumn: boolean;
  hasNumericColumns: boolean;
  sheetName: string;
  targetColName: string | null;
  targetColIndex: number;
  actualColName: string | null;
  actualColIndex: number;
  weightUnit: string;
  categories: {
    'سمك': CategorySummaryItem;
    'تسمين': CategorySummaryItem;
    'بط': CategorySummaryItem;
    'اليف': CategorySummaryItem;
  };
  grandTotal: {
    count: number;
    targetSum: number;
    actualSum: number;
    variance: number;
    variancePct: number;
  };
}

/**
 * Parses numeric cell value cleanly
 */
export function parseNumericCell(val: any): number {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  let s = val
    .toString()
    .replace(/[٠-٩٫٬]/g, (ch: string) => {
      const idx = '٠١٢٣٤٥٦٧٨٩'.indexOf(ch);
      if (idx !== -1) return String(idx);
      return ch === '٫' ? '.' : ch === '٬' ? ',' : ch;
    })
    .replace(/(?:kg|كجم|طن|ton|tons|gm|جرام|ر\.س|SAR|EGP|USD)/gi, '')
    .trim();

  if (!s) return 0;

  // Handle accounting parentheses negative values e.g. "(1,234.50)"
  const parenNeg = s.match(/^\(\s*([+-]?[\d,]+(?:\.\d+)?)\s*\)$/);
  if (parenNeg) s = `-${parenNeg[1]}`;

  // Remove thousands commas inside digits without gluing space-separated numbers together
  s = s.replace(/,/g, '');

  // If multiple space-separated tokens appear unexpectedly, parse the primary numeric token cleanly
  const firstTokenMatch = s.match(/^[+-]?\d+(?:\.\d+)?/);
  if (firstTokenMatch) {
    const n = parseFloat(firstTokenMatch[0]);
    return isNaN(n) ? 0 : n;
  }

  const anyNumMatch = s.match(/[+-]?\d+(?:\.\d+)?/);
  if (!anyNumMatch) return 0;
  const n = parseFloat(anyNumMatch[0]);
  return isNaN(n) ? 0 : n;
}

/**
 * Formats a number with Arabic locale / standard decimal commas
 */
export function formatWeight(num: number, unit: string = 'كجم', decimals: number = 3): string {
  if (num === null || num === undefined || isNaN(num)) return `0.000 ${unit}`.trim();
  const formatted = num.toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
  return `${formatted} ${unit}`.trim();
}

const EXCLUDED_WEIGHT_HEADER_RE =
  /diff|variance|dev|فرق|انحراف|%|percent|ratio|price|سعر|cost|تكلفة|كود|code|ident|supledger|subledger|batch|رقم|نوع الصنف|item_type/i;

/**
 * Detects target and actual numeric columns in a sheet
 */
export function detectWeightColumns(headers: string[], rows: any[][]): {
  targetColIndex: number;
  targetColName: string | null;
  actualColIndex: number;
  actualColName: string | null;
  unit: string;
} {
  let targetColIndex = -1;
  let targetColName: string | null = null;
  let actualColIndex = -1;
  let actualColName: string | null = null;
  let detectedUnit = 'كجم';

  // 1. Look by column header keywords
  headers.forEach((h, idx) => {
    const norm = (h || '').trim().toLowerCase();
    if (!norm) return;
    const isExcluded = EXCLUDED_WEIGHT_HEADER_RE.test(norm);

    if (
      !isExcluded &&
      /target_kg|target|planned|plan_qty|plan\b|std_qty|المستهدف|الكمية المستهدفة|الوزن المستهدف/i.test(
        norm
      )
    ) {
      targetColIndex = idx;
      targetColName = h;
    }
    if (
      !isExcluded &&
      /actual_kg|actual|act_qty|act\b|consumed|dispensed|used|الفعلي|الكمية الفعلية|الوزن الفعلي|المستهلك|الاستهلاك|الكمية|sum of|quantity|qty|weight/i.test(
        norm
      ) &&
      !norm.includes('target') &&
      !norm.includes('مستهدف')
    ) {
      actualColIndex = idx;
      actualColName = h;
    }
    if (/طن|ton/i.test(norm)) {
      detectedUnit = 'طن';
    } else if (/جرام|gm|gram/i.test(norm)) {
      detectedUnit = 'جم';
    }
  });

  // 2. If target / actual not matched by exact name, inspect non-excluded numeric columns
  if (targetColIndex === -1 || actualColIndex === -1) {
    const numericCandidates: { idx: number; hasDecimals: boolean }[] = [];

    for (let c = 0; c < headers.length; c++) {
      if (c === targetColIndex || c === actualColIndex) continue;
      const normHeader = (headers[c] || '').trim().toLowerCase();
      if (
        normHeader === 'م' ||
        normHeader === '#' ||
        normHeader === 'id' ||
        normHeader === 'no' ||
        EXCLUDED_WEIGHT_HEADER_RE.test(normHeader)
      ) {
        continue;
      }

      let numericCount = 0;
      let decimalCount = 0;
      let checkedCount = 0;

      for (let r = 0; r < Math.min(rows.length, 25); r++) {
        const val = rows[r]?.[c];
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          checkedCount++;
          const strVal = String(val).trim();
          const parsed = parseNumericCell(val);
          if (!isNaN(parsed) && /^[(+-]?[\d,]+(?:\.\d+)?\)?$/.test(strVal)) {
            numericCount++;
            if (/\.\d+/.test(strVal)) decimalCount++;
          }
        }
      }

      if (checkedCount > 0 && numericCount / checkedCount >= 0.7) {
        numericCandidates.push({ idx: c, hasDecimals: decimalCount > 0 });
      }
    }

    const decimalCandidates = numericCandidates.filter((c) => c.hasDecimals);
    const pool = decimalCandidates.length > 0 ? decimalCandidates : numericCandidates;

    if (actualColIndex === -1 && pool.length > 0) {
      actualColIndex = pool[pool.length - 1].idx;
      actualColName = headers[actualColIndex];
    }
    if (targetColIndex === -1 && pool.length > 1) {
      targetColIndex = pool[0].idx;
      targetColName = headers[targetColIndex];
    }
  }

  return {
    targetColIndex,
    targetColName,
    actualColIndex,
    actualColName,
    unit: detectedUnit,
  };
}

/**
 * Calculates complete summary totals for the four categories:
 * سمك، تسمين، بط، اليف + الإجمالي الكلي للأربعة
 */
export function calculateCategoryTotals(sheet: SheetData): OverallTotalsSummary {
  const itemTypeColIdx = findItemTypeColumnIndex(sheet.headers);
  const formulaColIdx = findFormulaColumnIndex(sheet.headers);

  // Ingredient Name column index
  const ingredientColIdx = sheet.headers.findIndex((h) =>
    /ingredient_name|اسم المكون|المكون|اسم الصنف|item_name/i.test(h.trim().toLowerCase())
  );
  // Code column index (prefer densest Ingredient_Code / كود الصنف column over Product_Code or sparse duplicate)
  const codeCandidates = sheet.headers
    .map((h, idx) => ({ h: h.trim().toLowerCase(), idx }))
    .filter(
      ({ h }) =>
        /code|كود|رمز/i.test(h) &&
        !h.includes('product') &&
        !h.includes('تركيبة') &&
        !h.includes('supledger') &&
        !h.includes('subledger')
    );
  let codeColIdx = -1;
  let bestCodeCount = -1;
  for (const { idx } of codeCandidates) {
    const count = sheet.rows.filter((r) => {
      const v = (r?.[idx] ?? '').toString().trim();
      return v && v !== '-' && v !== '—';
    }).length;
    if (count >= bestCodeCount) {
      bestCodeCount = count;
      codeColIdx = idx;
    }
  }

  const { targetColIndex, targetColName, actualColIndex, actualColName, unit } =
    detectWeightColumns(sheet.headers, sheet.rows);

  const categoriesConfig: Record<
    ItemType,
    {
      label: string;
      emoji: string;
      color: CategorySummaryItem['color'];
    }
  > = {
    'سمك': {
      label: 'السمك (Tilapia)',
      emoji: '🐟',
      color: {
        bg: 'bg-sky-50',
        border: 'border-sky-200',
        text: 'text-sky-900',
        badgeBg: 'bg-sky-100 text-sky-800 border-sky-300',
        ring: 'ring-sky-400',
        progress: 'bg-sky-500',
      },
    },
    'تسمين': {
      label: 'التسمين (GR, FN, ST, SS)',
      emoji: '🌾',
      color: {
        bg: 'bg-amber-50',
        border: 'border-amber-200',
        text: 'text-amber-950',
        badgeBg: 'bg-amber-100 text-amber-900 border-amber-300',
        ring: 'ring-amber-400',
        progress: 'bg-amber-500',
      },
    },
    'بط': {
      label: 'البط (Duck)',
      emoji: '🦆',
      color: {
        bg: 'bg-teal-50',
        border: 'border-teal-200',
        text: 'text-teal-900',
        badgeBg: 'bg-teal-100 text-teal-800 border-teal-300',
        ring: 'ring-teal-400',
        progress: 'bg-teal-500',
      },
    },
    'اليف': {
      label: 'الأليف (أصناف أخرى)',
      emoji: '🐾',
      color: {
        bg: 'bg-purple-50',
        border: 'border-purple-200',
        text: 'text-purple-900',
        badgeBg: 'bg-purple-100 text-purple-800 border-purple-300',
        ring: 'ring-purple-400',
        progress: 'bg-purple-500',
      },
    },
  };

  const initialCat = (key: ItemType): CategorySummaryItem => ({
    category: key,
    label: categoriesConfig[key].label,
    emoji: categoriesConfig[key].emoji,
    color: categoriesConfig[key].color,
    count: 0,
    percentOfTotalCount: 0,
    targetSum: 0,
    actualSum: 0,
    variance: 0,
    variancePct: 0,
    actualPercentOfGrand: 0,
    targetPercentOfGrand: 0,
    hasTargetData: targetColIndex !== -1,
    hasActualData: actualColIndex !== -1,
    sampleItems: [],
  });

  const categories: Record<ItemType, CategorySummaryItem> = {
    'سمك': initialCat('سمك'),
    'تسمين': initialCat('تسمين'),
    'بط': initialCat('بط'),
    'اليف': initialCat('اليف'),
  };

  let grandCount = 0;
  let grandTargetSum = 0;
  let grandActualSum = 0;

  // Iterate all rows
  sheet.rows.forEach((row) => {
    let catKey: ItemType = 'اليف';

    if (itemTypeColIdx !== -1) {
      const val = (row[itemTypeColIdx] ?? '').toString().trim();
      if (val === 'سمك' || val === 'تسمين' || val === 'بط' || val === 'اليف') {
        catKey = val as ItemType;
      }
    }

    const targetVal = targetColIndex !== -1 ? parseNumericCell(row[targetColIndex]) : 0;
    const actualVal = actualColIndex !== -1 ? parseNumericCell(row[actualColIndex]) : 0;

    categories[catKey].count += 1;
    categories[catKey].targetSum += targetVal;
    categories[catKey].actualSum += actualVal;

    grandCount += 1;
    grandTargetSum += targetVal;
    grandActualSum += actualVal;

    // Collect sample item info (product, ingredient, code)
    if (categories[catKey].sampleItems.length < 8) {
      categories[catKey].sampleItems.push({
        product: formulaColIdx !== -1 ? (row[formulaColIdx] ?? '').toString() : '',
        ingredient: ingredientColIdx !== -1 ? (row[ingredientColIdx] ?? '').toString() : '',
        code: codeColIdx !== -1 ? (row[codeColIdx] ?? '').toString() : '',
        target: targetVal,
        actual: actualVal,
      });
    }
  });

  // Calculate percentages and variances
  const catKeys: ItemType[] = ['سمك', 'تسمين', 'بط', 'اليف'];
  catKeys.forEach((k) => {
    const item = categories[k];
    item.variance = item.actualSum - item.targetSum;
    item.variancePct =
      item.targetSum > 0 ? (item.variance / item.targetSum) * 100 : 0;
    item.percentOfTotalCount =
      grandCount > 0 ? (item.count / grandCount) * 100 : 0;
    item.actualPercentOfGrand =
      grandActualSum > 0 ? (item.actualSum / grandActualSum) * 100 : 0;
    item.targetPercentOfGrand =
      grandTargetSum > 0 ? (item.targetSum / grandTargetSum) * 100 : 0;
  });

  const grandVariance = grandActualSum - grandTargetSum;
  const grandVariancePct =
    grandTargetSum > 0 ? (grandVariance / grandTargetSum) * 100 : 0;

  return {
    hasItemTypeColumn: itemTypeColIdx !== -1,
    hasNumericColumns: targetColIndex !== -1 || actualColIndex !== -1,
    sheetName: sheet.sheetName || 'ورقة العمل',
    targetColName,
    targetColIndex,
    actualColName,
    actualColIndex,
    weightUnit: unit,
    categories,
    grandTotal: {
      count: grandCount,
      targetSum: grandTargetSum,
      actualSum: grandActualSum,
      variance: grandVariance,
      variancePct: grandVariancePct,
    },
  };
}

/**
 * Generates an executive WhatsApp/Email copyable text summary
 */
export function generateTotalsReportText(
  summary: OverallTotalsSummary,
  docTitle: string = 'تقرير استهلاك وخلطات الأعلاف'
): string {
  const { categories, grandTotal, weightUnit, sheetName, targetColName, actualColName } = summary;

  const lines: string[] = [
    `📊 *ملخص إجماليات الأصناف والأوزان - ${docTitle}*`,
    `📄 ورقة العمل: ${sheetName}`,
    `━━━━━━━━━━━━━━━━━━━━━━━━━`,
  ];

  const order: ItemType[] = ['سمك', 'تسمين', 'بط', 'اليف'];

  order.forEach((k) => {
    const cat = categories[k];
    lines.push(`${cat.emoji} *إجمالي ${k} (${cat.count} صنف):*`);
    if (actualColName) {
      lines.push(`   • الفعلي (${actualColName}): ${formatWeight(cat.actualSum, weightUnit)}`);
    }
    if (targetColName) {
      lines.push(`   • المستهدف (${targetColName}): ${formatWeight(cat.targetSum, weightUnit)}`);
      const diffSign = cat.variance >= 0 ? '+' : '';
      lines.push(`   • الفرق: ${diffSign}${formatWeight(cat.variance, weightUnit)} (${diffSign}${cat.variancePct.toFixed(2)}%)`);
    }
    lines.push(`   • نسبة المساهمة من الإجمالي: ${cat.actualPercentOfGrand.toFixed(1)}%`);
    lines.push('');
  });

  lines.push(`━━━━━━━━━━━━━━━━━━━━━━━━━`);
  lines.push(`⭐ *الإجمالي الكلي للأربعة (Grand Total):*`);
  lines.push(`   • إجمالي عدد الأصناف / الصفوف: ${grandTotal.count} صنف`);
  if (actualColName) {
    lines.push(`   • إجمالي الفعلي العام: ${formatWeight(grandTotal.actualSum, weightUnit)}`);
  }
  if (targetColName) {
    lines.push(`   • إجمالي المستهدف العام: ${formatWeight(grandTotal.targetSum, weightUnit)}`);
    const netSign = grandTotal.variance >= 0 ? '+' : '';
    lines.push(`   • صافي الفرق الكلي: ${netSign}${formatWeight(grandTotal.variance, weightUnit)} (${netSign}${grandTotal.variancePct.toFixed(2)}%)`);
  }
  lines.push(`━━━━━━━━━━━━━━━━━━━━━━━━━`);
  lines.push(`✓ تم الإعداد والاستخراج آلياً بنظام محول PDF إلى Excel الذكي`);

  return lines.join('\n');
}
