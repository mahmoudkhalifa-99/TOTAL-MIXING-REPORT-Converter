import { SheetData } from '../types/converter';
import { findItemTypeColumnIndex, findFormulaColumnIndex } from './itemClassifier';
import { detectWeightColumns, parseNumericCell } from './categoryTotals';
import {
  findIngredientNameColumnIndex,
  isSummaryOrTotalRow,
  extractDateFromText,
} from './sadatControlCalculator';

/**
 * The fixed list of items shown in the "الصرف اليومي للكنترول" daily sheet, in its exact order.
 * `aliases` are the names that may appear in the uploaded report for that line (compared ignoring
 * case, spaces and punctuation). Corn / Corn B and Local soya 46% / Local soya 46% B are deliberately
 * SEPARATE lines here, exactly like the daily sheet.
 * To add or reorder a line, just edit this list.
 */
export interface DailySpendItemDef {
  en: string;
  ar: string;
  aliases: string[];
}

export const DAILY_SPEND_ITEMS: DailySpendItemDef[] = [
  { en: 'Corn', ar: 'الذرة', aliases: ['Corn', 'Yellow Corn'] },
  { en: 'local white corn (1)', ar: 'الذرة بيضاء', aliases: ['local white corn (1)', 'local white corn'] },
  { en: 'Local yellow corn', ar: 'الذرة صفراء محلى', aliases: ['Local yellow corn'] },
  { en: 'Sorghum bicolor', ar: 'الذرة سورجم', aliases: ['Sorghum bicolor', 'Sorghum'] },
  { en: 'Local soya 46%', ar: 'صويا 46%', aliases: ['Local soya 46%', 'Soya 46%', 'SOYA 46%'] },
  { en: 'Soya Oil', ar: 'زيت صويا', aliases: ['Soya Oil'] },
  { en: 'soya 44%', ar: 'صويا 44%', aliases: ['soya 44%'] },
  { en: 'HI Fat Soya 42%', ar: 'صويا هاى فات', aliases: ['HI Fat Soya 42%', 'Hi Fat Soya'] },
  { en: 'Soya full fat', ar: 'فولفات', aliases: ['Soya full fat', 'Soya full fat 35%'] },
  { en: 'Tunisian DCP', ar: 'دايكال تونسى', aliases: ['Tunisian DCP'] },
  { en: 'DCP', ar: 'دايكال تركى', aliases: ['DCP'] },
  { en: 'Soya B', ar: 'صويا', aliases: ['Soya B'] },
  { en: 'Turkish mono', ar: 'مونو تركى', aliases: ['Turkish mono'] },
  { en: 'Wheat Screenings', ar: 'خليط زوائد', aliases: ['Wheat Screenings'] },
  { en: 'Sunflower O/C 36%', ar: 'كسب عباد', aliases: ['Sunflower O/C 36%', 'Sunflower O/C'] },
  { en: 'Bran', ar: 'ردة', aliases: ['Bran', 'Bran 14%'] },
  { en: 'Water', ar: 'ماء', aliases: ['Water'] },
  { en: 'Ukrainian Corn 2', ar: 'الذرة اوكرانى', aliases: ['Ukrainian Corn 2', 'Ukrainian Corn'] },
  { en: 'Imported soya 46%', ar: 'صويا 46%', aliases: ['Imported soya 46%'] },
  { en: 'Corn B', ar: 'الذرة', aliases: ['Corn B'] },
  { en: 'Limistone', ar: 'حجر جيرى', aliases: ['Limistone', 'Limestone'] },
  { en: 'Rice Bran 14%', ar: 'رجيع الكون', aliases: ['Rice Bran 14%', 'Rice Bran'] },
  { en: 'MONO CALCIUM', ar: 'مونو كالسيوم', aliases: ['MONO CALCIUM', 'Mono calcium'] },
  { en: 'Local soya 46% B', ar: 'صويا', aliases: ['Local soya 46% B', 'SOYA 46% B', 'Soya 46% B'] },
  { en: 'Poultry Meal 60% Dark', ar: 'مسحوق بروتين', aliases: ['Poultry Meal 60% Dark'] },
  { en: 'Poultry meal feathers', ar: 'بروتين معصور', aliases: ['Poultry meal feathers'] },
  { en: 'Soya O/C 46% (D)', ar: 'Soya O/C 46% (D)', aliases: ['Soya O/C 46% (D)', 'Soya O/C 46% D'] },
];

/** Lower-case, letters+digits only, so "SOYA 46% B" == "soya 46 b" */
function compact(s: string): string {
  return (s ?? '')
    .toString()
    .toLowerCase()
    .replace(/[^a-z0-9\u0600-\u06ff]+/g, '');
}

export interface DailySpendRow {
  index: number;
  en: string;
  ar: string;
  qty: number;
  hasData: boolean;
}

export interface DailySpendReport {
  rows: DailySpendRow[];
  total: number;
  /** Items found in the file that are NOT in the daily list (name -> qty) */
  otherItems: { name: string; qty: number }[];
  otherTotal: number;
  fileTotal: number;
  unit: string;
  detectedDate: string;
}

export function calculateDailySpend(
  sheet: SheetData,
  metric: 'actual' | 'target',
  divisor: number,
  documentTitle?: string
): DailySpendReport {
  const rowsOut: DailySpendRow[] = DAILY_SPEND_ITEMS.map((d, i) => ({
    index: i,
    en: d.en,
    ar: d.ar,
    qty: 0,
    hasData: false,
  }));

  const aliasMap = new Map<string, number>();
  DAILY_SPEND_ITEMS.forEach((d, i) => {
    d.aliases.forEach((a) => {
      const k = compact(a);
      if (k && !aliasMap.has(k)) aliasMap.set(k, i);
    });
  });

  const itemTypeColIdx = findItemTypeColumnIndex(sheet.headers);
  const formulaColIdx = findFormulaColumnIndex(sheet.headers);
  const nameColIdx = findIngredientNameColumnIndex(sheet.headers, formulaColIdx, itemTypeColIdx);
  const { targetColIndex, actualColIndex, unit } = detectWeightColumns(sheet.headers, sheet.rows);
  const chosen =
    metric === 'actual'
      ? actualColIndex !== -1
        ? actualColIndex
        : targetColIndex
      : targetColIndex !== -1
      ? targetColIndex
      : actualColIndex;

  const others = new Map<string, number>();
  let fileTotal = 0;

  for (const row of sheet.rows) {
    if (isSummaryOrTotalRow(row)) continue;
    const name = nameColIdx !== -1 ? (row[nameColIdx] ?? '').toString().trim() : '';
    if (!name) continue;
    const rawQty = chosen !== -1 ? parseNumericCell(row[chosen]) : 0;
    const qty = divisor > 0 ? rawQty / divisor : rawQty;
    fileTotal += qty;

    const idx = aliasMap.get(compact(name));
    if (idx !== undefined) {
      rowsOut[idx].qty += qty;
      rowsOut[idx].hasData = true;
    } else {
      others.set(name, (others.get(name) ?? 0) + qty);
    }
  }

  const otherItems = Array.from(others.entries())
    .map(([name, qty]) => ({ name, qty }))
    .sort((a, b) => b.qty - a.qty);
  const otherTotal = otherItems.reduce((a, x) => a + x.qty, 0);
  const total = rowsOut.reduce((a, r) => a + r.qty, 0);

  return {
    rows: rowsOut,
    total,
    otherItems,
    otherTotal,
    fileTotal,
    unit,
    detectedDate: extractDateFromText(documentTitle || sheet?.sheetName),
  };
}
