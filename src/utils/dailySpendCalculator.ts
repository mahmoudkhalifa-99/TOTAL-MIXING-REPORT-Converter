import { SheetData } from '../types/converter';
import { findItemTypeColumnIndex, findFormulaColumnIndex } from './itemClassifier';
import { detectWeightColumns, parseNumericCell } from './categoryTotals';
import {
  findIngredientNameColumnIndex,
  findIngredientCodeColumns,
  isSummaryOrTotalRow,
  extractDateFromText,
} from './sadatControlCalculator';
import {
  MASTER_DRIVE_CODES_CATALOG,
  lookupDriveCatalogInfo,
  normalizeItemText,
} from './driveCodesCatalog';
import { loadStoredDriveOverrides, loadStoredCustomCatalog } from './controlPivotCalculator';

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
  /** Optional manual item (JDE) code / drive code typed by the user for a custom item */
  code?: string;
  driveCode?: string;
  custom?: boolean;
}

/** Items the user added himself (persisted in the browser) */
export interface DailySpendCustomItem {
  en: string;
  ar: string;
  code?: string;
  driveCode?: string;
}

/** Per-line manual edits of the codes (key = English name of the line) */
export type DailySpendCodeOverrides = Record<string, { code?: string; driveCode?: string }>;

export const DAILY_SPEND_CUSTOM_ITEMS_KEY = 'daily_spend_custom_items_v1';
export const DAILY_SPEND_CODE_OVERRIDES_KEY = 'daily_spend_code_overrides_v1';

export function loadDailySpendCustomItems(): DailySpendCustomItem[] {
  try {
    const raw = typeof localStorage !== 'undefined' ? localStorage.getItem(DAILY_SPEND_CUSTOM_ITEMS_KEY) : null;
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed.filter((x) => x && typeof x.en === 'string' && x.en.trim()) : [];
  } catch {
    return [];
  }
}

export function loadDailySpendCodeOverrides(): DailySpendCodeOverrides {
  try {
    const raw = typeof localStorage !== 'undefined' ? localStorage.getItem(DAILY_SPEND_CODE_OVERRIDES_KEY) : null;
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
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
  /** Item (JDE) code and Drive code, resolved from: manual edit > custom item > file > catalog */
  code: string;
  driveCode: string;
  custom: boolean;
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
  documentTitle?: string,
  customItems: DailySpendCustomItem[] = [],
  codeOverrides: DailySpendCodeOverrides = {}
): DailySpendReport {
  const allItems: DailySpendItemDef[] = [
    ...DAILY_SPEND_ITEMS,
    ...customItems
      .filter((c) => c.en && c.en.trim())
      .map<DailySpendItemDef>((c) => ({
        en: c.en.trim(),
        ar: (c.ar || '').trim(),
        aliases: [c.en.trim()],
        code: c.code,
        driveCode: c.driveCode,
        custom: true,
      })),
  ];

  const rowsOut: DailySpendRow[] = allItems.map((d, i) => ({
    index: i,
    en: d.en,
    ar: d.ar,
    qty: 0,
    hasData: false,
    code: '',
    driveCode: '',
    custom: !!d.custom,
  }));
  const fileCodes: string[] = allItems.map(() => '');

  const aliasMap = new Map<string, number>();
  allItems.forEach((d, i) => {
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

  const codeColIndices = findIngredientCodeColumns(sheet.headers, sheet.rows);
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
      if (!fileCodes[idx]) {
        for (const cIdx of codeColIndices) {
          const v = (row[cIdx] ?? '').toString().trim().replace(/,/g, '');
          if (/^\d{3,}$/.test(v)) {
            fileCodes[idx] = v;
            break;
          }
        }
      }
    } else {
      others.set(name, (others.get(name) ?? 0) + qty);
    }
  }

  // Resolve item (JDE) code + Drive code for every line
  const driveOverrides = loadStoredDriveOverrides();
  const customCatalog = loadStoredCustomCatalog();
  const combinedCatalog = [...customCatalog, ...MASTER_DRIVE_CODES_CATALOG];
  rowsOut.forEach((r, i) => {
    const def = allItems[i];
    const manual = codeOverrides[def.en] || {};

    let autoDrive = '';
    let autoJde = fileCodes[i] || '';
    // Variant lines such as "Corn B" fall back to the base item's codes when they have none of their own
    for (const nameToTry of [def.en, def.en.replace(/\s+b$/i, ''), def.ar]) {
      if (!nameToTry) continue;
      const info = lookupDriveCatalogInfo(nameToTry, autoJde, driveOverrides, customCatalog);
      if (info.driveCode) {
        autoDrive = info.driveCode;
        if (!autoJde) {
          const hit =
            combinedCatalog.find((e) => normalizeItemText(e.name) === normalizeItemText(info.fullName)) ||
            combinedCatalog.find((e) => e.driveCode === info.driveCode);
          if (hit) autoJde = hit.jdeCode || '';
        }
        break;
      }
    }

    r.code = (manual.code ?? def.code ?? autoJde ?? '').toString().trim();
    r.driveCode = (manual.driveCode ?? def.driveCode ?? autoDrive ?? '').toString().trim();
  });

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
