import { SheetData } from '../types/converter';
import { findItemTypeColumnIndex, findFormulaColumnIndex, ItemType } from './itemClassifier';
import { detectWeightColumns, parseNumericCell } from './categoryTotals';
import { extractCodeFromValue } from './codeSplitter';
import {
  lookupDriveCode,
  DriveCodeEntry,
  DRIVE_OVERRIDES_STORAGE_KEY,
  DRIVE_CUSTOM_CATALOG_STORAGE_KEY,
  groupItemByName,
  normalizeItemText as normalizeItemTextForGroup,
} from './driveCodesCatalog';

export interface ControlPivotRow {
  driveCode: string; // كود الدريف
  code: string;      // كود الصنف (JDE)
  name: string;
  fattening: number; // تسمين
  fish: number;      // منصرف سمك
  duck: number;      // منصرف بط
  pets: number;      // حيوانات أليفة
  total: number;     // الاجمالى
  occurrenceCount: number;
}

export interface ControlPivotSummary {
  rows: ControlPivotRow[];
  totals: {
    fattening: number;
    fish: number;
    duck: number;
    pets: number;
    grandTotal: number;
  };
  unit: string;
  hasActualData: boolean;
  hasTargetData: boolean;
  actualColName: string | null;
  targetColName: string | null;
  metricUsed: 'actual' | 'target';
  /** Total of Cleaning groups, kept OUT of rows/totals (own category) */
  cleaningTotal?: number;
}

function isSummaryOrTotalRow(row: any[]): boolean {
  const joined = row.map((c) => (c ?? '').toString().trim()).join(' ');
  return (
    joined.includes('إجمالي السمك') ||
    joined.includes('إجمالي التسمين') ||
    joined.includes('إجمالي البط') ||
    joined.includes('إجمالي الأليف') ||
    joined.includes('الإجمالي الكلي للأربعة') ||
    joined.includes('الإجمالي العام')
  );
}

function isCodeOrIdHeader(h: string): boolean {
  return (
    h.includes('code') ||
    h.includes('كود') ||
    h.includes('رمز') ||
    h.includes('رقم') ||
    h.includes('supledger') ||
    h.includes('subledger') ||
    h.includes('ident') ||
    h.includes('sku') ||
    h === 'id' ||
    h === 'م' ||
    h === '#' ||
    h === 'no' ||
    h.includes('نوع الصنف')
  );
}

/**
 * Finds the best Ingredient Name column index (prefers Ingredient_Name / اسم الصنف over Finished_Product).
 */
function findIngredientNameColumnIndex(headers: string[], formulaColIdx: number, itemTypeColIdx: number): number {
  const exactPriority = [
    'ingredient_name',
    'ingredient name',
    'ingredient',
    'اسم الصنف',
    'إسم الصنف',
    'اسم المكون',
    'المكون',
    'item_name',
    'item name',
    'material',
    'raw material',
    'الخامة',
    'الصنف',
  ];

  for (let i = 0; i < headers.length; i++) {
    if (i === itemTypeColIdx) continue;
    const h = (headers[i] ?? '').toString().trim().toLowerCase();
    if (isCodeOrIdHeader(h)) continue;
    if (exactPriority.some((p) => h === p || h.includes(p))) {
      return i;
    }
  }

  // Fallback to formula column or first non-code/non-type text column
  if (formulaColIdx !== -1) return formulaColIdx;
  for (let i = 0; i < headers.length; i++) {
    if (i === itemTypeColIdx) continue;
    const h = (headers[i] ?? '').toString().trim().toLowerCase();
    if (!isCodeOrIdHeader(h) && !h.includes('target') && !h.includes('actual') && !h.includes('المستهدف') && !h.includes('الفعلي')) {
      return i;
    }
  }
  return 0;
}

/**
 * Finds all Ingredient Code columns (prioritizing Ingredient_Code / كود الصنف over Product_Code)
 * and ranks columns within the same priority tier by populated code density descending
 * so a sparse accidental first code column never shadows the real codes column.
 */
function findIngredientCodeColumns(headers: string[], rows: any[][]): number[] {
  const primary: number[] = [];
  const secondary: number[] = [];

  headers.forEach((header, idx) => {
    const h = (header ?? '').toString().trim().toLowerCase();
    if (
      h.includes('product') ||
      h.includes('تركيبة') ||
      h.includes('supledger') ||
      h.includes('subledger') ||
      h.includes('drive') ||
      h.includes('دريف')
    ) {
      return;
    }
    if (
      h === 'ingredient_code' ||
      h === 'ingredient code' ||
      h === 'كود الصنف' ||
      h === 'كود_الصنف' ||
      h === 'رمز الصنف' ||
      h === 'كود المكون' ||
      h === 'item_code' ||
      h === 'item code' ||
      h === 'code' ||
      h === 'الكود'
    ) {
      primary.push(idx);
    } else if (h.includes('code') || h.includes('كود') || h.includes('رمز')) {
      secondary.push(idx);
    }
  });

  const countValidCodes = (colIdx: number) => {
    let count = 0;
    for (const row of rows) {
      const val = (row?.[colIdx] ?? '').toString().trim();
      if (val && val !== '-' && val !== '—') count++;
    }
    return count;
  };

  const sortByDensityDesc = (indices: number[]) =>
    [...indices].sort((a, b) => {
      const diff = countValidCodes(b) - countValidCodes(a);
      return diff !== 0 ? diff : b - a;
    });

  return [...sortByDensityDesc(primary), ...sortByDensityDesc(secondary)];
}

export type ControlPivotColumnKey =
  | 'driveCode'
  | 'code'
  | 'name'
  | 'fattening'
  | 'fish'
  | 'duck'
  | 'pets'
  | 'total';

export interface PivotColumnDef {
  key: ControlPivotColumnKey;
  label: string;
  width: number;
  isNumeric: boolean;
}

export const PIVOT_COLUMN_DEFS: Record<ControlPivotColumnKey, PivotColumnDef> = {
  driveCode: { key: 'driveCode', label: 'كود الدريف', width: 14, isNumeric: false },
  code: { key: 'code', label: 'كود الصنف', width: 15, isNumeric: false },
  name: { key: 'name', label: 'إسم الصنف', width: 30, isNumeric: false },
  fattening: { key: 'fattening', label: 'تسمين', width: 18, isNumeric: true },
  fish: { key: 'fish', label: 'منصرف سمك', width: 18, isNumeric: true },
  duck: { key: 'duck', label: 'منصرف بط', width: 18, isNumeric: true },
  pets: { key: 'pets', label: 'حيوانات أليفة', width: 18, isNumeric: true },
  total: { key: 'total', label: 'الاجمالى', width: 20, isNumeric: true },
};

export const DEFAULT_PIVOT_COLUMN_ORDER: ControlPivotColumnKey[] = [
  'driveCode',
  'code',
  'name',
  'fattening',
  'fish',
  'duck',
  'pets',
  'total',
];

export const PIVOT_COLUMN_ORDER_STORAGE_KEY = 'control_pivot_col_order_v1';

export function loadStoredPivotColumnOrder(): ControlPivotColumnKey[] {
  try {
    const raw =
      typeof localStorage !== 'undefined'
        ? localStorage.getItem(PIVOT_COLUMN_ORDER_STORAGE_KEY)
        : null;
    if (!raw) return [...DEFAULT_PIVOT_COLUMN_ORDER];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [...DEFAULT_PIVOT_COLUMN_ORDER];
    const validSet = new Set<ControlPivotColumnKey>(DEFAULT_PIVOT_COLUMN_ORDER);
    const cleaned: ControlPivotColumnKey[] = [];
    for (const item of parsed) {
      if (validSet.has(item) && !cleaned.includes(item)) {
        cleaned.push(item);
      }
    }
    for (const defKey of DEFAULT_PIVOT_COLUMN_ORDER) {
      if (!cleaned.includes(defKey)) {
        cleaned.push(defKey);
      }
    }
    return cleaned;
  } catch {
    return [...DEFAULT_PIVOT_COLUMN_ORDER];
  }
}

export function loadStoredDriveOverrides(): Record<string, string> {
  try {
    const raw = typeof localStorage !== 'undefined' ? localStorage.getItem(DRIVE_OVERRIDES_STORAGE_KEY) : null;
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

export function loadStoredCustomCatalog(): DriveCodeEntry[] {
  try {
    const raw = typeof localStorage !== 'undefined' ? localStorage.getItem(DRIVE_CUSTOM_CATALOG_STORAGE_KEY) : null;
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

/**
 * Builds the Excel-style Control Aggregation Pivot Table grouped by Ingredient (Code + Name)
 * Sorted descending by total (الأكبر فوق إلى الأصغر) by default.
 */
export function calculateControlPivot(
  sheet: SheetData,
  metric: 'actual' | 'target' = 'actual',
  divisor: number = 1,
  overrides?: Record<string, string>,
  customCatalog?: DriveCodeEntry[]
): ControlPivotSummary {
  if (!sheet || !sheet.headers || !sheet.rows || sheet.rows.length === 0) {
    return {
      rows: [],
      totals: { fattening: 0, fish: 0, duck: 0, pets: 0, grandTotal: 0 },
      unit: 'كجم',
      hasActualData: false,
      hasTargetData: false,
      actualColName: null,
      targetColName: null,
      metricUsed: metric,
      cleaningTotal: 0,
    };
  }

  const itemTypeColIdx = findItemTypeColumnIndex(sheet.headers);
  const formulaColIdx = findFormulaColumnIndex(sheet.headers);
  const nameColIdx = findIngredientNameColumnIndex(sheet.headers, formulaColIdx, itemTypeColIdx);
  const codeColIndices = findIngredientCodeColumns(sheet.headers, sheet.rows);

  const { targetColIndex, targetColName, actualColIndex, actualColName, unit } =
    detectWeightColumns(sheet.headers, sheet.rows);

  const chosenColIdx =
    metric === 'actual'
      ? actualColIndex !== -1
        ? actualColIndex
        : targetColIndex
      : targetColIndex !== -1
      ? targetColIndex
      : actualColIndex;

  const map = new Map<string, ControlPivotRow>();
  let cleaningTotal = 0;

  for (const row of sheet.rows) {
    if (!Array.isArray(row) || isSummaryOrTotalRow(row)) continue;

    // Extract raw ingredient name & possible embedded code
    let rawName = (row[nameColIdx] ?? '').toString().trim();
    if (!rawName) continue;

    // Extract code from dedicated code columns (pick first non-empty / non-dash value)
    let rawCode = '';
    for (const cIdx of codeColIndices) {
      const candidate = (row[cIdx] ?? '').toString().trim();
      if (candidate && candidate !== '-' && candidate !== '—') {
        rawCode = candidate;
        break;
      }
    }

    // Check if name still contains embedded code like "Corn (11342)"
    const split = extractCodeFromValue(rawName);
    if (split) {
      rawName = split.cleanText;
      if (!rawCode) {
        rawCode = split.code;
      }
    }

    // Group variants by name (Corn B -> Corn, Local soya 46% / SOYA 46% B -> soya 46%)
    const grouping = groupItemByName(rawName);
    rawName = grouping.name;

    // Determine ItemType
    let itemType: ItemType = 'اليف';
    if (itemTypeColIdx !== -1) {
      const tVal = (row[itemTypeColIdx] ?? '').toString().trim();
      if (tVal === 'سمك' || tVal === 'تسمين' || tVal === 'بط' || tVal === 'اليف' || tVal === 'تنظيف') {
        itemType = tVal;
      }
    }

    // Parse quantity
    const rawQty = chosenColIdx !== -1 ? parseNumericCell(row[chosenColIdx]) : 0;
    const qty = divisor > 0 ? rawQty / divisor : rawQty;

    // Cleaning groups are their own category: counted apart, never mixed into the item totals
    if (itemType === 'تنظيف') {
      cleaningTotal += qty;
      continue;
    }

    // Group key: include BOTH normalized code and normalized name so distinct items
    // sharing a JDE code (e.g. Corn vs Corn B [11342], Local soya 46% vs SOYA 46% B [913686]) stay separate.
    const normCode = rawCode.trim();
    const normName = rawName.toLowerCase().replace(/\s+/g, ' ').trim();
    const key = normCode ? `CODE:${normCode}::NAME:${normName}` : `NAME:${normName}`;

    let existing = map.get(key);
    if (!existing && grouping.grouped) {
      for (const entry of map.values()) {
        if (normalizeItemTextForGroup(entry.name) === normalizeItemTextForGroup(rawName)) {
          existing = entry;
          break;
        }
      }
    }
    // Merge ONLY when normalized ingredient name is identical (e.g. one row had code and another didn't)
    if (!existing && normCode) {
      const nameFallbackKey = `NAME:${normName}`;
      const existingByName = map.get(nameFallbackKey);
      if (existingByName) {
        map.delete(nameFallbackKey);
        existingByName.code = normCode;
        map.set(key, existingByName);
        existing = existingByName;
      }
    } else if (!existing && !normCode) {
      for (const entry of map.values()) {
        if (entry.name.toLowerCase().replace(/\s+/g, ' ').trim() === normName) {
          existing = entry;
          break;
        }
      }
    }

    if (!existing) {
      existing = {
        driveCode: '',
        code: normCode,
        name: rawName,
        fattening: 0,
        fish: 0,
        duck: 0,
        pets: 0,
        total: 0,
        occurrenceCount: 0,
      };
      map.set(key, existing);
    }

    if (itemType === 'تسمين') existing.fattening += qty;
    else if (itemType === 'سمك') existing.fish += qty;
    else if (itemType === 'بط') existing.duck += qty;
    else existing.pets += qty;

    existing.total += qty;
    existing.occurrenceCount += 1;
  }

  const effectiveOverrides = overrides ?? loadStoredDriveOverrides();
  const effectiveCatalog = customCatalog ?? loadStoredCustomCatalog();

  const rows = Array.from(map.values())
    .map((r) => ({
      ...r,
      driveCode: lookupDriveCode(r.name, r.code, effectiveOverrides, effectiveCatalog),
    }))
    .sort((a, b) => b.total - a.total);

  const totals = rows.reduce(
    (acc, r) => ({
      fattening: acc.fattening + r.fattening,
      fish: acc.fish + r.fish,
      duck: acc.duck + r.duck,
      pets: acc.pets + r.pets,
      grandTotal: acc.grandTotal + r.total,
    }),
    { fattening: 0, fish: 0, duck: 0, pets: 0, grandTotal: 0 }
  );

  return {
    rows,
    totals,
    unit,
    hasActualData: actualColIndex !== -1,
    hasTargetData: targetColIndex !== -1,
    actualColName,
    targetColName,
    metricUsed: metric,
    cleaningTotal,
  };
}

export function isWaterPivotRow(
  row: Pick<ControlPivotRow, 'name' | 'code' | 'driveCode'>
): boolean {
  const code = (row.code || '').trim();
  const driveCode = (row.driveCode || '').trim();
  if (code === '5050' || driveCode === '5050') return true;
  const norm = (row.name || '').trim().toLowerCase();
  return (
    /\bwater\b/i.test(norm) ||
    /(^|\s|-|_)(ال)?(ماء|مياه)($|\s|-|_)/.test(norm)
  );
}

export interface PivotTotalsGroup {
  fattening: number;
  fish: number;
  duck: number;
  pets: number;
  grandTotal: number;
}

export function sumPivotRows(rows: ControlPivotRow[]): PivotTotalsGroup {
  return rows.reduce(
    (acc, r) => ({
      fattening: acc.fattening + r.fattening,
      fish: acc.fish + r.fish,
      duck: acc.duck + r.duck,
      pets: acc.pets + r.pets,
      grandTotal: acc.grandTotal + r.total,
    }),
    { fattening: 0, fish: 0, duck: 0, pets: 0, grandTotal: 0 }
  );
}

/**
 * Builds a SheetData representation of the Control Pivot Table for Excel / CSV / Clipboard export.
 * Separates Water rows below "الإجمالي (بدون الماء)" followed by "الإجمالي الكلي (بالماء)" when Water exists.
 */
export function buildControlPivotSheetData(
  pivot: ControlPivotSummary,
  decimals: number = 6,
  columnOrder?: ControlPivotColumnKey[],
  customRows?: ControlPivotRow[]
): SheetData {
  const fmt = (val: number) => (Math.abs(val) < 1e-9 ? '-' : Number(val.toFixed(decimals)));
  const order = columnOrder && columnOrder.length > 0 ? columnOrder : loadStoredPivotColumnOrder();
  const sourceRows = customRows ?? pivot.rows;

  const headers = order.map((colKey) => PIVOT_COLUMN_DEFS[colKey].label);

  const mapDataRow = (r: ControlPivotRow): (string | number)[] =>
    order.map((colKey) => {
      switch (colKey) {
        case 'driveCode':
          return r.driveCode || '';
        case 'code':
          return r.code || '';
        case 'name':
          return r.name;
        case 'fattening':
          return fmt(r.fattening);
        case 'fish':
          return fmt(r.fish);
        case 'duck':
          return fmt(r.duck);
        case 'pets':
          return fmt(r.pets);
        case 'total':
          return Number(r.total.toFixed(decimals));
      }
    });

  const buildTotalRow = (label: string, totals: PivotTotalsGroup): (string | number)[] =>
    order.map((colKey) => {
      switch (colKey) {
        case 'driveCode':
        case 'code':
          return '';
        case 'name':
          return label;
        case 'fattening':
          return Number(totals.fattening.toFixed(decimals));
        case 'fish':
          return Number(totals.fish.toFixed(decimals));
        case 'duck':
          return Number(totals.duck.toFixed(decimals));
        case 'pets':
          return Number(totals.pets.toFixed(decimals));
        case 'total':
          return Number(totals.grandTotal.toFixed(decimals));
      }
    });

  const regularRows = sourceRows.filter((r) => !isWaterPivotRow(r));
  const waterRows = sourceRows.filter((r) => isWaterPivotRow(r));

  const totalsWithoutWater = sumPivotRows(regularRows);
  const totalsWithWater = sumPivotRows(sourceRows);

  const rows: (string | number)[][] = regularRows.map(mapDataRow);
  let totalsRow: (string | number)[];

  if (waterRows.length > 0) {
    rows.push(buildTotalRow('الإجمالي (بدون الماء)', totalsWithoutWater));
    for (const wRow of waterRows) {
      rows.push(mapDataRow(wRow));
    }
    totalsRow = buildTotalRow('الإجمالي الكلي (بالماء)', totalsWithWater);
  } else {
    totalsRow = buildTotalRow('الإجمالي الكلي', totalsWithWater);
  }

  return {
    sheetName: 'تجميع_الكنترول',
    description: 'صفحة تجميع الكنترول وإجمالي كل صنف موزعة حسب نوعه',
    isRtl: true,
    headers,
    rows,
    totalsRow,
  };
}
