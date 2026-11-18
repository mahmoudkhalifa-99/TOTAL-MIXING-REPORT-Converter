import { SheetData } from '../types/converter';
import {
  findItemTypeColumnIndex,
  findFormulaColumnIndex,
  classifyFormulaItemType,
  ItemType,
} from './itemClassifier';
import { detectWeightColumns, parseNumericCell } from './categoryTotals';
import { extractCodeFromValue } from './codeSplitter';
import {
  MASTER_DRIVE_CODES_CATALOG,
  lookupDriveCode,
  normalizeItemText,
  groupItemByName,
} from './driveCodesCatalog';
import {
  isWaterPivotRow,
  loadStoredDriveOverrides,
  loadStoredCustomCatalog,
} from './controlPivotCalculator';
import { extractSupledgerAndCleanFormula } from './formulaPropagator';

export interface SadatFormulaColumn {
  key: string;          // Unique formula name
  label: string;        // Display name
  itemType: ItemType;   // 'تسمين' | 'سمك' | 'بط' | 'اليف'
  productCode: string;  // Optional Product_Code / Batch number from PDF
  columnTotal: number;
}

export interface SadatMatrixRow {
  code: string;                 // كود الصنف (JDE)
  driveCode: string;            // كود الدريف
  name: string;                 // Raw ingredient name from PDF
  bilingualName: string;        // Enriched bilingual name (like "Limestone - حجر جيري")
  formulaValues: Record<string, number>; // formulaKey -> quantity
  fatteningTotal: number;       // اجمالي التسمين
  fishTotal: number;            // اجمالي سمك
  duckTotal: number;            // اجمالي البط
  petsTotal: number;            // اجمالي الاليف
  grandTotal: number;           // الإجمالي العام
  isWater: boolean;
}

export interface SadatControlReport {
  fatteningFormulas: SadatFormulaColumn[]; // تسمين
  fishFormulas: SadatFormulaColumn[];      // سمك
  duckFormulas: SadatFormulaColumn[];      // بط
  petsFormulas: SadatFormulaColumn[];      // اليف
  hasFattening: boolean;
  hasFish: boolean;
  hasDuck: boolean;
  hasPets: boolean;
  rows: SadatMatrixRow[];
  unit: string;
  hasActualData: boolean;
  hasTargetData: boolean;
  detectedDate: string;
}

export const SADAT_BATCH_CODES_STORAGE_KEY = 'sadat_control_batch_codes_v1';
export const SADAT_META_STORAGE_KEY = 'sadat_control_meta_v1';

export function isSummaryOrTotalRow(row: any[]): boolean {
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

export function findIngredientNameColumnIndex(
  headers: string[],
  formulaColIdx: number,
  itemTypeColIdx: number
): number {
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
    if (i === itemTypeColIdx || i === formulaColIdx) continue;
    const h = (headers[i] ?? '').toString().trim().toLowerCase();
    if (isCodeOrIdHeader(h)) continue;
    if (exactPriority.some((p) => h === p || h.includes(p))) {
      return i;
    }
  }

  for (let i = 0; i < headers.length; i++) {
    if (i === itemTypeColIdx || i === formulaColIdx) continue;
    const h = (headers[i] ?? '').toString().trim().toLowerCase();
    if (!isCodeOrIdHeader(h) && !h.includes('target') && !h.includes('actual') && !h.includes('المستهدف') && !h.includes('الفعلي')) {
      return i;
    }
  }
  return formulaColIdx !== -1 ? formulaColIdx : 0;
}

export function findIngredientCodeColumns(headers: string[], rows: any[][]): number[] {
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

function findSupledgerColumns(headers: string[]): number[] {
  const highPriority: number[] = [];
  const fallback: number[] = [];

  headers.forEach((h, idx) => {
    const norm = (h ?? '').toString().trim().toLowerCase();
    if (
      norm === 'supledger' ||
      norm === 'subledger' ||
      norm === 'سوب ليدجر' ||
      norm === 'رقم التشغيلة' ||
      norm === 'batch' ||
      norm === 'order_no' ||
      norm === 'order no'
    ) {
      highPriority.push(idx);
    } else if (
      norm === 'product_code' ||
      norm === 'product code' ||
      norm === 'كود التركيبة'
    ) {
      fallback.push(idx);
    }
  });

  return [...highPriority, ...fallback];
}

function isValidSupledgerValue(val: string): boolean {
  const cleaned = val.replace(/,/g, '').trim();
  // Supledger IDs in Control reports are numeric (4+ digits, e.g., 11336, 1504526, 908140)
  // Rejects weight/unit artifacts like "25KG", "Mono", "-"
  return /^[0-9]{4,}$/.test(cleaned);
}

function resolveBilingualName(rawName: string, jdeCode: string): string {
  const normRaw = normalizeItemText(rawName);
  if (!normRaw) return rawName;

  const hasArabic = /[\u0600-\u06FF]/.test(rawName);
  const hasLatin = /[a-zA-Z]/.test(rawName);
  if (hasArabic && hasLatin) return rawName;

  for (const entry of MASTER_DRIVE_CODES_CATALOG) {
    const entryNorm = normalizeItemText(entry.name);
    if (entryNorm === normRaw || entry.aliases?.some((a) => normalizeItemText(a) === normRaw)) {
      return entry.name;
    }
  }

  if (jdeCode) {
    const byJde = MASTER_DRIVE_CODES_CATALOG.filter((e) => e.jdeCode === jdeCode.trim());
    if (byJde.length === 1) {
      return byJde[0].name;
    }
  }

  return rawName;
}

export function extractDateFromText(text?: string): string {
  if (text) {
    const match = text.match(/(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/);
    if (match) return match[1].replace(/-/g, '/');
  }
  const now = new Date();
  const d = String(now.getDate()).padStart(2, '0');
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const y = now.getFullYear();
  return `${d}/${m}/${y}`;
}

export function calculateSadatControlReport(
  sheet: SheetData,
  metric: 'actual' | 'target' = 'actual',
  divisor: number = 1,
  documentTitle?: string
): SadatControlReport {
  const detectedDate = extractDateFromText(documentTitle || sheet?.sheetName);

  if (!sheet || !sheet.headers || !sheet.rows || sheet.rows.length === 0) {
    return {
      fatteningFormulas: [],
      fishFormulas: [],
      duckFormulas: [],
      petsFormulas: [],
      hasFattening: false,
      hasFish: false,
      hasDuck: false,
      hasPets: false,
      rows: [],
      unit: 'كجم',
      hasActualData: false,
      hasTargetData: false,
      detectedDate,
    };
  }

  const itemTypeColIdx = findItemTypeColumnIndex(sheet.headers);
  const formulaColIdx = findFormulaColumnIndex(sheet.headers);
  const nameColIdx = findIngredientNameColumnIndex(sheet.headers, formulaColIdx, itemTypeColIdx);
  const codeColIndices = findIngredientCodeColumns(sheet.headers, sheet.rows);
  const supledgerColIndices = findSupledgerColumns(sheet.headers);

  const { targetColIndex, actualColIndex, unit } = detectWeightColumns(
    sheet.headers,
    sheet.rows
  );

  const chosenColIdx =
    metric === 'actual'
      ? actualColIndex !== -1
        ? actualColIndex
        : targetColIndex
      : targetColIndex !== -1
      ? targetColIndex
      : actualColIndex;

  const formulaMetaMap = new Map<string, SadatFormulaColumn>();
  const rowMap = new Map<string, SadatMatrixRow>();

  const overrides = loadStoredDriveOverrides();
  const customCatalog = loadStoredCustomCatalog();

  for (const row of sheet.rows) {
    if (!Array.isArray(row) || isSummaryOrTotalRow(row)) continue;

    let rawName = (row[nameColIdx] ?? '').toString().trim();
    if (!rawName) continue;

    let rawCode = '';
    for (const cIdx of codeColIndices) {
      const candidate = (row[cIdx] ?? '').toString().trim();
      if (candidate && candidate !== '-' && candidate !== '—') {
        rawCode = candidate;
        break;
      }
    }

    const split = extractCodeFromValue(rawName);
    if (split) {
      rawName = split.cleanText;
      if (!rawCode) rawCode = split.code;
    }

    // Group variants by name (Corn B -> Corn, Local soya 46% / SOYA 46% B -> soya 46%)
    const grouping = groupItemByName(rawName);
    rawName = grouping.name;

    const rawFormulaCell =
      formulaColIdx !== -1
        ? (row[formulaColIdx] ?? '').toString().trim() || 'تركيبة غير محددة'
        : sheet.sheetName || 'التركيبة الرئيسية';

    const parsedFormula = extractSupledgerAndCleanFormula(rawFormulaCell);
    const rawFormula = parsedFormula.cleanName || rawFormulaCell;

    let itemType: ItemType = classifyFormulaItemType(rawFormula);
    if (itemTypeColIdx !== -1) {
      const tVal = (row[itemTypeColIdx] ?? '').toString().trim();
      if (tVal === 'سمك' || tVal.includes('سمك')) {
        itemType = 'سمك';
      } else if (tVal === 'تسمين' || tVal.includes('تسمين')) {
        itemType = 'تسمين';
      } else if (tVal === 'بط' || tVal.includes('بط')) {
        itemType = 'بط';
      } else if (tVal === 'اليف' || tVal === 'أليف' || tVal.includes('أليف') || tVal.includes('اليف')) {
        itemType = 'اليف';
      }
    }

    let productCode = parsedFormula.supledger || '';
    if (!productCode) {
      for (const sIdx of supledgerColIndices) {
        const pVal = (row[sIdx] ?? '').toString().trim();
        if (isValidSupledgerValue(pVal)) {
          productCode = pVal.replace(/,/g, '');
          break;
        }
      }
    }

    const rawQty = chosenColIdx !== -1 ? parseNumericCell(row[chosenColIdx]) : 0;
    const qty = divisor > 0 ? rawQty / divisor : rawQty;

    let fCol = formulaMetaMap.get(rawFormula);
    if (!fCol) {
      fCol = {
        key: rawFormula,
        label: rawFormula,
        itemType,
        productCode,
        columnTotal: 0,
      };
      formulaMetaMap.set(rawFormula, fCol);
    } else if (!fCol.productCode && productCode) {
      fCol.productCode = productCode;
    }
    fCol.columnTotal += qty;

    const normCode = rawCode.trim();
    const normName = rawName.toLowerCase().replace(/\s+/g, ' ').trim();
    const rowKey = normCode ? `CODE:${normCode}::NAME:${normName}` : `NAME:${normName}`;

    let existing = rowMap.get(rowKey);
    if (!existing && grouping.grouped) {
      // Name-grouped family: match by name only, ignoring differing item codes
      for (const entry of rowMap.values()) {
        if (normalizeItemText(entry.name) === normalizeItemText(rawName)) {
          existing = entry;
          break;
        }
      }
    }
    if (!existing && normCode) {
      const fallbackKey = `NAME:${normName}`;
      const byName = rowMap.get(fallbackKey);
      if (byName) {
        rowMap.delete(fallbackKey);
        byName.code = normCode;
        rowMap.set(rowKey, byName);
        existing = byName;
      }
    } else if (!existing && !normCode) {
      for (const entry of rowMap.values()) {
        if (entry.name.toLowerCase().replace(/\s+/g, ' ').trim() === normName) {
          existing = entry;
          break;
        }
      }
    }

    if (!existing) {
      const driveCode = lookupDriveCode(rawName, normCode, overrides, customCatalog);
      existing = {
        code: normCode,
        driveCode,
        name: rawName,
        bilingualName: resolveBilingualName(rawName, normCode),
        formulaValues: {},
        fatteningTotal: 0,
        fishTotal: 0,
        duckTotal: 0,
        petsTotal: 0,
        grandTotal: 0,
        isWater: isWaterPivotRow({ name: rawName, code: normCode, driveCode }),
      };
      rowMap.set(rowKey, existing);
    }

    existing.formulaValues[rawFormula] = (existing.formulaValues[rawFormula] || 0) + qty;
    if (itemType === 'تسمين') {
      existing.fatteningTotal += qty;
    } else if (itemType === 'سمك') {
      existing.fishTotal += qty;
    } else if (itemType === 'بط') {
      existing.duckTotal += qty;
    } else {
      existing.petsTotal += qty;
    }
    existing.grandTotal += qty;
  }

  const allFormulas = Array.from(formulaMetaMap.values());
  const byLabel = (a: SadatFormulaColumn, b: SadatFormulaColumn) =>
    a.label.localeCompare(b.label, 'en');

  const fatteningFormulas = allFormulas.filter((f) => f.itemType === 'تسمين').sort(byLabel);
  const fishFormulas = allFormulas.filter((f) => f.itemType === 'سمك').sort(byLabel);
  const duckFormulas = allFormulas.filter((f) => f.itemType === 'بط').sort(byLabel);
  const petsFormulas = allFormulas.filter((f) => f.itemType === 'اليف').sort(byLabel);

  const rows = Array.from(rowMap.values()).sort((a, b) => b.grandTotal - a.grandTotal);

  return {
    fatteningFormulas,
    fishFormulas,
    duckFormulas,
    petsFormulas,
    hasFattening: fatteningFormulas.length > 0,
    hasFish: fishFormulas.length > 0,
    hasDuck: duckFormulas.length > 0,
    hasPets: petsFormulas.length > 0,
    rows,
    unit,
    hasActualData: actualColIndex !== -1,
    hasTargetData: targetColIndex !== -1,
    detectedDate,
  };
}
