import { SheetData, ExtractedDocument } from '../types/converter';

export type ItemType = 'تسمين' | 'سمك' | 'بط' | 'اليف';

/**
 * Classifies an item based on formula name (اسم التركيبة) according to user rules:
 * 1. Contains "Tilapia" (case-insensitive) -> "سمك"
 * 2. Contains "GR", "FN", "ST", "SS" -> "تسمين"
 * 3. Contains "Duck" (case-insensitive) -> "بط"
 * 4. Otherwise -> "اليف"
 */
export function classifyFormulaItemType(formulaText: any): ItemType {
  if (formulaText === null || formulaText === undefined) return 'اليف';
  const text = formulaText.toString().trim();
  if (!text) return 'اليف';

  // 1. Tilapia rule -> سمك
  if (/tilapia|بلطي|أسماك|سمك/i.test(text)) {
    return 'سمك';
  }

  // 2. Duck rule -> بط (check before broiler codes so "Duck ST" or "Duck Grower" is classified as بط)
  if (/duck|بط/i.test(text)) {
    return 'بط';
  }

  // 3. Broiler / Fattening codes (GR, FN, ST, SS) -> تسمين
  // Matches standalone codes or codes prefixed/suffixed by space, dash, slash, underscore, or digits (e.g., "FN 13-4 A", "GR-21", "ST-19", "SS 23%", "FN13", "ST2", "GR3", "SS1")
  // Prevents false matches in ordinary words like "First", "Cost", "Stock", "Assistant", "Plastic"
  const broilerCodeRegex = /(?:^|[\s\-_/.,()#])(GR|FN|ST|SS)(?:[\s\-_/.,()#]|$|[0-9])/i;
  const broilerWordRegex = /(?:تسمين|بادي|بادئ|نامي|ناهي|برويد|broiler)/i;

  if (broilerCodeRegex.test(text) || broilerWordRegex.test(text)) {
    return 'تسمين';
  }

  // 4. Default -> اليف
  return 'اليف';
}

/**
 * Finds the column index for "اسم التركيبة" or similar formula / product column.
 */
export function findFormulaColumnIndex(headers: string[]): number {
  if (!headers || headers.length === 0) return -1;

  // Exact or high priority matches
  const highPriorityNames = [
    'اسم التركيبة',
    'التركيبة',
    'اسم الخلطة',
    'الخلطة',
    'finished_product',
    'finished product',
    'finished',
    'formula name',
    'formula_name',
    'formula',
    'recipe',
    'product name',
    'product_name',
    'اسم العلف',
    'نوع العلف',
  ];

  for (let i = 0; i < headers.length; i++) {
    const h = (headers[i] ?? '').toString().trim().toLowerCase();
    if (highPriorityNames.some((p) => h === p || h.includes(p))) {
      return i;
    }
  }

  // Medium priority matches (product, description, item name)
  const medPriorityNames = [
    'اسم الصنف',
    'إسم الصنف',
    'اسم المنتج',
    'المنتج',
    'item name',
    'description',
    'الوصف',
    'البيان',
    'المادة',
    'الصنف',
    'product',
  ];

  for (let i = 0; i < headers.length; i++) {
    const h = (headers[i] ?? '').toString().trim().toLowerCase();
    if (
      h.includes('code') ||
      h.includes('كود') ||
      h.includes('رمز') ||
      h.includes('رقم') ||
      h.includes('supledger') ||
      h.includes('subledger') ||
      h.includes('نوع الصنف')
    ) {
      continue;
    }
    if (medPriorityNames.some((p) => h === p || h.includes(p))) {
      return i;
    }
  }

  return -1;
}

/**
 * Finds the column index for "نوع الصنف" if already present
 */
export function findItemTypeColumnIndex(headers: string[]): number {
  if (!headers) return -1;
  return headers.findIndex((h) => {
    const norm = (h ?? '').toString().trim().toLowerCase();
    return norm === 'نوع الصنف' || norm === 'نوع_الصنف' || norm === 'item type' || norm === 'item_type' || norm === 'category';
  });
}

/**
 * Ensures a sheet has the "نوع الصنف" column populated according to formula rules.
 */
export function ensureItemTypeColumn(sheet: SheetData, documentContext?: Partial<ExtractedDocument>): {
  sheet: SheetData;
  addedOrUpdated: boolean;
  itemTypeColIndex: number;
} {
  if (!sheet.headers || sheet.headers.length === 0) {
    return { sheet, addedOrUpdated: false, itemTypeColIndex: -1 };
  }

  const existingColIndex = findItemTypeColumnIndex(sheet.headers);
  const formulaColIndex = findFormulaColumnIndex(sheet.headers);

  // Context fallback (e.g. if the whole sheet is a mixing report for one formula)
  const contextText = `${sheet.sheetName || ''} ${sheet.description || ''} ${documentContext?.documentTitle || ''} ${documentContext?.documentSummary || ''}`;

  let targetColIndex = existingColIndex;
  const newHeaders = [...sheet.headers];

  if (targetColIndex === -1) {
    // Insert "نوع الصنف" right after formula column if found, or right after the first column
    if (formulaColIndex !== -1) {
      targetColIndex = formulaColIndex + 1;
    } else {
      targetColIndex = Math.min(1, newHeaders.length);
    }
    newHeaders.splice(targetColIndex, 0, 'نوع الصنف');
  }

  const newRows = sheet.rows.map((row) => {
    const newRow = [...row];
    let formulaText = '';

    if (formulaColIndex !== -1 && row[formulaColIndex] !== undefined) {
      formulaText = (row[formulaColIndex] ?? '').toString();
    }

    // If formula text in that cell is blank, check other text cells in the row
    if (!formulaText.trim()) {
      for (let c = 0; c < row.length; c++) {
        if (c === existingColIndex) continue;
        const val = (row[c] ?? '').toString();
        // check if this cell has keywords
        if (/tilapia|بلطي|duck|بط|broiler|تسمين/i.test(val) || /(?:^|[\s\-_/])(GR|FN|ST|SS)(?:[\s\-_/]|$|[0-9])/i.test(val)) {
          formulaText = val;
          break;
        }
      }
    }

    // If still blank, use sheet / document context
    if (!formulaText.trim()) {
      formulaText = contextText;
    }

    const calculatedType = classifyFormulaItemType(formulaText);

    if (existingColIndex !== -1) {
      newRow[existingColIndex] = calculatedType;
    } else {
      newRow.splice(targetColIndex, 0, calculatedType);
    }

    return newRow;
  });

  // Handle totalsRow if present
  let newTotalsRow = sheet.totalsRow ? [...sheet.totalsRow] : undefined;
  if (newTotalsRow && newTotalsRow.length > 0) {
    if (existingColIndex === -1) {
      newTotalsRow.splice(targetColIndex, 0, '');
    } else {
      newTotalsRow[existingColIndex] = '';
    }
  }

  return {
    sheet: {
      ...sheet,
      headers: newHeaders,
      rows: newRows,
      totalsRow: newTotalsRow,
    },
    addedOrUpdated: true,
    itemTypeColIndex: targetColIndex,
  };
}
