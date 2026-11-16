import { SheetData } from '../types/converter';

/**
 * Splits text into clean name and code
 * Examples:
 * - "Anti ox ( 1472106 )" -> { cleanText: "Anti ox", code: "1472106" }
 * - "Bile Acid 200g (907948)" -> { cleanText: "Bile Acid 200g", code: "907948" }
 * - "Corn (11342)" -> { cleanText: "Corn", code: "11342" }
 * - "LIQUID METHONIN 1287795" -> { cleanText: "LIQUID METHONIN", code: "1287795" }
 * - "FN 13-4 A CODE 111" -> { cleanText: "FN 13-4 A", code: "111" }
 */
function isLikelyEnzymeOrUnitNumber(numStr: string): boolean {
  // Round multiples of 1000 like 1000, 2000, 10000, 20000, 50000, 100000
  // are enzyme activity units / potencies (e.g. "Ronzyme WX 2000", "Hi-phos GT 20000"), not JDE codes.
  return /^[1-9]0{3,}$/.test(numStr.trim());
}

function isIngredientCodeHeader(header: string): boolean {
  const h = (header ?? '').toString().trim().toLowerCase();
  if (!h) return false;
  if (
    h.includes('product') ||
    h.includes('تركيبة') ||
    h.includes('supledger') ||
    h.includes('subledger') ||
    h.includes('drive') ||
    h.includes('دريف')
  ) {
    return false;
  }
  return (
    h === 'ingredient_code' ||
    h === 'ingredient code' ||
    h === 'كود الصنف' ||
    h === 'كود_الصنف' ||
    h === 'item_code' ||
    h === 'item code' ||
    h === 'code' ||
    h === 'الكود' ||
    h === 'رمز الصنف' ||
    h === 'كود المكون'
  );
}

function isNonSplittableHeader(header: string): boolean {
  const h = (header ?? '').toString().trim().toLowerCase();
  if (!h) return false;
  return (
    h.includes('finished') ||
    h.includes('تركيبة') ||
    h.includes('خلطة') ||
    h.includes('formula') ||
    h.includes('supledger') ||
    h.includes('subledger') ||
    h.includes('نوع الصنف') ||
    h.includes('نوع_الصنف') ||
    h.includes('code') ||
    h.includes('كود') ||
    h.includes('target') ||
    h.includes('actual') ||
    h.includes('المستهدف') ||
    h.includes('الفعلي')
  );
}

function isValidCodeCell(val: any): boolean {
  const s = (val ?? '').toString().trim();
  return Boolean(s && s !== '-' && s !== '—' && s !== 'null' && s !== 'undefined');
}

export function extractCodeFromValue(
  val: any,
  options?: { strictOnly?: boolean }
): { cleanText: string; code: string } | null {
  if (val === null || val === undefined) return null;
  const str = val.toString().trim();
  if (!str) return null;

  // 1. Parentheses pattern for numeric codes only: "Anti ox ( 1472106 )" or "Corn (11342)"
  // Avoid stripping unit/weight/enzyme labels like "(25KG)", "(30%)", "(200g)", "( 20000 )"
  const parenMatch = str.match(/^(.*?)\s*\(\s*([0-9]{4,}[A-Za-z0-9_-]*)\s*\)\s*$/);
  if (parenMatch && parenMatch[1].trim() && parenMatch[2].trim()) {
    const candidateCode = parenMatch[2].trim();
    if (!isLikelyEnzymeOrUnitNumber(candidateCode)) {
      return {
        cleanText: parenMatch[1].trim(),
        code: candidateCode,
      };
    }
  }

  // 2. Name in parentheses followed by code: "(Bile Acid 200g) 907948" or "(corn 7.5%) 11342" or "(Biotrinic TOP 3)13855"
  const parenNameWithCodeMatch = str.match(/^\((.*?)\)\s*([0-9]{4,}[A-Za-z0-9_-]*)$/);
  if (parenNameWithCodeMatch && parenNameWithCodeMatch[1].trim() && parenNameWithCodeMatch[2].trim()) {
    const candidateCode = parenNameWithCodeMatch[2].trim();
    if (!isLikelyEnzymeOrUnitNumber(candidateCode)) {
      return {
        cleanText: parenNameWithCodeMatch[1].trim(),
        code: candidateCode,
      };
    }
  }

  // 3. Explicit CODE pattern: "FN 13-4 A CODE 111" or "Item CODE: 9821"
  const codeWordMatch = str.match(/^(.*?)\s+(?:CODE|كود)[:\s]+([0-9A-Za-z_-]+)\s*$/i);
  if (codeWordMatch && codeWordMatch[1].trim() && codeWordMatch[2].trim()) {
    return {
      cleanText: codeWordMatch[1].trim(),
      code: codeWordMatch[2].trim(),
    };
  }

  // 4. Trailing numeric code (at least 4 digits, preceded by space): "LIQUID METHONIN 1287795" or "Lysine HCL 13556"
  // Skip when strictOnly is true or when trailing number is a round enzyme activity figure (e.g. 2000, 10000, 20000)
  if (!options?.strictOnly) {
    const trailingNumMatch = str.match(/^(.*?)\s+([0-9]{4,})\s*$/);
    if (trailingNumMatch && trailingNumMatch[1].trim() && trailingNumMatch[2].trim()) {
      const candidateCode = trailingNumMatch[2].trim();
      if (!isLikelyEnzymeOrUnitNumber(candidateCode)) {
        return {
          cleanText: trailingNumMatch[1].trim(),
          code: candidateCode,
        };
      }
    }
  }

  return null;
}

/**
 * Merges duplicate ingredient code columns (for example, when an accidental split column
 * with 1-2 enzyme numbers appears before the real Ingredient_Code column) so the real
 * dense codes column is preserved and any falsely split numbers are restored to Ingredient_Name.
 */
export function consolidateDuplicateCodeColumns(sheet: SheetData): SheetData {
  if (!sheet?.headers || !sheet?.rows || sheet.rows.length === 0) {
    return sheet;
  }

  const codeColIndices: number[] = [];
  sheet.headers.forEach((h, idx) => {
    if (isIngredientCodeHeader(h)) {
      codeColIndices.push(idx);
    }
  });

  if (codeColIndices.length <= 1) {
    return sheet;
  }

  // Pick the column with the most valid codes (on tie, prefer the rightmost/original one)
  let primaryCol = codeColIndices[codeColIndices.length - 1];
  let bestCount = -1;
  for (const colIdx of codeColIndices) {
    let count = 0;
    for (const row of sheet.rows) {
      if (isValidCodeCell(row[colIdx])) {
        count++;
      }
    }
    if (count >= bestCount) {
      bestCount = count;
      primaryCol = colIdx;
    }
  }

  const dupCols = codeColIndices.filter((c) => c !== primaryCol);
  const nameCol = sheet.headers.findIndex((h) => {
    const norm = (h ?? '').toString().trim().toLowerCase();
    return (
      norm === 'ingredient_name' ||
      norm === 'ingredient name' ||
      norm === 'اسم الصنف' ||
      norm === 'إسم الصنف' ||
      norm === 'اسم المكون' ||
      norm === 'item_name'
    );
  });

  const newRows = sheet.rows.map((row) => {
    const nextRow = [...row];
    for (const dupCol of dupCols) {
      const dupVal = (nextRow[dupCol] ?? '').toString().trim();
      if (!isValidCodeCell(dupVal)) continue;

      const primaryVal = (nextRow[primaryCol] ?? '').toString().trim();
      if (!isValidCodeCell(primaryVal)) {
        if (!isLikelyEnzymeOrUnitNumber(dupVal)) {
          nextRow[primaryCol] = dupVal;
        }
      } else if (dupVal !== primaryVal && nameCol !== -1) {
        const nameStr = (nextRow[nameCol] ?? '').toString().trim();
        if (nameStr && !nameStr.includes(dupVal) && isLikelyEnzymeOrUnitNumber(dupVal)) {
          nextRow[nameCol] = `${nameStr} ${dupVal}`;
        }
      }
    }
    return nextRow;
  });

  const removeSet = new Set(dupCols);
  const newHeaders = sheet.headers.filter((_, idx) => !removeSet.has(idx));
  const cleanedRows = newRows.map((row) => row.filter((_, idx) => !removeSet.has(idx)));
  const newTotalsRow =
    sheet.totalsRow && sheet.totalsRow.length > 0
      ? sheet.totalsRow.filter((_, idx) => !removeSet.has(idx))
      : sheet.totalsRow;

  return {
    ...sheet,
    headers: newHeaders,
    rows: cleanedRows,
    totalsRow: newTotalsRow,
  };
}

/**
 * Automatically detects if a column in the sheet contains codes, and splits it into a new column.
 * If an Ingredient_Code column already exists with populated codes, avoids creating a duplicate column
 * and instead backfills missing cells in the existing code column.
 */
export function autoSplitCodeColumn(sheet: SheetData, specifiedColIndex?: number): {
  newSheet: SheetData;
  splitColumnIndex: number;
  newColumnName: string;
  countSplits: number;
} | null {
  if (!sheet.rows || sheet.rows.length === 0 || !sheet.headers) {
    return null;
  }

  const workingSheet = consolidateDuplicateCodeColumns(sheet);

  // If automatic mode and an Ingredient_Code column already exists with valid codes,
  // do NOT create a duplicate column — only backfill missing codes in that existing column.
  if (specifiedColIndex === undefined) {
    const existingCodeCol = workingSheet.headers.findIndex((h) => isIngredientCodeHeader(h));
    if (existingCodeCol !== -1) {
      const existingPopulated = workingSheet.rows.filter((r) => isValidCodeCell(r[existingCodeCol])).length;
      if (existingPopulated >= 2) {
        const nameCol = workingSheet.headers.findIndex((h) => {
          const norm = (h ?? '').toString().trim().toLowerCase();
          return (
            !isNonSplittableHeader(h) &&
            (norm.includes('ingredient') ||
              norm.includes('اسم الصنف') ||
              norm.includes('إسم الصنف') ||
              norm.includes('اسم المكون') ||
              norm.includes('item'))
          );
        });

        let backfilled = 0;
        if (nameCol !== -1) {
          const updatedRows = workingSheet.rows.map((row) => {
            const nextRow = [...row];
            if (!isValidCodeCell(nextRow[existingCodeCol])) {
              const split = extractCodeFromValue(nextRow[nameCol], { strictOnly: true });
              if (split) {
                nextRow[nameCol] = split.cleanText;
                nextRow[existingCodeCol] = split.code;
                backfilled++;
              }
            }
            return nextRow;
          });

          if (backfilled > 0 || workingSheet !== sheet) {
            return {
              newSheet: { ...workingSheet, rows: updatedRows },
              splitColumnIndex: nameCol,
              newColumnName: workingSheet.headers[existingCodeCol],
              countSplits: backfilled,
            };
          }
        }

        if (workingSheet !== sheet) {
          return {
            newSheet: workingSheet,
            splitColumnIndex: existingCodeCol,
            newColumnName: workingSheet.headers[existingCodeCol],
            countSplits: 0,
          };
        }
        return null;
      }
    }
  }

  let targetCol = -1;
  let maxMatches = 0;

  if (specifiedColIndex !== undefined && specifiedColIndex >= 0 && specifiedColIndex < workingSheet.headers.length) {
    targetCol = specifiedColIndex;
  } else {
    // Scan only non-formula, non-code columns to find the one with the most code patterns
    workingSheet.headers.forEach((header, colIdx) => {
      if (isNonSplittableHeader(header)) return;
      let matches = 0;
      workingSheet.rows.forEach((row) => {
        const cell = row[colIdx];
        if (extractCodeFromValue(cell)) {
          matches++;
        }
      });
      if (matches > maxMatches && matches >= 2) {
        maxMatches = matches;
        targetCol = colIdx;
      }
    });
  }

  if (targetCol === -1) {
    if (workingSheet !== sheet) {
      return {
        newSheet: workingSheet,
        splitColumnIndex: 0,
        newColumnName: 'Ingredient_Code',
        countSplits: 0,
      };
    }
    return null;
  }

  const originalHeader = (workingSheet.headers[targetCol] ?? '').toString();
  let newColName = 'Code';
  if (originalHeader.toLowerCase().includes('ingredient')) {
    newColName = 'Ingredient_Code';
  } else if (originalHeader.toLowerCase().includes('product')) {
    newColName = 'Product_Code';
  } else if (originalHeader.includes('صنف') || originalHeader.includes('مكون')) {
    newColName = 'كود_الصنف';
  } else {
    newColName = `${originalHeader || 'Item'}_Code`;
  }

  // Create new headers inserting newColName right after targetCol
  const newHeaders = [...workingSheet.headers];
  newHeaders.splice(targetCol + 1, 0, newColName);

  let countSplits = 0;
  // Create new rows
  const newRows = workingSheet.rows.map((row) => {
    const newRow = [...row];
    const cellVal = row[targetCol];
    const split = extractCodeFromValue(cellVal);

    if (split) {
      countSplits++;
      newRow[targetCol] = split.cleanText;
      newRow.splice(targetCol + 1, 0, split.code);
    } else {
      newRow.splice(targetCol + 1, 0, '');
    }
    return newRow;
  });

  // Handle totalsRow if present
  let newTotalsRow: (string | number)[] | undefined = undefined;
  if (workingSheet.totalsRow && workingSheet.totalsRow.length > 0) {
    newTotalsRow = [...workingSheet.totalsRow];
    newTotalsRow.splice(targetCol + 1, 0, '');
  }

  const newSheet: SheetData = {
    ...workingSheet,
    headers: newHeaders,
    rows: newRows,
    totalsRow: newTotalsRow,
  };

  return {
    newSheet,
    splitColumnIndex: targetCol,
    newColumnName: newColName,
    countSplits,
  };
}
