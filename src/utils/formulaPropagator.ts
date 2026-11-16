import { SheetData, ExtractedDocument } from '../types/converter';
import { findFormulaColumnIndex } from './itemClassifier';

/**
 * Extracts both a clean formula name and any trailing numeric supledger / batch ID:
 * E.g. "FN 19-9 A/2 Code111 (FN 19-9 A/2 Code111) 11336" -> { cleanName: "FN 19-9 A/2 Code111", supledger: "11336" }
 * E.g. "GR 28-6 A/8 (25KG) (GR 28-6 A/8 (25KG)) 1504526" -> { cleanName: "GR 28-6 A/8 (25KG)", supledger: "1504526" }
 * E.g. "(FN 19-9 A/7) 11308" -> { cleanName: "FN 19-9 A/7", supledger: "11308" }
 */
export function extractSupledgerAndCleanFormula(val: any): {
  cleanName: string;
  supledger: string;
} {
  if (val === null || val === undefined) return { cleanName: '', supledger: '' };
  const str = val.toString().trim();
  if (!str) return { cleanName: '', supledger: '' };

  let supledger = '';
  let work = str;

  // Match trailing supledger digits (4+ digits, optionally comma-separated) at the end
  const trailingNum = work.match(/^(.*?)(?:\s+|\)\s*|-\s*)([0-9][0-9,]{3,})\s*$/);
  if (trailingNum && trailingNum[1].trim()) {
    const candidateDigits = trailingNum[2].replace(/,/g, '');
    if (/^[0-9]{4,}$/.test(candidateDigits)) {
      supledger = candidateDigits;
      // Strip trailing digits AND any trailing separator punctuation like '-' or ':' (e.g. "A/3- 908140" -> "A/3")
      work = work
        .slice(0, work.length - trailingNum[2].length)
        .replace(/[\s\-:]+$/, '')
        .trim();
    }
  }

  // Collapse duplicated pattern: "X (X)" -> "X" (supports nested parens like "GR 28-6 A/8 (25KG) (GR 28-6 A/8 (25KG))")
  const halfDup = work.match(/^(.+?)\s*\(\s*\1\s*\)$/i);
  if (halfDup && halfDup[1].trim()) {
    work = halfDup[1].trim();
  } else {
    // Collapse unparenthesized Crystal Reports duplicated formula banner prefix:
    // e.g. "Tilapia 13-10 S (30%) Tilapia 13-10 S (30%) A/3" -> "Tilapia 13-10 S (30%) A/3"
    const prefixDup = work.match(/^(.{4,}?)\s+\1($|\s+.*)$/i);
    if (prefixDup && prefixDup[1].trim()) {
      work = `${prefixDup[1].trim()}${prefixDup[2] || ''}`.trim();
    } else if (work.startsWith('(') && work.endsWith(')')) {
      const inner = work.slice(1, -1).trim();
      if (inner) work = inner;
    }
  }

  return { cleanName: work.trim(), supledger };
}

/**
 * Cleans a formula header string if it contains duplicated parenthesis or batch codes:
 * E.g. "FN 19-9 A/1 Code111 (FN 19-9 A/1 Code111) 11336" -> "FN 19-9 A/1 Code111"
 */
export function cleanFormulaTitle(val: any): string {
  return extractSupledgerAndCleanFormula(val).cleanName;
}

/**
 * Checks if a row looks strictly like an empty section separator header row.
 * E.g. ["FN 19-9 A/1 Code111 (FN 19-9 A/1 Code111) 11336", "", "", "", ""]
 * where all other cells are completely blank or non-existent.
 */
export function isFormulaHeaderRow(row: any[]): {
  isHeader: boolean;
  formulaName: string;
  supledger: string;
} {
  if (!row || row.length === 0) return { isHeader: false, formulaName: '', supledger: '' };

  const firstCell = (row[0] ?? '').toString().trim();
  if (!firstCell) return { isHeader: false, formulaName: '', supledger: '' };

  // Don't treat totals rows as formula headers
  if (firstCell.includes('إجمالي') || firstCell.includes('المجموع') || /total|grand/i.test(firstCell)) {
    return { isHeader: false, formulaName: '', supledger: '' };
  }

  // Never treat raw ingredient names (e.g. "Duck Premix", "Rejected Pets") as formula section headers
  const looksLikeRawIngredient =
    /\b(?:premix|permix|permex|rejected|digest|powder|meal|bran|limestone|monocalcium|mono calcium|choline|methionin|lysine|threonine|salt|water|oil|soya|corn|ddg|wheat|screenings|mycofix|anti\s*ox)\b/i.test(
      firstCell
    ) && !/Code\s*[0-9]+/i.test(firstCell);
  if (looksLikeRawIngredient) {
    return { isHeader: false, formulaName: '', supledger: '' };
  }

  // Wrapped item-name fragment such as "Broiler 1.5 kg" is an ingredient name, not a formula header
  if (
    /(?:^|\s)\d+(?:[.,]\d+)?\s*(?:kg|kgs|كجم|كيلو|gm|g|gram|ton|طن|%)\s*$/i.test(firstCell) &&
    !/Code\s*[0-9]+/i.test(firstCell) &&
    !/\d+\s*-\s*\d+/.test(firstCell) &&
    !/(?:^|[\s\-_/(])(FN|GR|ST|SS)\s*[0-9]/i.test(firstCell)
  ) {
    return { isHeader: false, formulaName: '', supledger: '' };
  }

  // Count non-empty other cells
  let otherNonEmptyCount = 0;
  let hasValidQuantity = false;
  let extraSupledger = '';

  for (let i = 1; i < row.length; i++) {
    const c = (row[i] ?? '').toString().trim();
    if (c) {
      // Check if cell 1 holds "(Formula) 11336" (from PDF Ident + Name columns) with no numeric quantities elsewhere
      const parsedExtra = extractSupledgerAndCleanFormula(c);
      if (i === 1 && parsedExtra.supledger) {
        extraSupledger = parsedExtra.supledger;
        continue;
      }
      otherNonEmptyCount++;
      const num = parseFloat(c.replace(/,/g, ''));
      if (!isNaN(num) && num > 0) {
        hasValidQuantity = true;
      }
    }
  }

  // Known formula indicators (like "Code111", "FN", "Tilapia", "Duck", "GR", "ST", "SS", "Adult Dog")
  const hasFormulaKeywords =
    /Code\s*[0-9]+/i.test(firstCell) ||
    /(?:^|[\s\-_/])(FN|GR|ST|SS)[0-9\s\-_/]/i.test(firstCell) ||
    /\d+\s*-\s*\d+/.test(firstCell);

  if (hasFormulaKeywords && !hasValidQuantity && otherNonEmptyCount === 0) {
    const parsed = extractSupledgerAndCleanFormula(firstCell);
    return {
      isHeader: true,
      formulaName: parsed.cleanName,
      supledger: parsed.supledger || extraSupledger,
    };
  }

  return { isHeader: false, formulaName: '', supledger: '' };
}

function findExistingSupledgerCol(headers: string[]): number {
  return headers.findIndex((h) => {
    const norm = (h ?? '').toString().trim().toLowerCase();
    return (
      norm === 'supledger' ||
      norm === 'subledger' ||
      norm === 'سوب ليدجر' ||
      norm === 'رقم التشغيلة'
    );
  });
}

function isValidNumericSupledger(val: any): boolean {
  const s = (val ?? '').toString().replace(/,/g, '').trim();
  return /^[0-9]{4,}$/.test(s);
}

/**
 * Ensures that EVERY item / ingredient row in the sheet has the Formula Name (اسم التركيبة)
 * populated directly in front of it in the row, and preserves any extracted Supledger number.
 */
export function propagateFormulaNames(
  sheet: SheetData,
  documentContext?: Partial<ExtractedDocument>
): {
  sheet: SheetData;
  updatedCount: number;
  formulaColIndex: number;
} {
  if (!sheet.headers || sheet.headers.length === 0 || !sheet.rows || sheet.rows.length === 0) {
    return { sheet, updatedCount: 0, formulaColIndex: -1 };
  }

  let formulaCol = findFormulaColumnIndex(sheet.headers);
  let headers = [...sheet.headers];
  let rows = sheet.rows.map((r) => [...r]);

  const docFallbackFormula = cleanFormulaTitle(
    documentContext?.documentTitle || sheet.sheetName || ''
  );

  let updatedCount = 0;

  // Case A: A formula column already exists in headers
  if (formulaCol !== -1) {
    let supledgerCol = findExistingSupledgerCol(headers);
    let currentFormula = '';
    let currentSupledger = '';
    const extractedSupledgers: string[] = new Array(rows.length).fill('');
    let anyExtractedSupledger = false;

    rows = rows.map((row, rIdx) => {
      const cellVal = (row[formulaCol] ?? '').toString().trim();
      const rowSup = supledgerCol !== -1 ? (row[supledgerCol] ?? '').toString().trim() : '';

      if (cellVal && !cellVal.includes('إجمالي') && !cellVal.includes('المجموع') && !/total|grand/i.test(cellVal)) {
        const parsed = extractSupledgerAndCleanFormula(cellVal);
        currentFormula = parsed.cleanName;
        if (parsed.supledger) {
          currentSupledger = parsed.supledger;
        } else if (isValidNumericSupledger(rowSup)) {
          currentSupledger = rowSup.replace(/,/g, '');
        } else {
          currentSupledger = '';
        }
        row[formulaCol] = currentFormula;
      } else if (!cellVal) {
        if (currentFormula) {
          row[formulaCol] = currentFormula;
          updatedCount++;
        } else if (docFallbackFormula) {
          row[formulaCol] = docFallbackFormula;
          updatedCount++;
        }
      }

      const finalSup = isValidNumericSupledger(rowSup)
        ? rowSup.replace(/,/g, '')
        : currentSupledger;
      if (finalSup) {
        extractedSupledgers[rIdx] = finalSup;
        anyExtractedSupledger = true;
      }

      return row;
    });

    if (anyExtractedSupledger) {
      if (supledgerCol !== -1) {
        rows.forEach((r, idx) => {
          if (!isValidNumericSupledger(r[supledgerCol]) && extractedSupledgers[idx]) {
            r[supledgerCol] = extractedSupledgers[idx];
          }
        });
      } else {
        // Insert a Supledger column right after formulaCol
        const insertIdx = formulaCol + 1;
        headers.splice(insertIdx, 0, 'Supledger');
        rows.forEach((r, idx) => {
          r.splice(insertIdx, 0, extractedSupledgers[idx] || '');
        });
        if (sheet.totalsRow && sheet.totalsRow.length > 0) {
          const newTotals = [...sheet.totalsRow];
          newTotals.splice(insertIdx, 0, '');
          return {
            sheet: { ...sheet, headers, rows, totalsRow: newTotals },
            updatedCount,
            formulaColIndex: formulaCol,
          };
        }
      }
    }

    return {
      sheet: { ...sheet, headers, rows },
      updatedCount,
      formulaColIndex: formulaCol,
    };
  }

  // Case B: No formula column in headers yet
  const newColName = 'اسم التركيبة';
  headers.unshift(newColName);
  formulaCol = 0;

  let currentFormula = docFallbackFormula || '';
  let currentSupledger = '';
  const newRows: any[][] = [];
  const rowSupledgers: string[] = [];
  let anySupledger = false;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const headerCheck = isFormulaHeaderRow(row);

    if (headerCheck.isHeader && headerCheck.formulaName) {
      currentFormula = headerCheck.formulaName;
      currentSupledger = headerCheck.supledger;
      const hasOtherData = row.slice(2).some((c) => (c ?? '').toString().trim() !== '');
      if (hasOtherData) {
        newRows.push([currentFormula, ...row]);
        rowSupledgers.push(currentSupledger);
        if (currentSupledger) anySupledger = true;
      }
      continue;
    }

    newRows.push([currentFormula, ...row]);
    rowSupledgers.push(currentSupledger);
    if (currentSupledger) anySupledger = true;
    updatedCount++;
  }

  let newTotalsRow = sheet.totalsRow ? ['الإجمالي', ...sheet.totalsRow] : undefined;

  if (anySupledger) {
    headers.splice(1, 0, 'Supledger');
    newRows.forEach((r, idx) => {
      r.splice(1, 0, rowSupledgers[idx] || '');
    });
    if (newTotalsRow) {
      newTotalsRow.splice(1, 0, '');
    }
  }

  return {
    sheet: {
      ...sheet,
      headers,
      rows: newRows,
      totalsRow: newTotalsRow,
    },
    updatedCount,
    formulaColIndex: 0,
  };
}
