import ExcelJS from 'exceljs';
import {
  SadatControlReport,
  SadatFormulaColumn,
  SadatMatrixRow,
} from './sadatControlCalculator';
import {
  ControlPivotColumnKey,
  ControlPivotSummary,
  ControlPivotRow,
  PIVOT_COLUMN_DEFS,
  PivotTotalsGroup,
} from './controlPivotCalculator';

const THIN_BORDER_COLOR = 'FF94A3B8';
const DARK_BORDER_COLOR = 'FF0F172A';

function makeBorder(
  topStyle: ExcelJS.BorderStyle = 'thin',
  bottomStyle: ExcelJS.BorderStyle = 'thin',
  leftStyle: ExcelJS.BorderStyle = 'thin',
  rightStyle: ExcelJS.BorderStyle = 'thin',
  color: string = THIN_BORDER_COLOR
): Partial<ExcelJS.Borders> {
  return {
    top: { style: topStyle, color: { argb: color } },
    bottom: { style: bottomStyle, color: { argb: color } },
    left: { style: leftStyle, color: { argb: color } },
    right: { style: rightStyle, color: { argb: color } },
  };
}

function solidFill(argbHex: string): ExcelJS.Fill {
  return {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: argbHex },
  };
}

async function downloadWorkbook(wb: ExcelJS.Workbook, fileName: string) {
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export interface SadatStyledExportParams {
  report: SadatControlReport;
  regularRows: SadatMatrixRow[];
  waterRows: SadatMatrixRow[];
  displayedRows: SadatMatrixRow[];
  sumsWithoutWater: {
    formulaSums: Record<string, number>;
    fatteningTotal: number;
    fishTotal: number;
    duckTotal: number;
    petsTotal: number;
    grandTotal: number;
  };
  sumsWithWater: {
    formulaSums: Record<string, number>;
    fatteningTotal: number;
    fishTotal: number;
    duckTotal: number;
    petsTotal: number;
    grandTotal: number;
  };
  showDriveCode: boolean;
  useBilingualNames: boolean;
  decimals: number;
  grandTotalDecimals: number;
  displayUnit: string;
  meta: { title: string; date: string; subLabel: string };
  effectiveDate: string;
  batchCodeOverrides: Record<string, string>;
  fileName: string;
}

type SadatColSpec =
  | { kind: 'code'; header: string; width: number }
  | { kind: 'driveCode'; header: string; width: number }
  | { kind: 'name'; header: string; width: number }
  | { kind: 'formula'; formula: SadatFormulaColumn; header: string; width: number }
  | { kind: 'fatteningTotal'; header: string; width: number }
  | { kind: 'fishTotal'; header: string; width: number }
  | { kind: 'duckTotal'; header: string; width: number }
  | { kind: 'petsTotal'; header: string; width: number }
  | { kind: 'grandTotal'; header: string; width: number };

export async function exportSadatControlStyledExcel(params: SadatStyledExportParams) {
  const {
    report,
    regularRows,
    waterRows,
    displayedRows,
    sumsWithoutWater,
    sumsWithWater,
    showDriveCode,
    useBilingualNames,
    decimals,
    grandTotalDecimals,
    displayUnit,
    meta,
    effectiveDate,
    batchCodeOverrides,
    fileName,
  } = params;

  const wb = new ExcelJS.Workbook();
  wb.creator = 'محول PDF إلى Excel الذكي';
  wb.created = new Date();

  const ws = wb.addWorksheet('تقرير_كنترول_السادات', {
    views: [{ rightToLeft: true, state: 'frozen', ySplit: 3 }],
    pageSetup: {
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    },
  });

  const cols: SadatColSpec[] = [{ kind: 'code', header: 'كود الصنف', width: 13 }];
  if (showDriveCode) {
    cols.push({ kind: 'driveCode', header: 'كود الدريف', width: 13 });
  }
  cols.push({ kind: 'name', header: 'اسم الصنف', width: 32 });

  for (const f of report.fatteningFormulas) {
    cols.push({
      kind: 'formula',
      formula: f,
      header: f.label,
      width: Math.min(Math.max(f.label.length + 2, 13), 24),
    });
  }
  if (report.hasFattening) {
    cols.push({ kind: 'fatteningTotal', header: 'اجمالي التسمين', width: 16 });
  }

  for (const f of report.fishFormulas) {
    cols.push({
      kind: 'formula',
      formula: f,
      header: f.label,
      width: Math.min(Math.max(f.label.length + 2, 13), 24),
    });
  }
  if (report.hasFish) {
    cols.push({ kind: 'fishTotal', header: 'اجمالي سمك', width: 15 });
  }

  for (const f of report.duckFormulas) {
    cols.push({
      kind: 'formula',
      formula: f,
      header: f.label,
      width: Math.min(Math.max(f.label.length + 2, 13), 24),
    });
  }
  if (report.hasDuck) {
    cols.push({ kind: 'duckTotal', header: 'اجمالي البط', width: 15 });
  }

  for (const f of report.petsFormulas) {
    cols.push({
      kind: 'formula',
      formula: f,
      header: f.label,
      width: Math.min(Math.max(f.label.length + 2, 13), 24),
    });
  }
  if (report.hasPets) {
    cols.push({ kind: 'petsTotal', header: 'اجمالي الاليف', width: 15 });
  }

  cols.push({ kind: 'grandTotal', header: 'الإجمالي العام', width: 16 });

  ws.columns = cols.map((c) => ({ width: c.width }));
  const totalCols = cols.length;
  const fixedMetaColCount = showDriveCode ? 3 : 2;
  const numFmtQty = decimals === 6 ? '0.000000' : '0.000';
  const numFmtGrand = grandTotalDecimals === 4 ? '0.0000' : '0.000';

  // --- ROW 1: Top Yellow Banner ---
  const row1 = ws.addRow(new Array(totalCols).fill(''));
  row1.height = 52;
  for (let c = 1; c <= totalCols; c++) {
    const cell = row1.getCell(c);
    cell.fill = solidFill('FFFFFF00');
    cell.border = makeBorder('medium', 'medium', 'thin', 'thin', DARK_BORDER_COLOR);
  }

  const rightCardEnd = Math.min(2, totalCols);
  if (rightCardEnd >= 2) {
    ws.mergeCells(1, 1, 1, rightCardEnd);
  }
  const rightCard = row1.getCell(1);
  rightCard.value = 'شركة الدقهلية للدواجن\nمصنع أعلاف السادات\nإدارة المخازن\nمخزن الخامات';
  rightCard.fill = solidFill('FFFFFFFF');
  rightCard.font = { name: 'Cairo', size: 9, bold: true, color: { argb: 'FF0F172A' } };
  rightCard.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  rightCard.border = makeBorder('medium', 'medium', 'medium', 'thin', DARK_BORDER_COLOR);

  const leftCardStart = totalCols >= 6 ? totalCols - 1 : totalCols;
  if (leftCardStart < totalCols) {
    ws.mergeCells(1, leftCardStart, 1, totalCols);
  }
  const leftCard = row1.getCell(leftCardStart);
  leftCard.value = 'DakahliaPoultry';
  leftCard.fill = solidFill('FFFFFFFF');
  leftCard.font = { name: 'Cairo', size: 11, bold: true, color: { argb: 'FF002060' } };
  leftCard.alignment = { vertical: 'middle', horizontal: 'center' };
  leftCard.border = makeBorder('medium', 'medium', 'thin', 'medium', DARK_BORDER_COLOR);

  const titleStart = rightCardEnd + 1;
  const titleEnd = leftCardStart - 1;
  if (titleEnd >= titleStart) {
    if (titleEnd > titleStart) {
      ws.mergeCells(1, titleStart, 1, titleEnd);
    }
    const titleCell = row1.getCell(titleStart);
    const totalFormulas =
      report.fatteningFormulas.length +
      report.fishFormulas.length +
      report.duckFormulas.length +
      report.petsFormulas.length;
    titleCell.value = `${meta.title}\nالوحدة: ${displayUnit}  |  عدد التركيبات: ${totalFormulas}  |  عدد الخامات: ${displayedRows.length}`;
    titleCell.fill = solidFill('FFFFFF00');
    titleCell.font = { name: 'Cairo', size: 15, bold: true, color: { argb: 'FF020617' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  }

  // --- ROW 2: White Supledger & Date Row ---
  const row2 = ws.addRow(new Array(totalCols).fill(''));
  row2.height = 22;
  for (let c = 1; c <= totalCols; c++) {
    const cell = row2.getCell(c);
    cell.fill = solidFill('FFFFFFFF');
    cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF0F172A' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = makeBorder('thin', 'medium', 'thin', 'thin', DARK_BORDER_COLOR);
  }

  ws.mergeCells(2, 1, 2, fixedMetaColCount);
  row2.getCell(1).value = meta.subLabel || 'supledger';

  cols.forEach((col, idx) => {
    const cell = row2.getCell(idx + 1);
    if (col.kind === 'formula') {
      const rawVal = (batchCodeOverrides[col.formula.key] ?? col.formula.productCode ?? '')
        .replace(/,/g, '')
        .trim();
      if (/^[0-9]+$/.test(rawVal)) {
        cell.value = Number(rawVal);
        cell.numFmt = '0';
      } else {
        cell.value = rawVal || '';
      }
    } else if (col.kind === 'grandTotal') {
      cell.value = effectiveDate;
    }
  });

  // --- ROW 3: Dark Navy Table Headers ---
  const row3 = ws.addRow(cols.map((c) => c.header));
  row3.height = 34;
  cols.forEach((col, idx) => {
    const cell = row3.getCell(idx + 1);
    cell.fill = solidFill('FF002060');
    let fontColor = 'FFFFFFFF';
    if (col.kind === 'driveCode') fontColor = 'FFC7D2FE';
    else if (col.kind === 'fatteningTotal') fontColor = 'FFFFFF00';
    else if (col.kind === 'fishTotal') fontColor = 'FFFDE68A';
    else if (col.kind === 'duckTotal') fontColor = 'FF99F6E4';
    else if (col.kind === 'petsTotal') fontColor = 'FFE9D5FF';
    else if (col.kind === 'grandTotal') fontColor = 'FFA7F3D0';
    else if (col.kind === 'formula') {
      if (col.formula.itemType === 'تسمين') fontColor = 'FFFFFF00';
      else if (col.formula.itemType === 'سمك') fontColor = 'FFBAE6FD';
      else if (col.formula.itemType === 'بط') fontColor = 'FF99F6E4';
      else fontColor = 'FFE9D5FF';
    }
    cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: fontColor } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = makeBorder('medium', 'medium', 'thin', 'thin', 'FF475569');
  });

  const addMatrixDataRow = (r: SadatMatrixRow, isWater: boolean) => {
    const values = cols.map((col) => {
      switch (col.kind) {
        case 'code':
          return r.code || '';
        case 'driveCode':
          return r.driveCode || '';
        case 'name':
          return useBilingualNames ? r.bilingualName : r.name;
        case 'formula': {
          const v = r.formulaValues[col.formula.key] || 0;
          return Math.abs(v) < 1e-9 ? '-' : Number(v.toFixed(decimals));
        }
        case 'fatteningTotal':
          return Math.abs(r.fatteningTotal) < 1e-9 ? '-' : Number(r.fatteningTotal.toFixed(decimals));
        case 'fishTotal':
          return Math.abs(r.fishTotal) < 1e-9 ? '-' : Number(r.fishTotal.toFixed(decimals));
        case 'duckTotal':
          return Math.abs(r.duckTotal) < 1e-9 ? '-' : Number(r.duckTotal.toFixed(decimals));
        case 'petsTotal':
          return Math.abs(r.petsTotal) < 1e-9 ? '-' : Number(r.petsTotal.toFixed(decimals));
        case 'grandTotal':
          return Number(r.grandTotal.toFixed(grandTotalDecimals));
      }
    });

    const row = ws.addRow(values);
    row.height = 20;

    cols.forEach((col, idx) => {
      const cell = row.getCell(idx + 1);
      const val = values[idx];
      cell.border = makeBorder('thin', 'thin', 'thin', 'thin', THIN_BORDER_COLOR);

      if (col.kind === 'code') {
        cell.fill = solidFill(isWater ? 'FFE0F2FE' : 'FFFFFFFF');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF0F172A' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (col.kind === 'driveCode') {
        cell.fill = solidFill(isWater ? 'FFDBEAFE' : 'FFEEF2FF');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF312E81' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (col.kind === 'name') {
        cell.fill = solidFill(isWater ? 'FFE0F2FE' : 'FFFFFFFF');
        cell.font = {
          name: 'Cairo',
          size: 10,
          bold: true,
          color: { argb: isWater ? 'FF0C4A6E' : 'FF0F172A' },
        };
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (col.kind === 'fatteningTotal') {
        cell.fill = solidFill('FFFFE699');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF020617' } };
        cell.alignment = { vertical: 'middle', horizontal: typeof val === 'number' ? 'left' : 'center' };
        if (typeof val === 'number') cell.numFmt = numFmtQty;
      } else if (col.kind === 'fishTotal') {
        cell.fill = solidFill('FFF4B084');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF020617' } };
        cell.alignment = { vertical: 'middle', horizontal: typeof val === 'number' ? 'left' : 'center' };
        if (typeof val === 'number') cell.numFmt = numFmtQty;
      } else if (col.kind === 'duckTotal') {
        cell.fill = solidFill('FFB4C6E7');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF020617' } };
        cell.alignment = { vertical: 'middle', horizontal: typeof val === 'number' ? 'left' : 'center' };
        if (typeof val === 'number') cell.numFmt = numFmtQty;
      } else if (col.kind === 'petsTotal') {
        cell.fill = solidFill('FFE2D4F0');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF020617' } };
        cell.alignment = { vertical: 'middle', horizontal: typeof val === 'number' ? 'left' : 'center' };
        if (typeof val === 'number') cell.numFmt = numFmtQty;
      } else if (col.kind === 'grandTotal') {
        cell.fill = solidFill('FFA9D08E');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF020617' } };
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        cell.numFmt = numFmtGrand;
      } else {
        cell.fill = solidFill(isWater ? 'FFE0F2FE' : 'FFFFFFFF');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF0F172A' } };
        cell.alignment = { vertical: 'middle', horizontal: typeof val === 'number' ? 'left' : 'center' };
        if (typeof val === 'number') cell.numFmt = numFmtQty;
      }
    });
  };

  const addMatrixSummaryRow = (
    label: string,
    countVal: number,
    sums: SadatStyledExportParams['sumsWithWater'],
    isFinalGrandTotal: boolean
  ) => {
    const values = cols.map((col) => {
      switch (col.kind) {
        case 'code':
          return countVal;
        case 'driveCode':
          return '';
        case 'name':
          return label;
        case 'formula':
          return Number((sums.formulaSums[col.formula.key] || 0).toFixed(decimals));
        case 'fatteningTotal':
          return Number(sums.fatteningTotal.toFixed(decimals));
        case 'fishTotal':
          return Number(sums.fishTotal.toFixed(decimals));
        case 'duckTotal':
          return Number(sums.duckTotal.toFixed(decimals));
        case 'petsTotal':
          return Number(sums.petsTotal.toFixed(decimals));
        case 'grandTotal':
          return Number(sums.grandTotal.toFixed(grandTotalDecimals));
      }
    });

    const row = ws.addRow(values);
    row.height = 24;
    const baseFill = isFinalGrandTotal ? 'FFFFD966' : 'FFFEF3C7';

    cols.forEach((col, idx) => {
      const cell = row.getCell(idx + 1);
      cell.border = makeBorder('medium', 'medium', 'thin', 'thin', DARK_BORDER_COLOR);
      cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF020617' } };

      let fill = baseFill;
      if (col.kind === 'fatteningTotal') fill = 'FFFFD966';
      else if (col.kind === 'fishTotal') fill = 'FFE89968';
      else if (col.kind === 'duckTotal') fill = 'FF8FAADC';
      else if (col.kind === 'petsTotal') fill = 'FFC9B2E3';
      else if (col.kind === 'grandTotal') fill = 'FF8EC26D';

      cell.fill = solidFill(fill);

      if (col.kind === 'code') {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (col.kind === 'name' || col.kind === 'driveCode') {
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        cell.numFmt = col.kind === 'grandTotal' ? numFmtGrand : numFmtQty;
      }
    });
  };

  regularRows.forEach((r) => addMatrixDataRow(r, false));

  if (waterRows.length > 0) {
    addMatrixSummaryRow(`الإجمالي (بدون الماء) (${displayUnit})`, regularRows.length, sumsWithoutWater, false);
    waterRows.forEach((w) => addMatrixDataRow(w, true));
    addMatrixSummaryRow(`الإجمالي الكلي (بالماء) (${displayUnit})`, displayedRows.length, sumsWithWater, true);
  } else {
    addMatrixSummaryRow(`الإجمالي الكلي (${displayUnit})`, displayedRows.length, sumsWithWater, true);
  }

  const baseName = fileName.replace(/\.(xlsx|pdf|csv)$/i, '');
  await downloadWorkbook(wb, `${baseName}_كنترول_السادات.xlsx`);
}

export interface ControlPivotStyledExportParams {
  pivot: ControlPivotSummary;
  columnOrder: ControlPivotColumnKey[];
  regularRows: ControlPivotRow[];
  waterRows: ControlPivotRow[];
  displayedRows: ControlPivotRow[];
  totalsWithoutWater: PivotTotalsGroup;
  totalsWithWater: PivotTotalsGroup;
  decimals: number;
  displayUnitLabel: string;
  matchedCount: number;
  regularMatchedCount: number;
  filteredMatchedCount: number;
  fileName: string;
}

export async function exportControlPivotStyledExcel(params: ControlPivotStyledExportParams) {
  const {
    columnOrder,
    regularRows,
    waterRows,
    displayedRows,
    totalsWithoutWater,
    totalsWithWater,
    decimals,
    displayUnitLabel,
    matchedCount,
    regularMatchedCount,
    filteredMatchedCount,
    fileName,
  } = params;

  const wb = new ExcelJS.Workbook();
  wb.creator = 'محول PDF إلى Excel الذكي';
  wb.created = new Date();

  const ws = wb.addWorksheet('تجميع_الكنترول', {
    views: [{ rightToLeft: true, state: 'frozen', ySplit: 2 }],
    pageSetup: {
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    },
  });

  ws.columns = columnOrder.map((key) => ({
    width: Math.max(PIVOT_COLUMN_DEFS[key].width, 14),
  }));

  const numFmt = decimals === 6 ? '0.000000' : '0.000';
  const totalCols = columnOrder.length;

  // --- ROW 1: Report Title Banner ---
  const titleRow = ws.addRow(new Array(totalCols).fill(''));
  titleRow.height = 28;
  ws.mergeCells(1, 1, 1, totalCols);
  const titleCell = titleRow.getCell(1);
  titleCell.value = `صفحة تجميع الكنترول حسب الصنف ونوعه  —  الوحدة: ${displayUnitLabel}  (${matchedCount} من ${displayedRows.length} صنف)`;
  titleCell.fill = solidFill('FFF1F5F9');
  titleCell.font = { name: 'Cairo', size: 12, bold: true, color: { argb: 'FF002060' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  for (let c = 1; c <= totalCols; c++) {
    titleRow.getCell(c).border = makeBorder('medium', 'thin', 'thin', 'thin', DARK_BORDER_COLOR);
  }

  // --- ROW 2: Dark Navy Headers ---
  const headers = columnOrder.map((k) => PIVOT_COLUMN_DEFS[k].label);
  const headerRow = ws.addRow(headers);
  headerRow.height = 28;
  columnOrder.forEach((colKey, idx) => {
    const cell = headerRow.getCell(idx + 1);
    const isNumericCol =
      colKey === 'fattening' ||
      colKey === 'fish' ||
      colKey === 'duck' ||
      colKey === 'pets' ||
      colKey === 'total';
    cell.fill = solidFill('FF002060');
    cell.font = {
      name: 'Cairo',
      size: 11,
      bold: true,
      color: { argb: isNumericCol ? 'FFFCD34D' : 'FFFFFFFF' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = makeBorder('medium', 'medium', 'thin', 'thin', 'FF475569');
  });

  const addPivotDataRow = (row: ControlPivotRow, isWater: boolean) => {
    const values = columnOrder.map((colKey) => {
      switch (colKey) {
        case 'driveCode':
          return row.driveCode || '';
        case 'code':
          return row.code || '';
        case 'name':
          return row.name;
        case 'fattening':
        case 'fish':
        case 'duck':
        case 'pets':
          return Math.abs(row[colKey]) < 1e-9 ? '-' : Number(row[colKey].toFixed(decimals));
        case 'total':
          return Number(row.total.toFixed(decimals));
      }
    });

    const excelRow = ws.addRow(values);
    excelRow.height = 20;

    columnOrder.forEach((colKey, idx) => {
      const cell = excelRow.getCell(idx + 1);
      const val = values[idx];
      cell.border = makeBorder('thin', 'thin', 'thin', 'thin', THIN_BORDER_COLOR);

      if (colKey === 'driveCode') {
        cell.fill = solidFill(isWater ? 'FFDBEAFE' : 'FFEEF2FF');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF312E81' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (colKey === 'code') {
        cell.fill = solidFill(isWater ? 'FFE0F2FE' : 'FFFFFFFF');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF1E293B' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (colKey === 'name') {
        cell.fill = solidFill(isWater ? 'FFE0F2FE' : 'FFFFFFFF');
        cell.font = {
          name: 'Cairo',
          size: 10,
          bold: true,
          color: { argb: isWater ? 'FF0C4A6E' : 'FF0F172A' },
        };
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (colKey === 'total') {
        cell.fill = solidFill(isWater ? 'FFDBEAFE' : 'FFFFF2CC');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF020617' } };
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        cell.numFmt = numFmt;
      } else {
        cell.fill = solidFill(isWater ? 'FFE0F2FE' : 'FFFFFFFF');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF0F172A' } };
        cell.alignment = {
          vertical: 'middle',
          horizontal: typeof val === 'number' ? 'left' : 'center',
        };
        if (typeof val === 'number') cell.numFmt = numFmt;
      }
    });
  };

  const addPivotSummaryRow = (
    label: string,
    totalsGroup: PivotTotalsGroup,
    matchedCnt: number,
    totalCnt: number,
    isGrandFinal: boolean
  ) => {
    const values = columnOrder.map((colKey) => {
      switch (colKey) {
        case 'driveCode':
          return `${matchedCnt}/${totalCnt}`;
        case 'code':
          return totalCnt;
        case 'name':
          return label;
        case 'fattening':
        case 'fish':
        case 'duck':
        case 'pets':
          return Math.abs(totalsGroup[colKey]) < 1e-9
            ? '-'
            : Number(totalsGroup[colKey].toFixed(decimals));
        case 'total':
          return Number(totalsGroup.grandTotal.toFixed(decimals));
      }
    });

    const excelRow = ws.addRow(values);
    excelRow.height = 24;

    columnOrder.forEach((colKey, idx) => {
      const cell = excelRow.getCell(idx + 1);
      const val = values[idx];
      cell.border = makeBorder('medium', 'medium', 'thin', 'thin', DARK_BORDER_COLOR);

      if (colKey === 'total') {
        cell.fill = solidFill(isGrandFinal ? 'FFF59E0B' : 'FFFDE68A');
        cell.font = { name: 'Cairo', size: 11, bold: true, color: { argb: 'FF020617' } };
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        cell.numFmt = numFmt;
      } else {
        cell.fill = solidFill(isGrandFinal ? 'FFFDE68A' : 'FFFEF3C7');
        cell.font = { name: 'Cairo', size: 10, bold: true, color: { argb: 'FF0F172A' } };
        if (colKey === 'name') {
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if (colKey === 'code' || colKey === 'driveCode') {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        } else {
          cell.alignment = {
            vertical: 'middle',
            horizontal: typeof val === 'number' ? 'left' : 'center',
          };
          if (typeof val === 'number') cell.numFmt = numFmt;
        }
      }
    });
  };

  regularRows.forEach((r) => addPivotDataRow(r, false));

  if (waterRows.length > 0) {
    addPivotSummaryRow(
      `الإجمالي (بدون الماء) (${displayUnitLabel})`,
      totalsWithoutWater,
      regularMatchedCount,
      regularRows.length,
      false
    );
    waterRows.forEach((w) => addPivotDataRow(w, true));
    addPivotSummaryRow(
      `الإجمالي الكلي (بالماء) (${displayUnitLabel})`,
      totalsWithWater,
      filteredMatchedCount,
      displayedRows.length,
      true
    );
  } else {
    addPivotSummaryRow(
      `الإجمالي الكلي (${displayUnitLabel})`,
      totalsWithWater,
      filteredMatchedCount,
      displayedRows.length,
      true
    );
  }

  const baseName = fileName.replace(/\.(xlsx|pdf|csv)$/i, '');
  await downloadWorkbook(wb, `${baseName}_تجميع_الكنترول.xlsx`);
}

export interface DriveSummaryExportRow {
  code: string;
  driveCode: string;
  displayName: string;
  total: number;
  isWater?: boolean;
}

export interface DriveSummaryStyledExportParams {
  regularRows: DriveSummaryExportRow[];
  waterRows: DriveSummaryExportRow[];
  displayedRows: DriveSummaryExportRow[];
  totalWithoutWater: number;
  totalWithWater: number;
  separateWater: boolean;
  decimals: number;
  fileName: string;
}

/**
 * Exports the 4-column Drive Summary Sheet (كود الصنف | كود درايف | إسم الصنف | الاجمالى)
 * matching the bright yellow header (#FFFF00), black borders, centered bold cells.
 */
export async function exportDriveSummaryStyledExcel({
  regularRows,
  waterRows,
  displayedRows,
  totalWithoutWater,
  totalWithWater,
  separateWater,
  decimals,
  fileName,
}: DriveSummaryStyledExportParams): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'محول PDF إلى Excel الذكي';
  wb.created = new Date();

  const ws = wb.addWorksheet('ملخص_اكواد_الدرايف', {
    views: [{ rightToLeft: true, state: 'frozen', ySplit: 1 }],
  });

  ws.columns = [
    { width: 16 }, // كود الصنف
    { width: 14 }, // كود درايف
    { width: 38 }, // إسم الصنف
    { width: 20 }, // الاجمالى
  ];

  const headerRow = ws.addRow(['كود الصنف', 'كود درايف', 'إسم الصنف', 'الاجمالى']);
  headerRow.height = 26;
  headerRow.eachCell((cell) => {
    cell.fill = solidFill('FFFFFF00');
    cell.font = { name: 'Cairo', size: 12, bold: true, color: { argb: 'FF000000' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = makeBorder('medium', 'medium', 'thin', 'thin', 'FF000000');
  });

  const numFmt = '0.########';

  const appendDataRow = (r: DriveSummaryExportRow, isWaterRow: boolean) => {
    const numTotal = Math.abs(r.total) < 1e-12 ? 0 : Number(r.total.toFixed(decimals));
    const parsedCode = r.code && /^\d+$/.test(r.code.trim()) ? Number(r.code.trim()) : r.code || '';
    const parsedDrive =
      r.driveCode && /^\d+$/.test(r.driveCode.trim()) ? Number(r.driveCode.trim()) : r.driveCode || '';

    const excelRow = ws.addRow([parsedCode, parsedDrive, r.displayName, numTotal]);
    excelRow.height = 21;

    excelRow.eachCell((cell, colNumber) => {
      cell.fill = solidFill(isWaterRow ? 'FFE0F2FE' : 'FFFFFFFF');
      cell.font = { name: 'Cairo', size: 11, bold: true, color: { argb: 'FF000000' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = makeBorder('thin', 'thin', 'thin', 'thin', 'FF000000');
      if (colNumber === 4) {
        cell.numFmt = numFmt;
      }
    });
  };

  const appendSummaryRow = (label: string, totalVal: number, isGrand: boolean) => {
    const numTotal = Math.abs(totalVal) < 1e-12 ? 0 : Number(totalVal.toFixed(decimals));
    const excelRow = ws.addRow(['', '', label, numTotal]);
    excelRow.height = 24;

    excelRow.eachCell((cell, colNumber) => {
      cell.fill = solidFill(isGrand ? 'FFFFFF00' : 'FFFEF3C7');
      cell.font = { name: 'Cairo', size: 11, bold: true, color: { argb: 'FF000000' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = makeBorder('medium', 'medium', 'thin', 'thin', 'FF000000');
      if (colNumber === 4) {
        cell.numFmt = numFmt;
      }
    });
  };

  if (separateWater && waterRows.length > 0) {
    regularRows.forEach((r) => appendDataRow(r, false));
    appendSummaryRow('الإجمالي (بدون الماء)', totalWithoutWater, false);
    waterRows.forEach((w) => appendDataRow(w, true));
    appendSummaryRow('الإجمالي الكلي (بالماء)', totalWithWater, true);
  } else {
    displayedRows.forEach((r) => appendDataRow(r, Boolean(r.isWater)));
    appendSummaryRow('الإجمالي الكلي', totalWithWater, true);
  }

  const baseName = fileName.replace(/\.(xlsx|pdf|csv)$/i, '');
  await downloadWorkbook(wb, `${baseName}_ملخص_الدرايف.xlsx`);
}


export interface DailySpendExportRow {
  en: string;
  ar: string;
  qty: number;
  hasData: boolean;
}

export interface DailySpendStyledExportParams {
  rows: DailySpendExportRow[];
  total: number;
  decimals: number;
  title: string;
  qtyHeader: string;
  fileName: string;
}

/** Styled Excel export that mirrors the daily "الصرف اليومي للكنترول" sheet (English | quantity | Arabic). */
export async function exportDailySpendStyledExcel({
  rows,
  total,
  decimals,
  title,
  qtyHeader,
  fileName,
}: DailySpendStyledExportParams): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'TOTAL MIXING REPORT Converter';
  wb.created = new Date();

  const ws = wb.addWorksheet('الصرف_اليومي', {
    views: [{ rightToLeft: true, state: 'frozen', ySplit: 1 }],
  });
  ws.columns = [{ width: 32 }, { width: 22 }, { width: 28 }];

  const headerRow = ws.addRow([title, qtyHeader, 'الصنف']);
  headerRow.height = 30;
  headerRow.eachCell((cell) => {
    cell.fill = solidFill('FFFFFF00');
    cell.font = { name: 'Cairo', size: 12, bold: true, color: { argb: 'FF000000' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = makeBorder('medium', 'medium', 'medium', 'medium', 'FF000000');
  });

  const numFmt = '0.########';
  rows.forEach((r) => {
    const val = r.hasData ? Number(r.qty.toFixed(decimals)) : '-';
    const excelRow = ws.addRow([r.en, val, r.ar]);
    excelRow.height = 22;
    excelRow.eachCell((cell, colNumber) => {
      cell.fill = solidFill(colNumber === 2 ? 'FFFFFF00' : colNumber === 1 ? 'FFFDEBD0' : 'FFFFFFFF');
      cell.font = { name: 'Cairo', size: 11, bold: true, color: { argb: 'FF000000' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = makeBorder('thin', 'thin', 'thin', 'thin', 'FF000000');
      if (colNumber === 2 && typeof val === 'number') cell.numFmt = numFmt;
    });
  });

  const totalRow = ws.addRow(['الإجمالي', Number(total.toFixed(decimals)), '']);
  totalRow.height = 26;
  totalRow.eachCell((cell, colNumber) => {
    cell.fill = solidFill('FFDCFCE7');
    cell.font = { name: 'Cairo', size: 12, bold: true, color: { argb: 'FF000000' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = makeBorder('medium', 'medium', 'medium', 'medium', 'FF000000');
    if (colNumber === 2) cell.numFmt = numFmt;
  });

  await downloadWorkbook(wb, fileName);
}
