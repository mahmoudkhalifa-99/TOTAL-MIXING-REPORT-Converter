import React, { useState, useMemo } from 'react';
import {
  Download,
  Copy,
  Check,
  Search,
  Plus,
  Trash2,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Layers,
  FileSpreadsheet,
  FileText,
  Edit2,
  RefreshCw,
  Scissors,
  Sparkles,
  CheckCircle2,
  Tag,
  Filter,
  Info,
  Scale,
  Calculator,
  PieChart,
  GripVertical,
  Columns3,
  ChevronRight,
  ChevronLeft,
  TableProperties,
  ZoomIn,
  ZoomOut,
  ListOrdered,
} from 'lucide-react';
import { SheetData, ExtractedDocument } from '../types/converter';
import {
  exportToExcel,
  exportToCsv,
  copyTableToClipboard,
} from '../utils/excelExport';
import { autoSplitCodeColumn } from '../utils/codeSplitter';
import { propagateFormulaNames } from '../utils/formulaPropagator';
import {
  ensureItemTypeColumn,
  findItemTypeColumnIndex,
  findFormulaColumnIndex,
  classifyFormulaItemType,
} from '../utils/itemClassifier';
import { CategoryTotalsBar } from './CategoryTotalsBar';
import { ControlPivotView } from './ControlPivotView';
import { SadatControlReportView } from './SadatControlReportView';
import { DriveSummaryView } from './DriveSummaryView';
import { DailySpendView } from './DailySpendView';
import {
  calculateCategoryTotals,
  formatWeight,
  OverallTotalsSummary,
} from '../utils/categoryTotals';

interface ExcelGridProps {
  document: ExtractedDocument;
  fileName?: string;
  onUpdateDocument: (updated: ExtractedDocument) => void;
  onClearData?: () => void;
  pdfDataUrl?: string;
}

const FONT_SIZE_STORAGE_KEY = 'excel_grid_font_step_v1';

const FONT_SIZE_STEPS = [
  { px: 10, tableClass: 'text-[10px]' },
  { px: 11, tableClass: 'text-[11px]' },
  { px: 12, tableClass: 'text-xs' },
  { px: 13, tableClass: 'text-[13px]' },
  { px: 14, tableClass: 'text-sm' },
  { px: 15, tableClass: 'text-[15px]' },
  { px: 16, tableClass: 'text-base' },
  { px: 18, tableClass: 'text-lg' },
  { px: 20, tableClass: 'text-xl' },
] as const;

const DEFAULT_FONT_STEP_INDEX = 3; // 13px

export const ExcelGrid: React.FC<ExcelGridProps> = ({
  document,
  fileName = 'converted_document.xlsx',
  onUpdateDocument,
  onClearData,
  pdfDataUrl,
}) => {
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [copied, setCopied] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [sortConfig, setSortConfig] = useState<{
    columnIndex: number;
    direction: 'asc' | 'desc';
  } | null>(null);

  // Cell editing state
  const [editingCell, setEditingCell] = useState<{
    rowIndex: number;
    colIndex: number;
  } | null>(null);
  const [editValue, setEditValue] = useState('');

  // Header editing state
  const [editingHeaderIndex, setEditingHeaderIndex] = useState<number | null>(null);
  const [editHeaderValue, setEditHeaderValue] = useState('');

  // Selected row for actions
  const [selectedRowIndex, setSelectedRowIndex] = useState<number | null>(null);
  const [splitToast, setSplitToast] = useState<string | null>(null);
  const [categoryFilter, setCategoryFilter] = useState<string>('all');
  const [showItemTypeGuide, setShowItemTypeGuide] = useState<boolean>(false);

  // Tab view: 'grid' or 'split' (showing PDF alongside table)
  const [viewMode, setViewMode] = useState<'grid' | 'split'>('grid');
  // Page view: 'details' (detailed spreadsheet), 'control_pivot' (by type), 'sadat_control' (formula matrix), or 'drive_summary' (4-col drive code summary)
  const [activePage, setActivePage] = useState<'details' | 'control_pivot' | 'sadat_control' | 'drive_summary' | 'daily_spend'>('details');
  // Column reordering UI & drag state
  const [showColumnOrderBar, setShowColumnOrderBar] = useState(false);
  const [draggedColIndex, setDraggedColIndex] = useState<number | null>(null);
  const [dragOverColIndex, setDragOverColIndex] = useState<number | null>(null);

  // Table font size step (shared across all 3 pages)
  const [fontStepIndex, setFontStepIndex] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(FONT_SIZE_STORAGE_KEY);
      if (saved !== null) {
        const parsed = parseInt(saved, 10);
        if (!isNaN(parsed) && parsed >= 0 && parsed < FONT_SIZE_STEPS.length) {
          return parsed;
        }
      }
    } catch {
      // ignore storage errors
    }
    return DEFAULT_FONT_STEP_INDEX;
  });

  const updateFontStep = (nextIdx: number) => {
    const clamped = Math.max(0, Math.min(FONT_SIZE_STEPS.length - 1, nextIdx));
    setFontStepIndex(clamped);
    try {
      localStorage.setItem(FONT_SIZE_STORAGE_KEY, String(clamped));
    } catch {
      // ignore storage errors
    }
  };

  const currentFontStep = FONT_SIZE_STEPS[fontStepIndex] || FONT_SIZE_STEPS[DEFAULT_FONT_STEP_INDEX];

  const currentSheet = document.sheets[activeSheetIndex] || document.sheets[0];

  // Detect index of "نوع الصنف" column if present
  const itemTypeColIndex = useMemo(() => {
    if (!currentSheet || !currentSheet.headers) return -1;
    return findItemTypeColumnIndex(currentSheet.headers);
  }, [currentSheet]);

  // Count items per category (تسمين، سمك، بط، اليف)
  const categoryCounts = useMemo(() => {
    if (!currentSheet || itemTypeColIndex === -1) return null;
    const counts: Record<string, number> = {
      all: currentSheet.rows.length,
      'سمك': 0,
      'تسمين': 0,
      'بط': 0,
      'اليف': 0,
      'تنظيف': 0,
    };

    currentSheet.rows.forEach((row) => {
      const val = (row[itemTypeColIndex] ?? '').toString().trim();
      if (val === 'سمك' || val === 'تسمين' || val === 'بط' || val === 'اليف' || val === 'تنظيف') {
        counts[val]++;
      } else if (val) {
        counts['اليف']++;
      }
    });

    return counts;
  }, [currentSheet, itemTypeColIndex]);

  // Overall totals calculation for the 4 categories (سمك، تسمين، بط، اليف) & Grand Total
  const totalsSummary: OverallTotalsSummary = useMemo(() => {
    return calculateCategoryTotals(currentSheet);
  }, [currentSheet]);

  const isTotalsRowAppended = useMemo(() => {
    if (!currentSheet || !currentSheet.rows) return false;
    return currentSheet.rows.some((row) =>
      row.some(
        (cell) =>
          typeof cell === 'string' &&
          (cell.includes('إجمالي السمك') || cell.includes('الإجمالي الكلي للأربعة'))
      )
    );
  }, [currentSheet]);

  // Append or remove category totals rows directly in current sheet
  const handlePermanentlyAppendTotals = () => {
    if (!currentSheet) return;

    if (isTotalsRowAppended) {
      // Remove previously appended category totals rows
      const cleanedRows = currentSheet.rows.filter(
        (row) =>
          !row.some(
            (cell) =>
              typeof cell === 'string' &&
              (cell.includes('إجمالي السمك') ||
                cell.includes('إجمالي التسمين') ||
                cell.includes('إجمالي البط') ||
                cell.includes('إجمالي الأليف') ||
                cell.includes('الإجمالي الكلي للأربعة'))
          )
      );
      const newSheets = [...document.sheets];
      newSheets[activeSheetIndex] = { ...currentSheet, rows: cleanedRows };
      onUpdateDocument({ ...document, sheets: newSheets });
      setSplitToast('تمت إزالة صفوف مجاميع الأنواع الأربعة من الجدول.');
      setTimeout(() => setSplitToast(null), 3000);
      return;
    }

    const findLabelColumnIndex = (headers: string[]): number => {
      const idx = findFormulaColumnIndex(headers);
      if (idx !== -1) return idx;
      const fallback = headers.findIndex((h) => {
        const norm = (h ?? '').toString().trim().toLowerCase();
        if (!norm) return false;
        return !(
          norm.includes('code') ||
          norm.includes('كود') ||
          norm.includes('رمز') ||
          norm.includes('رقم') ||
          norm.includes('supledger') ||
          norm.includes('subledger') ||
          norm.includes('نوع الصنف') ||
          norm.includes('نوع_الصنف') ||
          norm.includes('target') ||
          norm.includes('actual') ||
          norm.includes('diff') ||
          norm.includes('المستهدف') ||
          norm.includes('الفعلي') ||
          norm.includes('الكمية') ||
          norm.includes('الوزن')
        );
      });
      return fallback !== -1 ? fallback : 0;
    };

    const targetIdx = totalsSummary.targetColIndex;
    const actualIdx = totalsSummary.actualColIndex;
    const itemTypeIdx = itemTypeColIndex;
    const labelCol = findLabelColumnIndex(currentSheet.headers);

    const makeSubtotalRow = (
      label: string,
      typeVal: string,
      targetVal: number,
      actualVal: number
    ) => {
      const row = new Array(currentSheet.headers.length).fill('');
      row[labelCol] = label;
      if (itemTypeIdx !== -1) row[itemTypeIdx] = typeVal;
      if (targetIdx !== -1) row[targetIdx] = targetVal.toFixed(3);
      if (actualIdx !== -1) row[actualIdx] = actualVal.toFixed(3);
      return row;
    };

    const subtotalRows: any[][] = [
      makeSubtotalRow(
        '🐟 إجمالي السمك (Fish Total)',
        'سمك',
        totalsSummary.categories['سمك'].targetSum,
        totalsSummary.categories['سمك'].actualSum
      ),
      makeSubtotalRow(
        '🌾 إجمالي التسمين (Broiler Total)',
        'تسمين',
        totalsSummary.categories['تسمين'].targetSum,
        totalsSummary.categories['تسمين'].actualSum
      ),
      makeSubtotalRow(
        '🦆 إجمالي البط (Duck Total)',
        'بط',
        totalsSummary.categories['بط'].targetSum,
        totalsSummary.categories['بط'].actualSum
      ),
      makeSubtotalRow(
        '🐾 إجمالي الأليف (Pets Total)',
        'اليف',
        totalsSummary.categories['اليف'].targetSum,
        totalsSummary.categories['اليف'].actualSum
      ),
      makeSubtotalRow(
        '🧹 Cleaning (تنظيف) - غير مشمول بالإجمالي الكلي',
        'تنظيف',
        totalsSummary.categories['تنظيف'].targetSum,
        totalsSummary.categories['تنظيف'].actualSum
      ),
      makeSubtotalRow(
        '⭐ الإجمالي الكلي للأربعة (Grand Total)',
        'الكل',
        totalsSummary.grandTotal.targetSum,
        totalsSummary.grandTotal.actualSum
      ),
    ];

    const newSheets = [...document.sheets];
    newSheets[activeSheetIndex] = {
      ...currentSheet,
      rows: [...currentSheet.rows, ...subtotalRows],
    };
    onUpdateDocument({ ...document, sheets: newSheets });
    setSplitToast('تم بنجاح إضافة صفوف مجاميع الأصناف الأربعة والإجمالي الكلي بأسفل الجدول!');
    setTimeout(() => setSplitToast(null), 5000);
  };

  // Helper to generate Excel-style column labels (A, B, C ... Z, AA, AB)
  const getColumnLetter = (colIndex: number) => {
    let result = '';
    let index = colIndex;
    while (index >= 0) {
      result = String.fromCharCode((index % 26) + 65) + result;
      index = Math.floor(index / 26) - 1;
    }
    return result;
  };

  // Filtered and sorted rows with persistent original indices
  const displayedRows = useMemo(() => {
    if (!currentSheet) return [];
    let items = currentSheet.rows.map((row, originalIndex) => ({
      row,
      originalIndex,
    }));

    // Category filter (نوع الصنف)
    if (itemTypeColIndex !== -1 && categoryFilter !== 'all') {
      items = items.filter(({ row }) => {
        const val = (row[itemTypeColIndex] ?? '').toString().trim();
        return val === categoryFilter;
      });
    }

    // Search filter
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      items = items.filter(({ row }) =>
        row.some((cell) => (cell ?? '').toString().toLowerCase().includes(q))
      );
    }

    // Sorting
    if (sortConfig !== null) {
      const { columnIndex, direction } = sortConfig;
      items.sort((a, b) => {
        const valA = a.row[columnIndex] ?? '';
        const valB = b.row[columnIndex] ?? '';

        const numA = parseFloat(valA.toString().replace(/,/g, ''));
        const numB = parseFloat(valB.toString().replace(/,/g, ''));

        if (!isNaN(numA) && !isNaN(numB)) {
          return direction === 'asc' ? numA - numB : numB - numA;
        }

        const strA = valA.toString();
        const strB = valB.toString();
        return direction === 'asc'
          ? strA.localeCompare(strB, 'ar')
          : strB.localeCompare(strA, 'ar');
      });
    }

    return items;
  }, [currentSheet, searchQuery, sortConfig, categoryFilter, itemTypeColIndex]);

  // Handle header click for sorting
  const handleSort = (colIndex: number) => {
    if (sortConfig && sortConfig.columnIndex === colIndex) {
      if (sortConfig.direction === 'asc') {
        setSortConfig({ columnIndex: colIndex, direction: 'desc' });
      } else {
        setSortConfig(null);
      }
    } else {
      setSortConfig({ columnIndex: colIndex, direction: 'asc' });
    }
  };

  // Handle cell edit commit
  const handleSaveCell = (
    originalRowIndex: number,
    colIndex: number,
    overrideValue?: string
  ) => {
    if (!currentSheet) return;
    const finalVal = overrideValue !== undefined ? overrideValue : editValue;
    const newSheets = [...document.sheets];
    const newRows = [...currentSheet.rows];
    const newRow = [...newRows[originalRowIndex]];
    newRow[colIndex] = finalVal;

    // If "نوع الصنف" column exists and user edited the formula or item name column, auto update classification
    if (itemTypeColIndex !== -1 && colIndex !== itemTypeColIndex) {
      const formulaCol = findFormulaColumnIndex(currentSheet.headers);
      if (colIndex === formulaCol || formulaCol === -1) {
        newRow[itemTypeColIndex] = classifyFormulaItemType(finalVal);
      }
    }

    newRows[originalRowIndex] = newRow;

    newSheets[activeSheetIndex] = {
      ...currentSheet,
      rows: newRows,
    };

    onUpdateDocument({
      ...document,
      sheets: newSheets,
    });
    setEditingCell(null);
  };

  // Handle header edit commit
  const handleSaveHeader = (colIndex: number) => {
    if (!currentSheet) return;
    const newSheets = [...document.sheets];
    const newHeaders = [...currentSheet.headers];
    newHeaders[colIndex] = editHeaderValue.trim() || `عمود ${colIndex + 1}`;

    newSheets[activeSheetIndex] = {
      ...currentSheet,
      headers: newHeaders,
    };

    onUpdateDocument({
      ...document,
      sheets: newSheets,
    });
    setEditingHeaderIndex(null);
  };

  // Add new empty row
  const handleAddRow = () => {
    if (!currentSheet) return;
    const emptyRow = new Array(currentSheet.headers.length).fill('');
    const newSheets = [...document.sheets];
    newSheets[activeSheetIndex] = {
      ...currentSheet,
      rows: [...currentSheet.rows, emptyRow],
    };
    onUpdateDocument({
      ...document,
      sheets: newSheets,
    });
  };

  // Delete selected row
  const handleDeleteRow = () => {
    if (!currentSheet || selectedRowIndex === null) return;
    const newSheets = [...document.sheets];
    const newRows = currentSheet.rows.filter((_, idx) => idx !== selectedRowIndex);
    newSheets[activeSheetIndex] = {
      ...currentSheet,
      rows: newRows,
    };
    onUpdateDocument({
      ...document,
      sheets: newSheets,
    });
    setSelectedRowIndex(null);
  };

  // Toggle sheet RTL / LTR
  const handleToggleRtl = () => {
    if (!currentSheet) return;
    const newSheets = [...document.sheets];
    newSheets[activeSheetIndex] = {
      ...currentSheet,
      isRtl: !currentSheet.isRtl,
    };
    onUpdateDocument({
      ...document,
      sheets: newSheets,
    });
  };

  // Reorder a column in current sheet (updates headers, rows, totalsRow & sortConfig)
  const handleReorderColumn = (fromIndex: number, toIndex: number) => {
    if (!currentSheet) return;
    const len = currentSheet.headers.length;
    if (
      fromIndex === toIndex ||
      fromIndex < 0 ||
      toIndex < 0 ||
      fromIndex >= len ||
      toIndex >= len
    ) {
      return;
    }

    const moveItem = <T,>(arr: T[]): T[] => {
      const copy = [...arr];
      while (copy.length < len) copy.push('' as unknown as T);
      const [moved] = copy.splice(fromIndex, 1);
      copy.splice(toIndex, 0, moved);
      return copy;
    };

    const newHeaders = moveItem(currentSheet.headers);
    const newRows = currentSheet.rows.map((row) => moveItem(row));
    const newTotalsRow =
      currentSheet.totalsRow && currentSheet.totalsRow.length > 0
        ? moveItem(currentSheet.totalsRow)
        : currentSheet.totalsRow;

    if (sortConfig) {
      let nextSortIdx = sortConfig.columnIndex;
      if (nextSortIdx === fromIndex) {
        nextSortIdx = toIndex;
      } else if (fromIndex < nextSortIdx && toIndex >= nextSortIdx) {
        nextSortIdx -= 1;
      } else if (fromIndex > nextSortIdx && toIndex <= nextSortIdx) {
        nextSortIdx += 1;
      }
      if (nextSortIdx !== sortConfig.columnIndex) {
        setSortConfig({ ...sortConfig, columnIndex: nextSortIdx });
      }
    }

    const newSheets = [...document.sheets];
    newSheets[activeSheetIndex] = {
      ...currentSheet,
      headers: newHeaders,
      rows: newRows,
      totalsRow: newTotalsRow,
    };
    onUpdateDocument({
      ...document,
      sheets: newSheets,
    });
  };

  // Copy to clipboard
  const handleCopyClipboard = async () => {
    if (!currentSheet) return;
    const success = await copyTableToClipboard(currentSheet);
    if (success) {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  // Export full workbook
  const handleDownloadExcel = () => {
    const baseName = fileName.replace(/\.[^/.]+$/, '');
    exportToExcel(document.sheets, `${baseName}.xlsx`, true);
  };

  // Export current sheet CSV
  const handleDownloadCsv = () => {
    if (!currentSheet) return;
    const baseName = fileName.replace(/\.[^/.]+$/, '');
    exportToCsv(
      currentSheet,
      `${baseName}_${currentSheet.sheetName || 'sheet'}.csv`
    );
  };

  // Split code from a column into a separate column
  const handleAutoSplitCodes = (specifiedColIndex?: number) => {
    if (!currentSheet) return;
    const result = autoSplitCodeColumn(currentSheet, specifiedColIndex);
    if (!result) {
      setSplitToast('لم يتم العثور على أكواد إضافية مدمجة داخل هذا العمود.');
      setTimeout(() => setSplitToast(null), 3000);
      return;
    }

    if (result.countSplits === 0 && result.newSheet === currentSheet) {
      setSplitToast('لم يتم العثور على أكواد إضافية مدمجة داخل هذا العمود.');
      setTimeout(() => setSplitToast(null), 3000);
      return;
    }

    const newSheets = [...document.sheets];
    newSheets[activeSheetIndex] = result.newSheet;

    onUpdateDocument({
      ...document,
      sheets: newSheets,
    });

    if (result.countSplits === 0) {
      setSplitToast('تم تنظيف وتوحيد عمود الأكواد وإزالة التكرار بنجاح.');
      setTimeout(() => setSplitToast(null), 4000);
      return;
    }

    setSplitToast(
      `تم بنجاح فصل ${result.countSplits} كود إلى عمود "${result.newColumnName}"!`
    );
    setTimeout(() => setSplitToast(null), 5000);
  };

  // Add / refresh "نوع الصنف" column and classify rows based on user rules
  const handleAutoClassifyItemTypes = () => {
    if (!currentSheet) return;
    const result = ensureItemTypeColumn(currentSheet, document);
    const newSheets = [...document.sheets];
    newSheets[activeSheetIndex] = result.sheet;

    onUpdateDocument({
      ...document,
      sheets: newSheets,
    });

    setSplitToast(
      'تم بنجاح إضافة وتصنيف عمود "نوع الصنف" (تسمين / سمك / بط / اليف) لجميع صفوف الجدول!'
    );
    setTimeout(() => setSplitToast(null), 6000);
  };

  // Ensure formula name is populated in front of every item row
  const handlePropagateFormulaNames = () => {
    if (!currentSheet) return;
    const formulaRes = propagateFormulaNames(currentSheet, document);
    const classifiedRes = ensureItemTypeColumn(formulaRes.sheet, document);
    const newSheets = [...document.sheets];
    newSheets[activeSheetIndex] = classifiedRes.sheet;

    onUpdateDocument({
      ...document,
      sheets: newSheets,
    });

    setSplitToast(
      'تم بنجاح تعبئة وتكرار اسم التركيبة أمام كل صنف وخامة في الجدول وتحديث تصنيفاتها!'
    );
    setTimeout(() => setSplitToast(null), 6000);
  };

  // Merge all sheets into one unified master sheet
  const handleMergeAllSheets = () => {
    if (!document.sheets || document.sheets.length <= 1) return;

    // Use headers with most columns as base
    let masterHeaders = [...document.sheets[0].headers];
    document.sheets.forEach((s) => {
      if (s.headers.length > masterHeaders.length) {
        masterHeaders = [...s.headers];
      }
    });

    const mergedRows: any[][] = [];
    document.sheets.forEach((s) => {
      s.rows.forEach((row) => {
        const aligned = new Array(masterHeaders.length).fill('');
        for (let i = 0; i < Math.min(row.length, masterHeaders.length); i++) {
          aligned[i] = row[i];
        }
        mergedRows.push(aligned);
      });
    });

    const mergedSheet: SheetData = {
      sheetName: 'تقرير_شامل_مدمج',
      description: 'تم دمج كافة التركيبات والأوراق في هذا الجدول الموحد',
      isRtl: document.sheets[0].isRtl,
      headers: masterHeaders,
      rows: mergedRows,
    };

    const formulaRes = propagateFormulaNames(mergedSheet, document);
    const classifiedRes = ensureItemTypeColumn(formulaRes.sheet, document);

    onUpdateDocument({
      ...document,
      sheets: [classifiedRes.sheet, ...document.sheets],
    });
    setActiveSheetIndex(0);

    setSplitToast(`تم بنجاح دمج جميع الأوراق في ورقة شاملة واحدة تضم ${mergedRows.length} صنفاً وتركيبة!`);
    setTimeout(() => setSplitToast(null), 6000);
  };



  if (!currentSheet) {
    return (
      <div className="p-8 text-center bg-white rounded-xl border border-slate-200">
        <p className="text-sm text-slate-500">لا توجد أوراق عمل متاحة.</p>
      </div>
    );
  }

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
      {/* Page Switcher & Direction Switch Bar */}
      <div className="px-4 sm:px-6 pt-3 pb-2 border-b border-slate-200 bg-slate-50/40 flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center flex-wrap gap-2">
          <button
            onClick={() => setActivePage('details')}
            className={`px-3.5 py-1.5 text-xs font-bold rounded-lg border transition-all flex items-center gap-1.5 cursor-pointer ${
              activePage === 'details'
                ? 'bg-emerald-600 text-white border-emerald-600 shadow-xs'
                : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
            }`}
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>جدول البيانات التفصيلي</span>
          </button>

          {totalsSummary.hasItemTypeColumn && (
            <button
              onClick={() => setActivePage('control_pivot')}
              className={`px-3.5 py-1.5 text-xs font-bold rounded-lg border transition-all flex items-center gap-1.5 cursor-pointer ${
                activePage === 'control_pivot'
                  ? 'bg-[#002060] text-amber-300 border-[#002060] shadow-xs'
                  : 'bg-white text-slate-800 border-slate-200 hover:bg-indigo-50 hover:border-indigo-300'
              }`}
            >
              <Scale className="w-3.5 h-3.5" />
              <span>صفحة تجميع الكنترول (حسب الصنف ونوعه)</span>
            </button>
          )}

          {(totalsSummary.hasItemTypeColumn || findFormulaColumnIndex(currentSheet.headers) !== -1) && (
            <button
              onClick={() => setActivePage('sadat_control')}
              className={`px-3.5 py-1.5 text-xs font-bold rounded-lg border transition-all flex items-center gap-1.5 cursor-pointer ${
                activePage === 'sadat_control'
                  ? 'bg-[#ffff00] text-slate-950 border-slate-900 shadow-xs'
                  : 'bg-white text-slate-800 border-slate-200 hover:bg-yellow-50 hover:border-yellow-400'
              }`}
            >
              <TableProperties className="w-3.5 h-3.5" />
              <span>تقرير كنترول السادات (حسب التركيبة)</span>
            </button>
          )}

          <button
            onClick={() => setActivePage('drive_summary')}
            className={`px-3.5 py-1.5 text-xs font-bold rounded-lg border transition-all flex items-center gap-1.5 cursor-pointer ${
              activePage === 'drive_summary'
                ? 'bg-[#ffff00] text-slate-950 border-slate-950 shadow-xs'
                : 'bg-white text-slate-800 border-slate-200 hover:bg-yellow-50 hover:border-yellow-400'
            }`}
          >
            <ListOrdered className="w-3.5 h-3.5" />
            <span>ملخص الأصناف (كود درايف)</span>
          </button>

          <button
            onClick={() => setActivePage('daily_spend')}
            className={`px-3.5 py-1.5 text-xs font-bold rounded-lg border transition-all flex items-center gap-1.5 cursor-pointer ${
              activePage === 'daily_spend'
                ? 'bg-[#ffff00] text-slate-950 border-slate-950 shadow-xs'
                : 'bg-white text-slate-800 border-slate-200 hover:bg-yellow-50 hover:border-yellow-400'
            }`}
          >
            <ListOrdered className="w-3.5 h-3.5" />
            <span>الصرف اليومي (أصناف مختصرة)</span>
          </button>
        </div>

        <div className="flex items-center gap-2.5 shrink-0 flex-wrap">
          {/* Font Size Zoom Control (works across all 3 pages) */}
          <div
            className="inline-flex items-center bg-white border border-slate-200 rounded-lg p-0.5 shadow-2xs"
            title="تكبير أو تصغير حجم الخط في الصفحات"
          >
            <button
              type="button"
              onClick={() => updateFontStep(fontStepIndex - 1)}
              disabled={fontStepIndex === 0}
              className="px-2 py-1 text-xs font-bold text-slate-700 hover:bg-slate-100 rounded-md transition-colors disabled:opacity-35 disabled:cursor-not-allowed cursor-pointer flex items-center gap-1"
              title="تصغير الخط"
            >
              <ZoomOut className="w-3.5 h-3.5 text-slate-600" />
              <span>A-</span>
            </button>

            <button
              type="button"
              onClick={() => updateFontStep(DEFAULT_FONT_STEP_INDEX)}
              className="px-2 py-1 text-xs font-mono font-bold text-slate-800 hover:bg-slate-50 border-x border-slate-200 cursor-pointer tabular-nums"
              title="إرجاع حجم الخط الافتراضي"
            >
              {currentFontStep.px}px
            </button>

            <button
              type="button"
              onClick={() => updateFontStep(fontStepIndex + 1)}
              disabled={fontStepIndex === FONT_SIZE_STEPS.length - 1}
              className="px-2 py-1 text-xs font-bold text-slate-700 hover:bg-slate-100 rounded-md transition-colors disabled:opacity-35 disabled:cursor-not-allowed cursor-pointer flex items-center gap-1"
              title="تكبير الخط"
            >
              <ZoomIn className="w-3.5 h-3.5 text-emerald-700" />
              <span>A+</span>
            </button>
          </div>

          {activePage === 'details' && (
            <button
              onClick={handleToggleRtl}
              className="text-xs text-slate-600 hover:text-emerald-700 font-semibold flex items-center gap-1 shrink-0 cursor-pointer"
              title="تغيير اتجاه الجدول بين اليمين لليسار واليسار لليمين"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>الاتجاه: {currentSheet.isRtl ? 'يمين إلى يسار (RTL)' : 'يسار إلى يمين (LTR)'}</span>
            </button>
          )}

          {onClearData && (
            <button
              type="button"
              onClick={onClearData}
              className="px-3 py-1.5 text-xs font-bold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-lg transition-colors flex items-center gap-1.5 shrink-0 cursor-pointer"
              title="مسح البيانات الحالية والبدء بملف جديد"
            >
              <Trash2 className="w-3.5 h-3.5 text-rose-600" />
              <span>مسح البيانات</span>
            </button>
          )}
        </div>
      </div>

      {activePage === 'control_pivot' && totalsSummary.hasItemTypeColumn ? (
        <ControlPivotView
          sheet={currentSheet}
          fileName={fileName}
          tableFontClass={currentFontStep.tableClass}
        />
      ) : activePage === 'sadat_control' && (totalsSummary.hasItemTypeColumn || findFormulaColumnIndex(currentSheet.headers) !== -1) ? (
        <SadatControlReportView
          sheet={currentSheet}
          fileName={fileName}
          documentTitle={document.documentTitle}
          tableFontClass={currentFontStep.tableClass}
          fontPx={currentFontStep.px}
          canIncreaseFont={fontStepIndex < FONT_SIZE_STEPS.length - 1}
          canDecreaseFont={fontStepIndex > 0}
          onIncreaseFont={() => updateFontStep(fontStepIndex + 1)}
          onDecreaseFont={() => updateFontStep(fontStepIndex - 1)}
          onResetFont={() => updateFontStep(DEFAULT_FONT_STEP_INDEX)}
        />
      ) : activePage === 'daily_spend' ? (
        <DailySpendView
          sheet={currentSheet}
          fileName={fileName}
          documentTitle={document.documentTitle}
        />
      ) : activePage === 'drive_summary' ? (
        <DriveSummaryView
          sheet={currentSheet}
          fileName={fileName}
          tableFontClass={currentFontStep.tableClass}
          fontPx={currentFontStep.px}
          canIncreaseFont={fontStepIndex < FONT_SIZE_STEPS.length - 1}
          canDecreaseFont={fontStepIndex > 0}
          onIncreaseFont={() => updateFontStep(fontStepIndex + 1)}
          onDecreaseFont={() => updateFontStep(fontStepIndex - 1)}
          onResetFont={() => updateFontStep(DEFAULT_FONT_STEP_INDEX)}
        />
      ) : (
      <>
      {/* Prominent Category Totals Dashboard Bar */}
      {totalsSummary.hasItemTypeColumn && (
        <div className="px-4 sm:px-6 pt-2">
          <CategoryTotalsBar
            summary={totalsSummary}
            docTitle={document.documentTitle}
            activeCategoryFilter={categoryFilter}
            onSelectCategoryFilter={(cat) => setCategoryFilter(cat)}
            onAppendTotalsToSheet={handlePermanentlyAppendTotals}
            isTotalsRowAppended={isTotalsRowAppended}
          />
        </div>
      )}

      {/* Main Content: Split View or Full Grid */}
      <div className={viewMode === 'split' ? 'grid grid-cols-1 lg:grid-cols-2 divide-y lg:divide-y-0 lg:divide-x lg:divide-x-reverse divide-slate-200' : ''}>
        {/* PDF Viewer Side (if in split mode) */}
        {viewMode === 'split' && pdfDataUrl && (
          <div className="p-4 bg-slate-100 flex flex-col h-[550px] overflow-hidden">
            <div className="flex items-center justify-between pb-2 mb-2 border-b border-slate-200">
              <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                <FileText className="w-4 h-4 text-emerald-700" />
                <span>المستند الأصلي (للمقارنة والتحقق)</span>
              </span>
              <span className="text-xs text-slate-500">معاينة مباشرة</span>
            </div>
            <div className="flex-1 bg-white rounded-lg border border-slate-200 overflow-hidden">
              <iframe
                src={pdfDataUrl}
                className="w-full h-full border-none"
                title="معاينة ملف PDF"
              />
            </div>
          </div>
        )}

        {/* Interactive Spreadsheet Grid */}
        <div className="flex-1 flex flex-col min-w-0">
          {/* Grid Toolbar */}
          <div className="p-3 sm:px-6 bg-white border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs">
            {/* Search Input */}
            <div className="relative flex-1 min-w-[200px] max-w-sm">
              <Search className="w-4 h-4 text-slate-400 absolute right-2.5 top-2.5" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="بحث وتصفية في الجدول..."
                className="w-full pr-8 pl-3 py-1.5 border border-slate-200 rounded-lg text-xs focus:outline-hidden focus:border-emerald-500"
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery('')}
                  className="absolute left-2.5 top-2.5 text-slate-400 hover:text-slate-600"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Grid Editing & Export Actions */}
            <div className="flex items-center flex-wrap gap-2">
              <button
                onClick={handleAddRow}
                className="px-2.5 py-1.5 text-slate-700 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-md transition-colors flex items-center gap-1 font-semibold cursor-pointer"
                title="إضافة صف فارغ جديد إلى الجدول"
              >
                <Plus className="w-3.5 h-3.5 text-emerald-600" />
                <span>إضافة صف</span>
              </button>

              {selectedRowIndex !== null && (
                <button
                  onClick={handleDeleteRow}
                  className="px-2.5 py-1.5 text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-md transition-colors flex items-center gap-1 font-semibold cursor-pointer"
                  title="حذف الصف المحدد حالياً"
                >
                  <Trash2 className="w-3.5 h-3.5 text-rose-600" />
                  <span>حذف الصف ({selectedRowIndex + 1})</span>
                </button>
              )}

              {pdfDataUrl && (
                <button
                  onClick={() => setViewMode(viewMode === 'grid' ? 'split' : 'grid')}
                  className={`px-2.5 py-1.5 font-semibold rounded-md border transition-colors flex items-center gap-1 cursor-pointer ${
                    viewMode === 'split'
                      ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                      : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                  }`}
                  title="معاينة الملف الأصلي جنباً إلى جنب مع جدول الإكسل"
                >
                  <FileText className="w-3.5 h-3.5" />
                  <span>{viewMode === 'split' ? 'إخفاء المعاينة' : 'معاينة PDF'}</span>
                </button>
              )}

              <button
                onClick={handleCopyClipboard}
                className="px-2.5 py-1.5 font-semibold bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 rounded-md transition-colors flex items-center gap-1 cursor-pointer"
                title="نسخ الجدول"
              >
                {copied ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                    <span className="text-emerald-700">تم النسخ!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5 text-slate-500" />
                    <span>نسخ</span>
                  </>
                )}
              </button>

              <button
                onClick={handleDownloadCsv}
                className="px-2.5 py-1.5 font-semibold bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 rounded-md transition-colors flex items-center gap-1 cursor-pointer"
                title="تحميل CSV"
              >
                <Download className="w-3.5 h-3.5 text-slate-500" />
                <span>CSV</span>
              </button>

              <button
                onClick={() => setShowColumnOrderBar((v) => !v)}
                className={`px-2.5 py-1.5 font-semibold rounded-md border transition-colors flex items-center gap-1 cursor-pointer ${
                  showColumnOrderBar
                    ? 'bg-slate-800 text-white border-slate-800'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                }`}
                title="تغيير ترتيب أعمدة الجدول"
              >
                <Columns3 className="w-3.5 h-3.5" />
                <span>ترتيب الأعمدة</span>
              </button>

              <button
                onClick={handleDownloadExcel}
                className="px-3 py-1.5 font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-md shadow-xs transition-all flex items-center gap-1.5 cursor-pointer"
                title="تحميل مصنف إكسل كامل بصيغة .xlsx"
              >
                <Download className="w-3.5 h-3.5" />
                <span>تحميل Excel (.xlsx)</span>
              </button>
            </div>
          </div>

          {/* Interactive Column Ordering Bar */}
          {showColumnOrderBar && (
            <div className="px-4 sm:px-6 py-2.5 bg-slate-50 border-b border-slate-200 space-y-2">
              <div className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                <Columns3 className="w-3.5 h-3.5 text-emerald-700" />
                <span>اسحب رأس أي عمود في الجدول أو استخدم الأسهم لتغيير ترتيب الأعمدة:</span>
              </div>
              <div className="flex flex-wrap items-center gap-1.5">
                {currentSheet.headers.map((header, colIdx) => (
                  <div
                    key={`${colIdx}_${header}`}
                    draggable
                    onDragStart={() => setDraggedColIndex(colIdx)}
                    onDragOver={(e) => {
                      e.preventDefault();
                      setDragOverColIndex(colIdx);
                    }}
                    onDrop={(e) => {
                      e.preventDefault();
                      if (draggedColIndex !== null) {
                        handleReorderColumn(draggedColIndex, colIdx);
                      }
                      setDraggedColIndex(null);
                      setDragOverColIndex(null);
                    }}
                    onDragEnd={() => {
                      setDraggedColIndex(null);
                      setDragOverColIndex(null);
                    }}
                    className={`inline-flex items-center gap-1 px-2 py-1 rounded-md border text-xs font-bold bg-white shadow-2xs cursor-grab active:cursor-grabbing ${
                      dragOverColIndex === colIdx && draggedColIndex !== colIdx
                        ? 'border-emerald-500 ring-2 ring-emerald-200'
                        : 'border-slate-200 text-slate-800'
                    }`}
                  >
                    <GripVertical className="w-3 h-3 text-slate-400" />
                    <span className="text-[10px] font-mono text-slate-400">
                      {getColumnLetter(colIdx)}
                    </span>
                    <span className="max-w-[140px] truncate">{header}</span>
                    <div className="flex items-center border-r border-slate-200 pr-1 mr-0.5 gap-0.5">
                      <button
                        onClick={() =>
                          handleReorderColumn(colIdx, currentSheet.isRtl ? colIdx - 1 : colIdx + 1)
                        }
                        disabled={
                          currentSheet.isRtl
                            ? colIdx === 0
                            : colIdx === currentSheet.headers.length - 1
                        }
                        className="p-0.5 rounded hover:bg-slate-100 disabled:opacity-25 disabled:cursor-not-allowed cursor-pointer"
                        title="تحريك لليمين"
                      >
                        <ChevronRight className="w-3.5 h-3.5 text-slate-700" />
                      </button>
                      <button
                        onClick={() =>
                          handleReorderColumn(colIdx, currentSheet.isRtl ? colIdx + 1 : colIdx - 1)
                        }
                        disabled={
                          currentSheet.isRtl
                            ? colIdx === currentSheet.headers.length - 1
                            : colIdx === 0
                        }
                        className="p-0.5 rounded hover:bg-slate-100 disabled:opacity-25 disabled:cursor-not-allowed cursor-pointer"
                        title="تحريك لليسار"
                      >
                        <ChevronLeft className="w-3.5 h-3.5 text-slate-700" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Split notification banner if triggered */}
          {splitToast && (
            <div className="px-4 py-2.5 bg-emerald-50 border-b border-emerald-200 text-emerald-900 text-xs font-semibold flex items-center justify-between animate-in fade-in duration-150">
              <div className="flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                <span>{splitToast}</span>
              </div>
              <button
                onClick={() => setSplitToast(null)}
                className="text-emerald-700 hover:text-emerald-900 px-1 font-bold"
              >
                ✕
              </button>
            </div>
          )}

          {/* Table Container */}
          <div
            className={`overflow-x-auto max-h-[560px] ${
              currentSheet.isRtl ? 'direction-rtl' : 'direction-ltr'
            }`}
            dir={currentSheet.isRtl ? 'rtl' : 'ltr'}
          >
            <table className={`w-full border-collapse text-right ${currentFontStep.tableClass}`}>
              {/* Excel Column Letters Bar (A, B, C...) */}
              <thead className="sticky top-0 z-20 bg-slate-100 text-slate-500 border-b border-slate-300">
                <tr className="font-mono text-[11px]">
                  <th className="w-12 px-2 py-1.5 text-center bg-slate-200 border-r border-slate-300 font-semibold select-none">
                    #
                  </th>
                  {currentSheet.headers.map((_, colIdx) => (
                    <th
                      key={colIdx}
                      className="px-3 py-1.5 text-center border-r border-slate-200 select-none uppercase"
                    >
                      {getColumnLetter(colIdx)}
                    </th>
                  ))}
                </tr>

                {/* Actual Table Header Names */}
                <tr className="bg-slate-50 text-slate-800 font-bold border-b border-slate-300">
                  <th className="w-12 px-2 py-2.5 text-center bg-slate-100 border-r border-slate-300 select-none">
                    0
                  </th>
                  {currentSheet.headers.map((header, colIdx) => {
                    const isSorted = sortConfig?.columnIndex === colIdx;
                    const isDragOver = dragOverColIndex === colIdx && draggedColIndex !== colIdx;
                    return (
                      <th
                        key={colIdx}
                        draggable={editingHeaderIndex !== colIdx}
                        onDragStart={() => setDraggedColIndex(colIdx)}
                        onDragOver={(e) => {
                          e.preventDefault();
                          setDragOverColIndex(colIdx);
                        }}
                        onDragLeave={() => {
                          if (dragOverColIndex === colIdx) setDragOverColIndex(null);
                        }}
                        onDrop={(e) => {
                          e.preventDefault();
                          if (draggedColIndex !== null) {
                            handleReorderColumn(draggedColIndex, colIdx);
                          }
                          setDraggedColIndex(null);
                          setDragOverColIndex(null);
                        }}
                        onDragEnd={() => {
                          setDraggedColIndex(null);
                          setDragOverColIndex(null);
                        }}
                        className={`px-3 py-2.5 border-r border-slate-200 hover:bg-slate-100 transition-colors cursor-pointer select-none group ${
                          isDragOver ? 'ring-2 ring-inset ring-emerald-500 bg-emerald-50' : ''
                        }`}
                        onClick={() => handleSort(colIdx)}
                      >
                        <div className="flex items-center justify-between gap-2">
                          <div className="flex items-center gap-1 min-w-0">
                            <span title="اسحب لتغيير ترتيب العمود" className="inline-flex shrink-0">
                              <GripVertical className="w-3 h-3 text-slate-300 group-hover:text-slate-500 cursor-grab" />
                            </span>
                            {editingHeaderIndex === colIdx ? (
                              <input
                                type="text"
                                value={editHeaderValue}
                                onChange={(e) => setEditHeaderValue(e.target.value)}
                                onBlur={() => handleSaveHeader(colIdx)}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') handleSaveHeader(colIdx);
                                  if (e.key === 'Escape') setEditingHeaderIndex(null);
                                }}
                                autoFocus
                                className="px-1.5 py-0.5 border border-emerald-500 rounded text-xs bg-white text-slate-900 w-full"
                                onClick={(e) => e.stopPropagation()}
                              />
                            ) : (
                              <span
                                className="font-bold truncate text-slate-900"
                                onDoubleClick={(e) => {
                                  e.stopPropagation();
                                  setEditingHeaderIndex(colIdx);
                                  setEditHeaderValue(header);
                                }}
                                title="انقر مرتين لتعديل اسم العمود، أو اسحب لتغيير مكانه"
                              >
                                {header}
                              </span>
                            )}
                          </div>

                          <div className="flex items-center gap-1 shrink-0">
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                handleAutoSplitCodes(colIdx);
                              }}
                              className="opacity-0 group-hover:opacity-100 p-0.5 text-slate-400 hover:text-amber-600 rounded transition-opacity"
                              title="فصل الأكواد من هذا العمود إلى عمود مستقل"
                            >
                              <Scissors className="w-3 h-3" />
                            </button>

                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setEditingHeaderIndex(colIdx);
                                setEditHeaderValue(header);
                              }}
                              className="opacity-0 group-hover:opacity-100 p-0.5 text-slate-400 hover:text-slate-600 rounded transition-opacity"
                              title="تعديل اسم العمود"
                            >
                              <Edit2 className="w-3 h-3" />
                            </button>

                            {isSorted ? (
                              sortConfig?.direction === 'asc' ? (
                                <ArrowUp className="w-3.5 h-3.5 text-emerald-600" />
                              ) : (
                                <ArrowDown className="w-3.5 h-3.5 text-emerald-600" />
                              )
                            ) : (
                              <ArrowUpDown className="w-3 h-3 text-slate-300 opacity-0 group-hover:opacity-100 transition-opacity" />
                            )}
                          </div>
                        </div>
                      </th>
                    );
                  })}
                </tr>
              </thead>

              {/* Data Rows */}
              <tbody className="divide-y divide-slate-100 bg-white">
                {displayedRows.length === 0 ? (
                  <tr>
                    <td
                      colSpan={currentSheet.headers.length + 1}
                      className="px-6 py-8 text-center text-slate-400 text-xs"
                    >
                      {searchQuery
                        ? 'لم يتم العثور على نتائج مطابقة لكلمة البحث.'
                        : 'لا توجد بيانات في هذه الورقة.'}
                    </td>
                  </tr>
                ) : (
                  displayedRows.map(({ row, originalIndex }, displayIdx) => {
                    const isSelected = selectedRowIndex === originalIndex;
                    const isSubtotalRow = row.some(
                      (cell) =>
                        typeof cell === 'string' &&
                        (cell.includes('إجمالي السمك') ||
                          cell.includes('إجمالي التسمين') ||
                          cell.includes('إجمالي البط') ||
                          cell.includes('إجمالي الأليف') ||
                          cell.includes('الإجمالي الكلي للأربعة'))
                    );
                    const isGrandTotalRow = row.some(
                      (cell) =>
                        typeof cell === 'string' && cell.includes('الإجمالي الكلي للأربعة')
                    );

                    return (
                      <tr
                        key={originalIndex}
                        onClick={() => setSelectedRowIndex(originalIndex)}
                        className={`transition-colors cursor-pointer ${
                          isGrandTotalRow
                            ? 'bg-emerald-100/90 font-extrabold text-emerald-950 border-y-2 border-emerald-400'
                            : isSubtotalRow
                            ? 'bg-amber-50/90 font-bold text-slate-900 border-y border-amber-300'
                            : isSelected
                            ? 'bg-emerald-50/80'
                            : displayIdx % 2 === 0
                            ? 'bg-white hover:bg-slate-50'
                            : 'bg-slate-50/40 hover:bg-slate-100/50'
                        }`}
                      >
                        {/* Row Index Number */}
                        <td className="w-12 px-2 py-2 text-center text-[11px] font-mono font-bold tabular-nums text-slate-600 bg-slate-50/80 border-r border-slate-200 select-none">
                          {isGrandTotalRow ? '⭐' : isSubtotalRow ? 'Σ' : displayIdx + 1}
                        </td>

                        {/* Cell Contents */}
                        {currentSheet.headers.map((_, colIdx) => {
                          const cellValue = row[colIdx] ?? '';
                          const isEditing =
                            editingCell?.rowIndex === originalIndex &&
                            editingCell?.colIndex === colIdx;
                          const isItemType = colIdx === itemTypeColIndex;

                          // Check if looks numeric for tabular-nums font
                          const isNumeric =
                            !isItemType &&
                            (typeof cellValue === 'number' ||
                              (!isNaN(Number(cellValue)) && cellValue !== ''));

                          const valStr = (cellValue ?? '').toString().trim();

                          return (
                            <td
                              key={colIdx}
                              onDoubleClick={() => {
                                setEditingCell({ rowIndex: originalIndex, colIndex: colIdx });
                                setEditValue(valStr);
                              }}
                              className={`px-3 py-2 border-r border-slate-100 whitespace-nowrap font-bold text-slate-900 ${
                                isNumeric ? 'font-mono tabular-nums' : ''
                              }`}
                            >
                              {isEditing ? (
                                isItemType ? (
                                  <select
                                    value={editValue}
                                    onChange={(e) => {
                                      const newVal = e.target.value;
                                      setEditValue(newVal);
                                      handleSaveCell(originalIndex, colIdx, newVal);
                                    }}
                                    onBlur={() => handleSaveCell(originalIndex, colIdx)}
                                    autoFocus
                                    className="px-2 py-1 border-2 border-indigo-500 rounded text-xs bg-white text-slate-900 w-full focus:outline-hidden font-bold"
                                  >
                                    <option value="تسمين">🌾 تسمين</option>
                                    <option value="سمك">🐟 سمك</option>
                                    <option value="بط">🦆 بط</option>
                                    <option value="اليف">🐾 اليف</option>
                                    <option value="تنظيف">🧹 تنظيف (Cleaning)</option>
                                  </select>
                                ) : (
                                  <input
                                    type="text"
                                    value={editValue}
                                    onChange={(e) => setEditValue(e.target.value)}
                                    onBlur={() => handleSaveCell(originalIndex, colIdx)}
                                    onKeyDown={(e) => {
                                      if (e.key === 'Enter')
                                        handleSaveCell(originalIndex, colIdx);
                                      if (e.key === 'Escape') setEditingCell(null);
                                    }}
                                    autoFocus
                                    className="px-2 py-1 border border-emerald-500 rounded text-xs bg-white text-slate-900 w-full focus:outline-hidden"
                                  />
                                )
                              ) : isItemType ? (
                                valStr === 'سمك' ? (
                                  <span
                                    title="نوع الصنف: سمك (Tilapia) - انقر مرتين للتغيير"
                                    className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-bold bg-sky-100 text-sky-800 border border-sky-300 shadow-2xs"
                                  >
                                    <span>🐟</span>
                                    <span>سمك</span>
                                  </span>
                                ) : valStr === 'تسمين' ? (
                                  <span
                                    title="نوع الصنف: تسمين (GR, FN, ST, SS) - انقر مرتين للتغيير"
                                    className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-bold bg-amber-100 text-amber-900 border border-amber-300 shadow-2xs"
                                  >
                                    <span>🌾</span>
                                    <span>تسمين</span>
                                  </span>
                                ) : valStr === 'بط' ? (
                                  <span
                                    title="نوع الصنف: بط (Duck) - انقر مرتين للتغيير"
                                    className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-bold bg-teal-100 text-teal-800 border border-teal-300 shadow-2xs"
                                  >
                                    <span>🦆</span>
                                    <span>بط</span>
                                  </span>
                                ) : (
                                  <span
                                    title="نوع الصنف: اليف (أصناف أخرى) - انقر مرتين للتغيير"
                                    className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-bold bg-purple-100 text-purple-800 border border-purple-300 shadow-2xs"
                                  >
                                    <span>🐾</span>
                                    <span>اليف</span>
                                  </span>
                                )
                              ) : (
                                <span title="انقر مرتين لتعديل الخلية">
                                  {cellValue === '' ? (
                                    <span className="text-slate-300">-</span>
                                  ) : (
                                    cellValue.toString()
                                  )}
                                </span>
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })
                )}

                {/* Built-in Category Subtotal Rows for the 4 categories (سمك، تسمين، بط، أليف) */}
                {totalsSummary.hasItemTypeColumn && !isTotalsRowAppended && (
                  <>
                    {(['سمك', 'تسمين', 'بط', 'اليف', 'تنظيف'] as const).map((catKey) => {
                      const cat = totalsSummary.categories[catKey];
                      const targetIdx = totalsSummary.targetColIndex;
                      const actualIdx = totalsSummary.actualColIndex;
                      const formulaIdx = findFormulaColumnIndex(currentSheet.headers);
                      const fallbackLabelIdx = currentSheet.headers.findIndex((h) => {
                        const norm = (h ?? '').toString().trim().toLowerCase();
                        return Boolean(
                          norm &&
                          !norm.includes('code') &&
                          !norm.includes('كود') &&
                          !norm.includes('رمز') &&
                          !norm.includes('رقم') &&
                          !norm.includes('supledger') &&
                          !norm.includes('subledger') &&
                          !norm.includes('نوع الصنف') &&
                          !norm.includes('target') &&
                          !norm.includes('actual')
                        );
                      });
                      const nameCol =
                        formulaIdx !== -1
                          ? formulaIdx
                          : fallbackLabelIdx !== -1
                          ? fallbackLabelIdx
                          : 0;

                      const rowBg =
                        catKey === 'سمك'
                          ? 'bg-sky-50/80 text-sky-950 border-t border-sky-200'
                          : catKey === 'تسمين'
                          ? 'bg-amber-50/80 text-amber-950 border-t border-amber-200'
                          : catKey === 'بط'
                          ? 'bg-teal-50/80 text-teal-950 border-t border-teal-200'
                          : 'bg-purple-50/80 text-purple-950 border-t border-purple-200';

                      return (
                        <tr key={catKey} className={`font-semibold ${rowBg}`}>
                          <td className="w-12 px-2 py-2 text-center text-xs font-mono border-r border-slate-200 select-none">
                            {cat.emoji}
                          </td>
                          {currentSheet.headers.map((_, colIdx) => {
                            if (colIdx === nameCol) {
                              return (
                                <td
                                  key={colIdx}
                                  className="px-3 py-2 border-r border-slate-200 font-bold"
                                >
                                  إجمالي {catKey} ({cat.count} صنف)
                                </td>
                              );
                            }
                            if (colIdx === itemTypeColIndex) {
                              return (
                                <td
                                  key={colIdx}
                                  className="px-3 py-2 border-r border-slate-200 font-bold text-xs"
                                >
                                  {catKey}
                                </td>
                              );
                            }
                            if (colIdx === targetIdx && cat.hasTargetData) {
                              return (
                                <td
                                  key={colIdx}
                                  className="px-3 py-2 border-r border-slate-200 font-mono font-bold tabular-nums"
                                >
                                  {cat.targetSum.toFixed(3)}
                                </td>
                              );
                            }
                            if (colIdx === actualIdx && cat.hasActualData) {
                              return (
                                <td
                                  key={colIdx}
                                  className="px-3 py-2 border-r border-slate-200 font-mono font-bold tabular-nums"
                                >
                                  {cat.actualSum.toFixed(3)}
                                </td>
                              );
                            }
                            return (
                              <td
                                key={colIdx}
                                className="px-3 py-2 border-r border-slate-200 text-slate-400"
                              >
                                -
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}

                    {/* Grand Total Row for the 4 Categories */}
                    <tr className="bg-gradient-to-r from-emerald-100 via-emerald-50 to-teal-100 font-extrabold text-emerald-950 border-t-2 border-b-2 border-emerald-400">
                      <td className="w-12 px-2 py-2.5 text-center text-xs font-mono bg-emerald-200 border-r border-emerald-300 select-none">
                        ⭐
                      </td>
                      {currentSheet.headers.map((_, colIdx) => {
                        const formulaIdx = findFormulaColumnIndex(currentSheet.headers);
                        const fallbackLabelIdx = currentSheet.headers.findIndex((h) => {
                          const norm = (h ?? '').toString().trim().toLowerCase();
                          return Boolean(
                            norm &&
                            !norm.includes('code') &&
                            !norm.includes('كود') &&
                            !norm.includes('رمز') &&
                            !norm.includes('رقم') &&
                            !norm.includes('supledger') &&
                            !norm.includes('subledger') &&
                            !norm.includes('نوع الصنف') &&
                            !norm.includes('target') &&
                            !norm.includes('actual')
                          );
                        });
                        const nameCol =
                          formulaIdx !== -1
                            ? formulaIdx
                            : fallbackLabelIdx !== -1
                            ? fallbackLabelIdx
                            : 0;
                        const targetIdx = totalsSummary.targetColIndex;
                        const actualIdx = totalsSummary.actualColIndex;

                        if (colIdx === nameCol) {
                          return (
                            <td
                              key={colIdx}
                              className="px-3 py-2.5 border-r border-emerald-300 font-extrabold text-emerald-950"
                            >
                              ⭐ الإجمالي الكلي للأربعة (Grand Total)
                            </td>
                          );
                        }
                        if (colIdx === itemTypeColIndex) {
                          return (
                            <td
                              key={colIdx}
                              className="px-3 py-2.5 border-r border-emerald-300 font-extrabold text-emerald-900"
                            >
                              الكل ({totalsSummary.grandTotal.count})
                            </td>
                          );
                        }
                        if (colIdx === targetIdx && totalsSummary.grandTotal.targetSum > 0) {
                          return (
                            <td
                              key={colIdx}
                              className="px-3 py-2.5 border-r border-emerald-300 font-mono font-extrabold tabular-nums text-emerald-950"
                            >
                              {totalsSummary.grandTotal.targetSum.toFixed(3)}
                            </td>
                          );
                        }
                        if (colIdx === actualIdx && totalsSummary.grandTotal.actualSum > 0) {
                          return (
                            <td
                              key={colIdx}
                              className="px-3 py-2.5 border-r border-emerald-300 font-mono font-extrabold tabular-nums text-emerald-950"
                            >
                              {totalsSummary.grandTotal.actualSum.toFixed(3)}
                            </td>
                          );
                        }
                        return (
                          <td
                            key={colIdx}
                            className="px-3 py-2.5 border-r border-emerald-300 text-slate-400"
                          >
                            -
                          </td>
                        );
                      })}
                    </tr>
                  </>
                )}

                {/* Legacy Sheet Totals Row if present and not overlapping */}
                {currentSheet.totalsRow && currentSheet.totalsRow.length > 0 && !totalsSummary.hasItemTypeColumn && (
                  <tr className="bg-slate-100/90 font-bold text-slate-900 border-t-2 border-slate-300">
                    <td className="w-12 px-2 py-2.5 text-center text-[11px] font-mono text-emerald-800 bg-slate-200 border-r border-slate-300 select-none">
                      Σ
                    </td>
                    {currentSheet.headers.map((_, colIdx) => {
                      const totalVal = currentSheet.totalsRow?.[colIdx] ?? '';
                      return (
                        <td
                          key={colIdx}
                          className="px-3 py-2.5 border-r border-slate-200 font-mono tabular-nums text-emerald-950 font-bold"
                        >
                          {totalVal}
                        </td>
                      );
                    })}
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Grid Footer with tips */}
          <div className="p-3 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between text-xs text-slate-500">
            <div className="flex items-center gap-3">
              <span>💡 انقر نقراً مزدوجاً على أي خلية لتعديل محتواها مباشرة</span>
              <span aria-hidden="true">·</span>
              <span>انقر على اسم العمود لترتيب البيانات</span>
            </div>
            <div>
              <span>
                إجمالي الصفوف:{' '}
                <strong className="text-slate-800 font-mono tabular-nums">
                  {currentSheet.rows.length}
                </strong>
              </span>
            </div>
          </div>
        </div>
      </div>
      </>
      )}
    </div>
  );
};
