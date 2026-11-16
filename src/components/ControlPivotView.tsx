import React, { useMemo, useRef, useState } from 'react';
import {
  Download,
  Copy,
  Check,
  Search,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Scale,
  FileSpreadsheet,
  Upload,
  CheckCircle2,
  GripVertical,
  Columns3,
  ChevronRight,
  ChevronLeft,
  RotateCcw,
} from 'lucide-react';
import { SheetData } from '../types/converter';
import {
  calculateControlPivot,
  buildControlPivotSheetData,
  ControlPivotRow,
  ControlPivotColumnKey,
  PivotTotalsGroup,
  PIVOT_COLUMN_DEFS,
  DEFAULT_PIVOT_COLUMN_ORDER,
  PIVOT_COLUMN_ORDER_STORAGE_KEY,
  loadStoredPivotColumnOrder,
  loadStoredDriveOverrides,
  loadStoredCustomCatalog,
  isWaterPivotRow,
  sumPivotRows,
} from '../utils/controlPivotCalculator';
import {
  DriveCodeEntry,
  DRIVE_OVERRIDES_STORAGE_KEY,
  DRIVE_CUSTOM_CATALOG_STORAGE_KEY,
  normalizeItemText,
  parseDriveCodesFile,
} from '../utils/driveCodesCatalog';
import {
  exportToCsv,
  copyTableToClipboard,
} from '../utils/excelExport';
import { exportControlPivotStyledExcel } from '../utils/styledExcelExport';

interface ControlPivotViewProps {
  sheet: SheetData;
  fileName?: string;
  tableFontClass?: string;
}

type SortField = ControlPivotColumnKey;

export const ControlPivotView: React.FC<ControlPivotViewProps> = ({
  sheet,
  fileName = 'control_summary.xlsx',
  tableFontClass = 'text-xs sm:text-sm',
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [metric, setMetric] = useState<'actual' | 'target'>('actual');

  const baseIsKg = useMemo(() => {
    const joined = (sheet.headers || []).join(' ').toLowerCase();
    return joined.includes('kg') || joined.includes('كجم') || (!joined.includes('ton') && !joined.includes('طن'));
  }, [sheet.headers]);

  const [useTons, setUseTons] = useState<boolean>(baseIsKg);
  // Default sort: الاجمالى من الأكبر فوق إلى الأصغر (descending)
  const [sortConfig, setSortConfig] = useState<{
    field: SortField;
    direction: 'asc' | 'desc';
  }>({ field: 'total', direction: 'desc' });
  const [copied, setCopied] = useState(false);

  // Column ordering state (drag & drop + quick arrow buttons, persisted in localStorage)
  const [columnOrder, setColumnOrder] = useState<ControlPivotColumnKey[]>(() =>
    loadStoredPivotColumnOrder()
  );
  const [showColumnOrderBar, setShowColumnOrderBar] = useState(false);
  const [draggedColKey, setDraggedColKey] = useState<ControlPivotColumnKey | null>(null);
  const [dragOverColKey, setDragOverColKey] = useState<ControlPivotColumnKey | null>(null);

  // Persistent Drive code overrides & imported custom catalog
  const [overrides, setOverrides] = useState<Record<string, string>>(() => loadStoredDriveOverrides());
  const [customCatalog, setCustomCatalog] = useState<DriveCodeEntry[]>(() => loadStoredCustomCatalog());
  const [editingRowKey, setEditingRowKey] = useState<string | null>(null);
  const [editingDriveValue, setEditingDriveValue] = useState('');
  const [importToast, setImportToast] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const divisor = useTons && baseIsKg ? 1000 : 1;
  const decimals = useTons && baseIsKg ? 6 : 3;

  const pivot = useMemo(
    () => calculateControlPivot(sheet, metric, divisor, overrides, customCatalog),
    [sheet, metric, divisor, overrides, customCatalog]
  );

  const matchedCount = useMemo(
    () => pivot.rows.filter((r) => Boolean(r.driveCode)).length,
    [pivot.rows]
  );

  const displayedRows = useMemo(() => {
    let items = [...pivot.rows];

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      items = items.filter(
        (r) =>
          r.name.toLowerCase().includes(q) ||
          r.code.toLowerCase().includes(q) ||
          r.driveCode.toLowerCase().includes(q)
      );
    }

    if (sortConfig) {
      items.sort((a, b) => {
        const valA = a[sortConfig.field];
        const valB = b[sortConfig.field];
        if (sortConfig.field === 'driveCode' || sortConfig.field === 'code') {
          const numA = parseFloat(String(valA));
          const numB = parseFloat(String(valB));
          if (!isNaN(numA) && !isNaN(numB)) {
            return sortConfig.direction === 'asc' ? numA - numB : numB - numA;
          }
        }
        if (typeof valA === 'number' && typeof valB === 'number') {
          return sortConfig.direction === 'asc' ? valA - valB : valB - valA;
        }
        return sortConfig.direction === 'asc'
          ? String(valA).localeCompare(String(valB), 'ar')
          : String(valB).localeCompare(String(valA), 'ar');
      });
    }

    return items;
  }, [pivot.rows, searchQuery, sortConfig]);

  const regularRows = useMemo(
    () => displayedRows.filter((r) => !isWaterPivotRow(r)),
    [displayedRows]
  );

  const waterRows = useMemo(
    () => displayedRows.filter((r) => isWaterPivotRow(r)),
    [displayedRows]
  );

  const regularMatchedCount = useMemo(
    () => regularRows.filter((r) => Boolean(r.driveCode)).length,
    [regularRows]
  );

  const filteredMatchedCount = useMemo(
    () => displayedRows.filter((r) => Boolean(r.driveCode)).length,
    [displayedRows]
  );

  const totalsWithoutWater = useMemo(() => sumPivotRows(regularRows), [regularRows]);
  const totalsWithWater = useMemo(() => sumPivotRows(displayedRows), [displayedRows]);

  const persistColumnOrder = (nextOrder: ControlPivotColumnKey[]) => {
    setColumnOrder(nextOrder);
    try {
      localStorage.setItem(PIVOT_COLUMN_ORDER_STORAGE_KEY, JSON.stringify(nextOrder));
    } catch (err) {
      console.error('Failed to save pivot column order', err);
    }
  };

  const moveColumnByOffset = (colKey: ControlPivotColumnKey, delta: -1 | 1) => {
    const idx = columnOrder.indexOf(colKey);
    const targetIdx = idx + delta;
    if (idx === -1 || targetIdx < 0 || targetIdx >= columnOrder.length) return;
    const next = [...columnOrder];
    const [removed] = next.splice(idx, 1);
    next.splice(targetIdx, 0, removed);
    persistColumnOrder(next);
  };

  const reorderColumnToTarget = (fromKey: ControlPivotColumnKey, toKey: ControlPivotColumnKey) => {
    if (fromKey === toKey) return;
    const fromIdx = columnOrder.indexOf(fromKey);
    const toIdx = columnOrder.indexOf(toKey);
    if (fromIdx === -1 || toIdx === -1) return;
    const next = [...columnOrder];
    const [removed] = next.splice(fromIdx, 1);
    next.splice(toIdx, 0, removed);
    persistColumnOrder(next);
  };

  const handleResetColumnOrder = () => {
    persistColumnOrder([...DEFAULT_PIVOT_COLUMN_ORDER]);
  };

  const isDefaultColumnOrder = useMemo(
    () => columnOrder.every((k, idx) => DEFAULT_PIVOT_COLUMN_ORDER[idx] === k),
    [columnOrder]
  );

  const handleSort = (field: SortField) => {
    setSortConfig((prev) => {
      if (prev.field === field) {
        return {
          field,
          direction: prev.direction === 'desc' ? 'asc' : 'desc',
        };
      }
      return {
        field,
        direction: field === 'name' || field === 'code' || field === 'driveCode' ? 'asc' : 'desc',
      };
    });
  };

  const renderSortIcon = (field: SortField) => {
    if (sortConfig.field !== field) {
      return <ArrowUpDown className="w-3 h-3 opacity-50" />;
    }
    return sortConfig.direction === 'asc' ? (
      <ArrowUp className="w-3 h-3 text-amber-300" />
    ) : (
      <ArrowDown className="w-3 h-3 text-amber-300" />
    );
  };

  const formatCell = (val: number) => {
    if (Math.abs(val) < 1e-9) return '-';
    return val.toFixed(decimals);
  };

  const getRowKey = (row: ControlPivotRow) => `${row.code}::${normalizeItemText(row.name)}`;

  const handleStartEditDriveCode = (row: ControlPivotRow) => {
    setEditingRowKey(getRowKey(row));
    setEditingDriveValue(row.driveCode || '');
  };

  const handleSaveDriveCode = (row: ControlPivotRow) => {
    const key = getRowKey(row);
    const updated = { ...overrides, [key]: editingDriveValue.trim() };
    setOverrides(updated);
    try {
      localStorage.setItem(DRIVE_OVERRIDES_STORAGE_KEY, JSON.stringify(updated));
    } catch (err) {
      console.error('Failed to save drive code override', err);
    }
    setEditingRowKey(null);
  };

  const handleImportCatalogFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const entries = await parseDriveCodesFile(file);
      if (entries.length === 0) {
        setImportToast('لم يتم العثور على أكواد صالحة في الملف. تأكد من وجود أعمدة: اسم الصنف، كود الصنف دريف.');
      } else {
        const merged = [...entries, ...customCatalog];
        setCustomCatalog(merged);
        try {
          localStorage.setItem(DRIVE_CUSTOM_CATALOG_STORAGE_KEY, JSON.stringify(merged.slice(0, 1500)));
        } catch (err) {
          console.error('Failed to persist custom catalog', err);
        }
        setImportToast(`تم استيراد ${entries.length} صنفاً وتحديث أكواد الدريف تلقائياً!`);
      }
      setTimeout(() => setImportToast(null), 5000);
    } catch (err) {
      console.error('Error importing drive codes sheet', err);
      setImportToast('تعذر قراءة ملف الأكواد. يرجى رفع ملف Excel (.xlsx) أو CSV صحيح.');
      setTimeout(() => setImportToast(null), 4000);
    } finally {
      e.target.value = '';
    }
  };

  const pivotSheetData = useMemo(
    () => buildControlPivotSheetData(pivot, decimals, columnOrder, displayedRows),
    [pivot, decimals, columnOrder, displayedRows]
  );

  const handleDownloadPivotExcel = async () => {
    await exportControlPivotStyledExcel({
      pivot,
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
    });
  };

  const handleDownloadPivotCsv = () => {
    const baseName = fileName.replace(/\.(xlsx|pdf|csv)$/i, '');
    exportToCsv(pivotSheetData, `${baseName}_تجميع_الكنترول.csv`);
  };

  const handleCopyPivot = async () => {
    const ok = await copyTableToClipboard(pivotSheetData);
    if (ok) {
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    }
  };

  const displayUnitLabel = useTons && baseIsKg ? 'طن' : pivot.unit;

  const renderDataCell = (
    colKey: ControlPivotColumnKey,
    row: ControlPivotRow,
    isLastCol: boolean,
    isWaterHighlight: boolean = false
  ) => {
    const borderClass = isLastCol ? '' : 'border-l border-slate-300';
    const rowKey = getRowKey(row);
    const isEditingDrive = editingRowKey === rowKey;
    const cellBg = isWaterHighlight ? 'bg-sky-50/80' : 'bg-white';

    switch (colKey) {
      case 'driveCode':
        return (
          <td
            key={colKey}
            onClick={() => !isEditingDrive && handleStartEditDriveCode(row)}
            className={`px-2 py-1.5 ${borderClass} font-mono text-center font-extrabold text-indigo-900 tabular-nums whitespace-nowrap ${
              isWaterHighlight ? 'bg-sky-100/70' : 'bg-indigo-50/40'
            } cursor-pointer hover:bg-indigo-100/60 transition-colors`}
            title="انقر لتعديل أو إدخال كود الدريف يدوياً"
          >
            {isEditingDrive ? (
              <input
                type="text"
                autoFocus
                value={editingDriveValue}
                onChange={(e) => setEditingDriveValue(e.target.value)}
                onBlur={() => handleSaveDriveCode(row)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleSaveDriveCode(row);
                  if (e.key === 'Escape') setEditingRowKey(null);
                }}
                onClick={(e) => e.stopPropagation()}
                className="w-20 px-1.5 py-0.5 text-center font-mono text-xs border border-indigo-500 rounded bg-white text-slate-900 focus:outline-hidden"
              />
            ) : row.driveCode ? (
              row.driveCode
            ) : (
              <span className="text-slate-300 font-normal text-xs">+ كود</span>
            )}
          </td>
        );

      case 'code':
        return (
          <td
            key={colKey}
            className={`px-3 py-1.5 ${borderClass} font-mono text-center font-bold text-slate-900 tabular-nums whitespace-nowrap ${cellBg}`}
          >
            {row.code || ''}
          </td>
        );

      case 'name':
        return (
          <td
            key={colKey}
            className={`px-4 py-1.5 ${borderClass} font-bold ${
              isWaterHighlight ? 'text-sky-950' : 'text-slate-950'
            } ${cellBg}`}
          >
            {row.name}
          </td>
        );

      case 'fattening':
      case 'fish':
      case 'duck':
      case 'pets':
        return (
          <td
            key={colKey}
            className={`px-3 py-1.5 ${borderClass} font-mono tabular-nums text-left font-bold text-slate-950 ${cellBg}`}
            dir="ltr"
          >
            {formatCell(row[colKey])}
          </td>
        );

      case 'total':
        return (
          <td
            key={colKey}
            className={`px-4 py-1.5 ${borderClass} font-mono tabular-nums text-left font-extrabold text-slate-950 ${
              isWaterHighlight ? 'bg-sky-100/90' : 'bg-[#fff2cc]'
            }`}
            dir="ltr"
          >
            {row.total.toFixed(decimals)}
          </td>
        );
    }
  };

  const renderFooterCell = (
    colKey: ControlPivotColumnKey,
    isLastCol: boolean,
    totalsGroup: PivotTotalsGroup,
    label: string,
    matchedCnt: number,
    totalCnt: number,
    totalHighlightClass: string
  ) => {
    const borderClass = isLastCol ? '' : 'border-l border-slate-400';

    switch (colKey) {
      case 'driveCode':
        return (
          <td key={colKey} className={`px-3 py-2.5 ${borderClass} text-center font-mono text-[0.9em]`}>
            {matchedCnt}/{totalCnt}
          </td>
        );
      case 'code':
        return (
          <td key={colKey} className={`px-3 py-2.5 ${borderClass} text-center font-mono`}>
            {totalCnt}
          </td>
        );
      case 'name':
        return (
          <td key={colKey} className={`px-4 py-2.5 ${borderClass}`}>
            {label}
          </td>
        );
      case 'fattening':
      case 'fish':
      case 'duck':
      case 'pets':
        return (
          <td
            key={colKey}
            className={`px-3 py-2.5 ${borderClass} font-mono tabular-nums text-left`}
            dir="ltr"
          >
            {formatCell(totalsGroup[colKey])}
          </td>
        );
      case 'total':
        return (
          <td
            key={colKey}
            className={`px-4 py-2.5 ${borderClass} font-mono tabular-nums text-left ${totalHighlightClass}`}
            dir="ltr"
          >
            {totalsGroup.grandTotal.toFixed(decimals)}
          </td>
        );
    }
  };

  return (
    <div className="space-y-4 animate-in fade-in duration-200" dir="rtl">
      {/* Control Toolbar */}
      <div className="px-2 sm:px-4 pt-3 flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-[#002060] text-amber-300 flex items-center justify-center shrink-0">
            <Scale className="w-4 h-4" />
          </div>
          <div>
            <h4 className="text-sm sm:text-base font-bold text-slate-900">
              صفحة تجميع الكنترول حسب الصنف ونوعه
            </h4>
            <p className="text-xs text-slate-500">
              إجمالي كل صنف في عمود نوعه مع استكمال كود الدريف تلقائياً ({matchedCount} من {pivot.rows.length} صنف · الوحدة: {displayUnitLabel})
            </p>
          </div>
        </div>

        {/* Export, Import Codes & Copy buttons */}
        <div className="flex items-center flex-wrap gap-2">
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleImportCatalogFile}
            className="hidden"
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            className="px-3 py-1.5 text-xs font-semibold bg-indigo-50 text-indigo-800 border border-indigo-200 hover:bg-indigo-100 rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer"
            title="استيراد أو تحديث جدول أكواد الدريف من ملف Excel أو CSV"
          >
            <Upload className="w-3.5 h-3.5 text-indigo-600" />
            <span>استيراد شيت أكواد الدريف</span>
          </button>

          <button
            onClick={handleCopyPivot}
            className="px-3 py-1.5 text-xs font-semibold bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer"
          >
            {copied ? (
              <>
                <Check className="w-3.5 h-3.5 text-emerald-600" />
                <span className="text-emerald-700">تم النسخ!</span>
              </>
            ) : (
              <>
                <Copy className="w-3.5 h-3.5 text-slate-500" />
                <span>نسخ تجميع الكنترول</span>
              </>
            )}
          </button>

          <button
            onClick={handleDownloadPivotCsv}
            className="px-3 py-1.5 text-xs font-semibold bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer"
          >
            <Download className="w-3.5 h-3.5 text-slate-500" />
            <span>CSV</span>
          </button>

          <button
            onClick={handleDownloadPivotExcel}
            className="px-3.5 py-1.5 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg shadow-xs transition-all flex items-center gap-1.5 cursor-pointer"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>تصدير صفحة الكنترول (.xlsx)</span>
          </button>
        </div>
      </div>

      {importToast && (
        <div className="mx-4 sm:mx-6 px-3.5 py-2 bg-emerald-50 border border-emerald-200 rounded-lg text-xs font-semibold text-emerald-900 flex items-center gap-2">
          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>{importToast}</span>
        </div>
      )}

      {/* Filter / Unit / Metric / Column Order Controls */}
      <div className="px-2 sm:px-4 flex flex-wrap items-center justify-between gap-3">
        <div className="relative flex-1 min-w-[220px] max-w-sm">
          <Search className="w-4 h-4 text-slate-400 absolute right-2.5 top-2.5" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="بحث بكود الدريف أو كود الصنف أو اسمه..."
            className="w-full pr-8 pl-3 py-1.5 border border-slate-200 rounded-lg text-xs focus:outline-hidden focus:border-emerald-500 bg-white"
          />
          {searchQuery && (
            <button
              onClick={() => setSearchQuery('')}
              className="absolute left-2.5 top-2 text-slate-400 hover:text-slate-600 text-xs"
            >
              ✕
            </button>
          )}
        </div>

        <div className="flex items-center flex-wrap gap-2 text-xs">
          {/* Column Reorder Toggle */}
          <button
            onClick={() => setShowColumnOrderBar((v) => !v)}
            className={`px-2.5 py-1.5 rounded-lg border font-bold transition-colors flex items-center gap-1.5 cursor-pointer ${
              showColumnOrderBar
                ? 'bg-[#002060] text-amber-300 border-[#002060]'
                : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
            }`}
            title="إعادة ترتيب أعمدة الجدول وتصدير الإكسل"
          >
            <Columns3 className="w-3.5 h-3.5" />
            <span>ترتيب الأعمدة</span>
          </button>

          {/* Metric switcher (Actual vs Target) */}
          {pivot.hasActualData && pivot.hasTargetData && (
            <div className="inline-flex rounded-lg border border-slate-200 bg-slate-100 p-0.5">
              <button
                onClick={() => setMetric('actual')}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  metric === 'actual'
                    ? 'bg-white text-slate-900 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                الفعلي (Actual)
              </button>
              <button
                onClick={() => setMetric('target')}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  metric === 'target'
                    ? 'bg-white text-slate-900 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                المستهدف (Target)
              </button>
            </div>
          )}

          {/* Unit switcher (Ton ÷1000 vs KG) */}
          {baseIsKg && (
            <div className="inline-flex rounded-lg border border-slate-200 bg-slate-100 p-0.5">
              <button
                onClick={() => setUseTons(true)}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  useTons
                    ? 'bg-[#002060] text-amber-300 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                بالطن (÷ 1000)
              </button>
              <button
                onClick={() => setUseTons(false)}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  !useTons
                    ? 'bg-[#002060] text-amber-300 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                بالكيلوجرام ({pivot.unit})
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Column Ordering Interactive Bar */}
      {showColumnOrderBar && (
        <div className="mx-4 sm:mx-6 p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
              <Columns3 className="w-3.5 h-3.5 text-[#002060]" />
              <span>اسحب رأس أي عمود في الجدول أو استخدم الأسهم لتغيير ترتيب الأعمدة (يُحفظ تلقائياً في الإكسل أيضاً):</span>
            </span>
            {!isDefaultColumnOrder && (
              <button
                onClick={handleResetColumnOrder}
                className="px-2.5 py-1 text-[11px] font-bold text-slate-700 bg-white border border-slate-200 hover:bg-slate-100 rounded-md flex items-center gap-1 cursor-pointer"
              >
                <RotateCcw className="w-3 h-3" />
                <span>إعادة الترتيب الافتراضي</span>
              </button>
            )}
          </div>
          <div className="flex flex-wrap items-center gap-1.5">
            {columnOrder.map((colKey, idx) => (
              <div
                key={colKey}
                draggable
                onDragStart={() => setDraggedColKey(colKey)}
                onDragOver={(e) => {
                  e.preventDefault();
                  setDragOverColKey(colKey);
                }}
                onDrop={(e) => {
                  e.preventDefault();
                  if (draggedColKey) reorderColumnToTarget(draggedColKey, colKey);
                  setDraggedColKey(null);
                  setDragOverColKey(null);
                }}
                onDragEnd={() => {
                  setDraggedColKey(null);
                  setDragOverColKey(null);
                }}
                className={`inline-flex items-center gap-1 px-2 py-1 rounded-md border text-xs font-bold bg-white shadow-2xs cursor-grab active:cursor-grabbing transition-all ${
                  dragOverColKey === colKey && draggedColKey !== colKey
                    ? 'border-[#002060] ring-2 ring-amber-300'
                    : 'border-slate-200 text-slate-800'
                }`}
              >
                <GripVertical className="w-3 h-3 text-slate-400" />
                <span className="text-[10px] font-mono text-slate-400">{idx + 1}.</span>
                <span>{PIVOT_COLUMN_DEFS[colKey].label}</span>
                <div className="flex items-center border-r border-slate-200 pr-1 mr-0.5 gap-0.5">
                  <button
                    onClick={() => moveColumnByOffset(colKey, -1)}
                    disabled={idx === 0}
                    className="p-0.5 rounded hover:bg-slate-100 disabled:opacity-25 disabled:cursor-not-allowed cursor-pointer"
                    title="تحريك لليمين"
                  >
                    <ChevronRight className="w-3.5 h-3.5 text-slate-700" />
                  </button>
                  <button
                    onClick={() => moveColumnByOffset(colKey, 1)}
                    disabled={idx === columnOrder.length - 1}
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

      {/* Excel-Style Control Aggregation Table with Dynamic Column Ordering */}
      <div className="px-2 sm:px-4 pb-4">
        <div className="border border-slate-300 rounded-lg overflow-x-auto shadow-xs bg-white">
          <table className={`w-full text-right border-collapse ${tableFontClass}`}>
            <thead>
              <tr className="bg-[#002060] text-amber-300 select-none">
                {columnOrder.map((colKey, idx) => {
                  const def = PIVOT_COLUMN_DEFS[colKey];
                  const isLast = idx === columnOrder.length - 1;
                  const isDragOver = dragOverColKey === colKey && draggedColKey !== colKey;
                  const widthClass =
                    colKey === 'name'
                      ? 'min-w-[210px] px-4'
                      : colKey === 'driveCode' || colKey === 'code'
                      ? 'w-28 px-3 text-center'
                      : colKey === 'total'
                      ? 'px-4 text-center'
                      : 'px-3 text-center';

                  return (
                    <th
                      key={colKey}
                      draggable
                      onDragStart={() => setDraggedColKey(colKey)}
                      onDragOver={(e) => {
                        e.preventDefault();
                        setDragOverColKey(colKey);
                      }}
                      onDragLeave={() => {
                        if (dragOverColKey === colKey) setDragOverColKey(null);
                      }}
                      onDrop={(e) => {
                        e.preventDefault();
                        if (draggedColKey) {
                          reorderColumnToTarget(draggedColKey, colKey);
                        }
                        setDraggedColKey(null);
                        setDragOverColKey(null);
                      }}
                      onDragEnd={() => {
                        setDraggedColKey(null);
                        setDragOverColKey(null);
                      }}
                      onClick={() => handleSort(colKey)}
                      title="انقر للترتيب، أو اسحب رأس العمود لإعادة ترتيب الأعمدة"
                      className={`py-3 font-extrabold ${
                        isLast ? '' : 'border-l border-slate-600/60'
                      } cursor-pointer hover:bg-[#002b80] transition-colors whitespace-nowrap group ${widthClass} ${
                        isDragOver ? 'ring-2 ring-inset ring-amber-300 bg-[#003399]' : ''
                      }`}
                    >
                      <div
                        className={`flex items-center gap-1.5 ${
                          colKey === 'name' ? 'justify-between' : 'justify-center'
                        }`}
                      >
                        <span className="inline-flex items-center gap-1">
                          <GripVertical className="w-3 h-3 opacity-35 group-hover:opacity-90 cursor-grab shrink-0" />
                          <span>{def.label}</span>
                        </span>
                        {renderSortIcon(colKey)}
                      </div>
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-300">
              {displayedRows.length === 0 ? (
                <tr>
                  <td
                    colSpan={columnOrder.length}
                    className="py-10 text-center text-slate-500 text-sm"
                  >
                    لا توجد أصناف مطابقة للبحث.
                  </td>
                </tr>
              ) : (
                <>
                  {regularRows.map((row: ControlPivotRow, idx: number) => {
                    const rowKey = getRowKey(row);
                    return (
                      <tr
                        key={`${rowKey}_${idx}`}
                        className="hover:bg-slate-50/90 transition-colors"
                      >
                        {columnOrder.map((colKey, colIdx) =>
                          renderDataCell(colKey, row, colIdx === columnOrder.length - 1, false)
                        )}
                      </tr>
                    );
                  })}

                  {/* Total WITHOUT Water (only separated when Water row exists) */}
                  {waterRows.length > 0 && (
                    <tr className="bg-[#fff2cc] border-t-2 border-slate-700 font-extrabold text-slate-950">
                      {columnOrder.map((colKey, colIdx) =>
                        renderFooterCell(
                          colKey,
                          colIdx === columnOrder.length - 1,
                          totalsWithoutWater,
                          `الإجمالي (بدون الماء) (${displayUnitLabel})`,
                          regularMatchedCount,
                          regularRows.length,
                          'bg-[#ffe699]'
                        )
                      )}
                    </tr>
                  )}

                  {/* Separated Water row(s) placed below الإجمالي (بدون الماء) */}
                  {waterRows.map((row: ControlPivotRow, idx: number) => {
                    const rowKey = getRowKey(row);
                    return (
                      <tr
                        key={`water_${rowKey}_${idx}`}
                        className="hover:bg-sky-100/70 transition-colors border-t border-sky-200"
                      >
                        {columnOrder.map((colKey, colIdx) =>
                          renderDataCell(colKey, row, colIdx === columnOrder.length - 1, true)
                        )}
                      </tr>
                    );
                  })}

                  {/* Final Grand Total Row (With Water if Water exists, else standard Grand Total) */}
                  <tr className="bg-[#ffe699] border-t-2 border-slate-800 font-extrabold text-slate-950">
                    {columnOrder.map((colKey, colIdx) =>
                      renderFooterCell(
                        colKey,
                        colIdx === columnOrder.length - 1,
                        totalsWithWater,
                        waterRows.length > 0
                          ? `الإجمالي الكلي (بالماء) (${displayUnitLabel})`
                          : `الإجمالي الكلي (${displayUnitLabel})`,
                        filteredMatchedCount,
                        displayedRows.length,
                        'bg-[#ffd966]'
                      )
                    )}
                  </tr>
                </>
              )}
            </tbody>
          </table>
        </div>
        <p className="mt-2 text-[11px] text-slate-500">
          💡 الترتيب الافتراضي في عمود <strong>الاجمالى</strong> من الأكبر فوق إلى الأصغر، ويمكنك سحب أي رأس عمود أو استخدام زر <strong>ترتيب الأعمدة</strong> لإعادة ترتيب الأعمدة بحرية.
        </p>
      </div>
    </div>
  );
};
