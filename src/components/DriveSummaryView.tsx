import React, { useMemo, useRef, useState } from 'react';
import {
  Download,
  Copy,
  Check,
  Search,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  FileSpreadsheet,
  Upload,
  CheckCircle2,
  Printer,
  Maximize2,
  Minimize2,
  ZoomIn,
  ZoomOut,
  ListOrdered,
} from 'lucide-react';
import { SheetData } from '../types/converter';
import {
  calculateControlPivot,
  ControlPivotRow,
  isWaterPivotRow,
  loadStoredDriveOverrides,
  loadStoredCustomCatalog,
} from '../utils/controlPivotCalculator';
import {
  DriveCodeEntry,
  DRIVE_OVERRIDES_STORAGE_KEY,
  DRIVE_CUSTOM_CATALOG_STORAGE_KEY,
  lookupDriveCatalogInfo,
  normalizeItemText,
  parseDriveCodesFile,
} from '../utils/driveCodesCatalog';
import { exportToCsv, copyTableToClipboard } from '../utils/excelExport';
import {
  DriveSummaryExportRow,
  exportDriveSummaryStyledExcel,
} from '../utils/styledExcelExport';

interface DriveSummaryViewProps {
  sheet: SheetData;
  fileName?: string;
  tableFontClass?: string;
  fontPx?: number;
  canIncreaseFont?: boolean;
  canDecreaseFont?: boolean;
  onIncreaseFont?: () => void;
  onDecreaseFont?: () => void;
  onResetFont?: () => void;
}

type NameMode = 'short' | 'full' | 'original';
type SortField = 'default' | 'code' | 'driveCode' | 'name' | 'total';

interface SummaryDisplayRow extends ControlPivotRow {
  shortName: string;
  fullName: string;
  displayName: string;
  catalogOrder: number;
  isWater: boolean;
}

function formatCleanNumber(val: number, decimals: number): string {
  if (Math.abs(val) < 1e-12) return '0';
  const fixed = val.toFixed(decimals);
  return fixed.replace(/\.?0+$/, '') || '0';
}

function getDriveSortRank(driveCode: string): { group: number; num: number; str: string } {
  const trimmed = (driveCode || '').trim();
  if (!trimmed) {
    return { group: 9, num: Number.MAX_SAFE_INTEGER, str: '' };
  }
  const num = parseFloat(trimmed);
  if (isNaN(num)) {
    return { group: 8, num: Number.MAX_SAFE_INTEGER, str: trimmed };
  }
  // Macros (< 300) and Rejected macros (1001..1099) appear before Additives (300..999)
  if (num < 300 || (num >= 1000 && num < 1100)) {
    return { group: 0, num, str: trimmed };
  }
  if (num >= 300 && num < 1000) {
    return { group: 1, num, str: trimmed };
  }
  return { group: 2, num, str: trimmed };
}

export const DriveSummaryView: React.FC<DriveSummaryViewProps> = ({
  sheet,
  fileName = 'drive_summary.xlsx',
  tableFontClass = 'text-[13px]',
  fontPx,
  canIncreaseFont = false,
  canDecreaseFont = false,
  onIncreaseFont,
  onDecreaseFont,
  onResetFont,
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [metric, setMetric] = useState<'actual' | 'target'>('actual');

  const baseIsKg = useMemo(() => {
    const joined = (sheet.headers || []).join(' ').toLowerCase();
    return (
      joined.includes('kg') ||
      joined.includes('كجم') ||
      (!joined.includes('ton') && !joined.includes('طن'))
    );
  }, [sheet.headers]);

  const [useTons, setUseTons] = useState<boolean>(baseIsKg);
  const [nameMode, setNameMode] = useState<NameMode>('short');
  const [separateWater, setSeparateWater] = useState<boolean>(true);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);
  const [sortConfig, setSortConfig] = useState<{
    field: SortField;
    direction: 'asc' | 'desc';
  }>({ field: 'default', direction: 'asc' });

  const [copied, setCopied] = useState(false);
  const [overrides, setOverrides] = useState<Record<string, string>>(() =>
    loadStoredDriveOverrides()
  );
  const [customCatalog, setCustomCatalog] = useState<DriveCodeEntry[]>(() =>
    loadStoredCustomCatalog()
  );
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

  const enrichedRows: SummaryDisplayRow[] = useMemo(() => {
    return pivot.rows.map((r) => {
      const info = lookupDriveCatalogInfo(r.name, r.code, overrides, customCatalog);
      const displayName =
        nameMode === 'short'
          ? info.shortName || r.name
          : nameMode === 'full'
          ? info.fullName || r.name
          : r.name;

      return {
        ...r,
        driveCode: info.driveCode || r.driveCode,
        shortName: info.shortName || r.name,
        fullName: info.fullName || r.name,
        displayName,
        catalogOrder: info.catalogOrder,
        isWater: isWaterPivotRow({ ...r, driveCode: info.driveCode || r.driveCode }),
      };
    });
  }, [pivot.rows, overrides, customCatalog, nameMode]);

  const displayedRows = useMemo(() => {
    let items = [...enrichedRows];

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      items = items.filter(
        (r) =>
          r.displayName.toLowerCase().includes(q) ||
          r.name.toLowerCase().includes(q) ||
          r.code.toLowerCase().includes(q) ||
          r.driveCode.toLowerCase().includes(q)
      );
    }

    const dir = sortConfig.direction === 'asc' ? 1 : -1;

    items.sort((a, b) => {
      if (sortConfig.field === 'default') {
        const catDiff = a.catalogOrder - b.catalogOrder;
        const rankA = getDriveSortRank(a.driveCode);
        const rankB = getDriveSortRank(b.driveCode);
        if (rankA.group !== rankB.group) return rankA.group - rankB.group;
        if (rankA.group === 0 && catDiff !== 0) return catDiff;
        if (rankA.num !== rankB.num) return rankA.num - rankB.num;
        return b.total - a.total;
      }

      if (sortConfig.field === 'total') {
        return (a.total - b.total) * dir;
      }

      if (sortConfig.field === 'code' || sortConfig.field === 'driveCode') {
        const valA = sortConfig.field === 'code' ? a.code : a.driveCode;
        const valB = sortConfig.field === 'code' ? b.code : b.driveCode;
        if (!valA && !valB) return b.total - a.total;
        if (!valA) return 1;
        if (!valB) return -1;
        const numA = parseFloat(valA);
        const numB = parseFloat(valB);
        if (!isNaN(numA) && !isNaN(numB)) {
          return (numA - numB) * dir;
        }
        return valA.localeCompare(valB, 'ar') * dir;
      }

      return a.displayName.localeCompare(b.displayName, 'ar') * dir;
    });

    return items;
  }, [enrichedRows, searchQuery, sortConfig]);

  const { regularRows, waterRows, totalWithoutWater, totalWithWater } = useMemo(() => {
    const reg = separateWater ? displayedRows.filter((r) => !r.isWater) : displayedRows;
    const wat = separateWater ? displayedRows.filter((r) => r.isWater) : [];
    const sumReg = reg.reduce((s, r) => s + r.total, 0);
    const sumAll = displayedRows.reduce((s, r) => s + r.total, 0);
    return {
      regularRows: reg,
      waterRows: wat,
      totalWithoutWater: sumReg,
      totalWithWater: sumAll,
    };
  }, [displayedRows, separateWater]);

  const matchedCount = useMemo(
    () => displayedRows.filter((r) => Boolean(r.driveCode)).length,
    [displayedRows]
  );

  const handleSort = (field: Exclude<SortField, 'default'>) => {
    setSortConfig((prev) => {
      if (prev.field === field) {
        if (prev.direction === 'asc') {
          return { field, direction: 'desc' };
        }
        return { field: 'default', direction: 'asc' };
      }
      return {
        field,
        direction: field === 'total' ? 'desc' : 'asc',
      };
    });
  };

  const renderSortIcon = (field: Exclude<SortField, 'default'>) => {
    if (sortConfig.field !== field) {
      return <ArrowUpDown className="w-3.5 h-3.5 text-slate-700 opacity-60" />;
    }
    return sortConfig.direction === 'asc' ? (
      <ArrowUp className="w-3.5 h-3.5 text-slate-950" />
    ) : (
      <ArrowDown className="w-3.5 h-3.5 text-slate-950" />
    );
  };

  const getRowKey = (row: SummaryDisplayRow) =>
    `${row.code}::${normalizeItemText(row.name)}`;

  const handleStartEditDrive = (row: SummaryDisplayRow) => {
    setEditingRowKey(getRowKey(row));
    setEditingDriveValue(row.driveCode || '');
  };

  const handleSaveDriveCode = (row: SummaryDisplayRow) => {
    const cleaned = editingDriveValue.trim();
    const compositeKey = `${(row.code || '').trim()}::${normalizeItemText(row.name)}`;
    const nameOnlyKey = normalizeItemText(row.name);

    const nextOverrides = { ...overrides };
    if (cleaned) {
      nextOverrides[compositeKey] = cleaned;
      if (nameOnlyKey) nextOverrides[nameOnlyKey] = cleaned;
    } else {
      delete nextOverrides[compositeKey];
      delete nextOverrides[nameOnlyKey];
    }

    setOverrides(nextOverrides);
    try {
      localStorage.setItem(DRIVE_OVERRIDES_STORAGE_KEY, JSON.stringify(nextOverrides));
    } catch {
      // ignore storage errors
    }
    setEditingRowKey(null);
  };

  const handleImportDriveFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const entries = await parseDriveCodesFile(file);
      if (entries.length === 0) {
        setImportToast('لم يتم العثور على أكواد صالحة في الملف المختار.');
        setTimeout(() => setImportToast(null), 4000);
        return;
      }

      const mergedMap = new Map<string, DriveCodeEntry>();
      customCatalog.forEach((item) => {
        mergedMap.set(`${item.jdeCode}::${normalizeItemText(item.name)}`, item);
      });
      entries.forEach((item) => {
        mergedMap.set(`${item.jdeCode}::${normalizeItemText(item.name)}`, item);
      });
      const mergedList = Array.from(mergedMap.values());

      setCustomCatalog(mergedList);
      try {
        localStorage.setItem(
          DRIVE_CUSTOM_CATALOG_STORAGE_KEY,
          JSON.stringify(mergedList)
        );
      } catch {
        // ignore storage errors
      }
      setImportToast(`تم استيراد ${entries.length} كود درايف وتحديث الجدول تلقائياً!`);
      setTimeout(() => setImportToast(null), 5000);
    } catch {
      setImportToast('حدث خطأ أثناء قراءة ملف أكواد الدريف.');
      setTimeout(() => setImportToast(null), 4000);
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const buildPlainSheetData = (): SheetData => {
    const toRow = (r: SummaryDisplayRow): (string | number)[] => [
      r.code || '',
      r.driveCode || '',
      r.displayName,
      formatCleanNumber(r.total, decimals),
    ];

    const rows: (string | number)[][] = [];
    let totalsRow: (string | number)[];

    if (separateWater && waterRows.length > 0) {
      regularRows.forEach((r) => rows.push(toRow(r)));
      rows.push([
        '',
        '',
        'الإجمالي (بدون الماء)',
        formatCleanNumber(totalWithoutWater, decimals),
      ]);
      waterRows.forEach((w) => rows.push(toRow(w)));
      totalsRow = [
        '',
        '',
        'الإجمالي الكلي (بالماء)',
        formatCleanNumber(totalWithWater, decimals),
      ];
    } else {
      displayedRows.forEach((r) => rows.push(toRow(r)));
      totalsRow = [
        '',
        '',
        'الإجمالي الكلي',
        formatCleanNumber(totalWithWater, decimals),
      ];
    }

    return {
      sheetName: 'ملخص_اكواد_الدرايف',
      description: 'صفحة ملخص الأصناف حسب كود الصنف وكود درايف والإجمالي',
      isRtl: true,
      headers: ['كود الصنف', 'كود درايف', 'إسم الصنف', 'الاجمالى'],
      rows,
      totalsRow,
    };
  };

  const handleCopy = async () => {
    const ok = await copyTableToClipboard(buildPlainSheetData());
    if (ok) {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  const handleExportCsv = () => {
    const baseName = fileName.replace(/\.(xlsx|pdf|csv)$/i, '');
    exportToCsv(buildPlainSheetData(), `${baseName}_ملخص_الدرايف.csv`);
  };

  const handleExportExcel = async () => {
    const mapExport = (r: SummaryDisplayRow): DriveSummaryExportRow => ({
      code: r.code,
      driveCode: r.driveCode,
      displayName: r.displayName,
      total: r.total,
      isWater: r.isWater,
    });

    await exportDriveSummaryStyledExcel({
      regularRows: regularRows.map(mapExport),
      waterRows: waterRows.map(mapExport),
      displayedRows: displayedRows.map(mapExport),
      totalWithoutWater,
      totalWithWater,
      separateWater,
      decimals,
      fileName,
    });
  };

  const displayUnitLabel = useTons && baseIsKg ? 'طن' : pivot.unit;

  const renderDataRow = (row: SummaryDisplayRow, isWaterRow: boolean) => {
    const rowKey = getRowKey(row);
    const isEditing = editingRowKey === rowKey;

    return (
      <tr
        key={rowKey}
        className={`transition-colors ${
          isWaterRow ? 'bg-sky-50 hover:bg-sky-100/80' : 'bg-white hover:bg-yellow-50/70'
        }`}
      >
        {/* 1. كود الصنف */}
        <td className="px-3 py-1.5 text-center border border-slate-950 font-mono font-bold text-slate-950 tabular-nums whitespace-nowrap">
          {row.code || '-'}
        </td>

        {/* 2. كود درايف (قابل للتعديل بالنقر) */}
        <td
          className="px-3 py-1.5 text-center border border-slate-950 font-mono font-bold text-slate-950 tabular-nums whitespace-nowrap cursor-pointer hover:bg-amber-50"
          onClick={() => {
            if (!isEditing) handleStartEditDrive(row);
          }}
          title="انقر لتعديل كود درايف وحفظه تلقائياً"
        >
          {isEditing ? (
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
              className="w-20 px-1.5 py-0.5 text-center font-mono font-bold text-slate-950 bg-white border-2 border-amber-500 rounded outline-hidden"
            />
          ) : row.driveCode ? (
            row.driveCode
          ) : (
            <span className="text-rose-700 font-bold">+ كود</span>
          )}
        </td>

        {/* 3. إسم الصنف */}
        <td
          className="px-4 py-1.5 text-center border border-slate-950 font-bold text-slate-950"
          title={row.name !== row.displayName ? `الأصل في PDF: ${row.name}` : undefined}
        >
          {row.displayName}
        </td>

        {/* 4. الاجمالى */}
        <td
          className="px-4 py-1.5 text-center border border-slate-950 font-mono font-bold text-slate-950 tabular-nums whitespace-nowrap"
          dir="ltr"
        >
          {formatCleanNumber(row.total, decimals)}
        </td>
      </tr>
    );
  };

  return (
    <div
      className={
        isFullscreen
          ? 'fixed inset-0 z-50 bg-white flex flex-col overflow-hidden'
          : 'flex flex-col'
      }
    >
      {/* Top Toolbar */}
      <div className="p-3 sm:px-6 bg-slate-50 border-b border-slate-200 flex flex-col gap-3">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2.5">
            <div className="p-2 rounded-lg bg-[#ffff00] border border-slate-950 text-slate-950 shadow-xs">
              <ListOrdered className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-extrabold text-slate-900">
                ملخص الأصناف (كود الصنف · كود درايف · الاجمالى)
              </h3>
              <p className="text-xs font-semibold text-slate-600">
                إجمالي الأصناف: {displayedRows.length} · مطابق لكود درايف: {matchedCount}/
                {displayedRows.length} · الوحدة: {displayUnitLabel}
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 text-xs">
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleImportDriveFile}
              className="hidden"
            />
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="px-2.5 py-1.5 font-bold text-indigo-800 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 rounded-md transition-colors flex items-center gap-1.5 cursor-pointer"
              title="استيراد ملف إكسل يحتوي على أكواد درايف لتحديث الكتالوج"
            >
              <Upload className="w-3.5 h-3.5" />
              <span>استيراد أكواد درايف</span>
            </button>

            <button
              type="button"
              onClick={handleCopy}
              className="px-2.5 py-1.5 font-bold bg-white text-slate-800 border border-slate-200 hover:bg-slate-100 rounded-md transition-colors flex items-center gap-1.5 cursor-pointer"
            >
              {copied ? (
                <>
                  <Check className="w-3.5 h-3.5 text-emerald-600" />
                  <span className="text-emerald-700">تم النسخ!</span>
                </>
              ) : (
                <>
                  <Copy className="w-3.5 h-3.5 text-slate-600" />
                  <span>نسخ</span>
                </>
              )}
            </button>

            <button
              type="button"
              onClick={handleExportCsv}
              className="px-2.5 py-1.5 font-bold bg-white text-slate-800 border border-slate-200 hover:bg-slate-100 rounded-md transition-colors flex items-center gap-1.5 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5 text-slate-600" />
              <span>CSV</span>
            </button>

            <button
              type="button"
              onClick={() => window.print()}
              className="px-2.5 py-1.5 font-bold bg-white text-slate-800 border border-slate-200 hover:bg-slate-100 rounded-md transition-colors flex items-center gap-1.5 cursor-pointer"
            >
              <Printer className="w-3.5 h-3.5 text-slate-600" />
              <span>طباعة</span>
            </button>

            {isFullscreen && onIncreaseFont && onDecreaseFont && onResetFont && (
              <div className="inline-flex items-center bg-white border border-slate-200 rounded-md p-0.5">
                <button
                  type="button"
                  onClick={onDecreaseFont}
                  disabled={!canDecreaseFont}
                  className="px-1.5 py-1 text-xs font-bold text-slate-700 hover:bg-slate-100 rounded disabled:opacity-35 cursor-pointer"
                  title="تصغير الخط"
                >
                  <ZoomOut className="w-3.5 h-3.5" />
                </button>
                <button
                  type="button"
                  onClick={onResetFont}
                  className="px-1.5 py-1 text-xs font-mono font-bold text-slate-800 border-x border-slate-200 cursor-pointer"
                >
                  {fontPx ?? 13}px
                </button>
                <button
                  type="button"
                  onClick={onIncreaseFont}
                  disabled={!canIncreaseFont}
                  className="px-1.5 py-1 text-xs font-bold text-slate-700 hover:bg-slate-100 rounded disabled:opacity-35 cursor-pointer"
                  title="تكبير الخط"
                >
                  <ZoomIn className="w-3.5 h-3.5 text-emerald-700" />
                </button>
              </div>
            )}

            <button
              type="button"
              onClick={() => setIsFullscreen((v) => !v)}
              className="px-2.5 py-1.5 font-bold bg-white text-slate-800 border border-slate-200 hover:bg-slate-100 rounded-md transition-colors flex items-center gap-1.5 cursor-pointer"
              title={isFullscreen ? 'تصغير الشاشة' : 'عرض بملء الشاشة'}
            >
              {isFullscreen ? (
                <>
                  <Minimize2 className="w-3.5 h-3.5" />
                  <span>تصغير</span>
                </>
              ) : (
                <>
                  <Maximize2 className="w-3.5 h-3.5" />
                  <span>ملء الشاشة</span>
                </>
              )}
            </button>

            <button
              type="button"
              onClick={handleExportExcel}
              className="px-3 py-1.5 font-bold text-slate-950 bg-[#ffff00] hover:bg-yellow-300 border border-slate-950 rounded-md shadow-xs transition-all flex items-center gap-1.5 cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>تصدير Excel (.xlsx)</span>
            </button>
          </div>
        </div>

        {/* Filters & Controls Row */}
        <div className="flex flex-wrap items-center justify-between gap-2.5 text-xs">
          {/* Search Input */}
          <div className="relative flex-1 min-w-[220px] max-w-sm">
            <Search className="w-4 h-4 text-slate-400 absolute right-2.5 top-2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="بحث بكود الصنف أو كود درايف أو إسم الصنف..."
              className="w-full pr-8 pl-7 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-bold focus:outline-hidden focus:border-slate-900"
            />
            {searchQuery && (
              <button
                type="button"
                onClick={() => setSearchQuery('')}
                className="absolute left-2.5 top-1.5 text-slate-400 hover:text-slate-700 font-bold"
              >
                ✕
              </button>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {/* Actual vs Target Toggle */}
            {pivot.hasActualData && pivot.hasTargetData && (
              <div className="inline-flex items-center bg-white border border-slate-200 rounded-lg p-0.5">
                <button
                  type="button"
                  onClick={() => setMetric('actual')}
                  className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                    metric === 'actual'
                      ? 'bg-slate-900 text-white'
                      : 'text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  الفعلي (Actual)
                </button>
                <button
                  type="button"
                  onClick={() => setMetric('target')}
                  className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                    metric === 'target'
                      ? 'bg-slate-900 text-white'
                      : 'text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  المستهدف (Target)
                </button>
              </div>
            )}

            {/* Ton vs KG Toggle */}
            {baseIsKg && (
              <div className="inline-flex items-center bg-white border border-slate-200 rounded-lg p-0.5">
                <button
                  type="button"
                  onClick={() => setUseTons(true)}
                  className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                    useTons
                      ? 'bg-[#002060] text-amber-300'
                      : 'text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  بالطن (÷ 1000)
                </button>
                <button
                  type="button"
                  onClick={() => setUseTons(false)}
                  className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                    !useTons
                      ? 'bg-[#002060] text-amber-300'
                      : 'text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  بالكيلوجرام (كجم)
                </button>
              </div>
            )}

            {/* Name Format Mode */}
            <div className="inline-flex items-center bg-white border border-slate-200 rounded-lg p-0.5">
              <button
                type="button"
                onClick={() => setNameMode('short')}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  nameMode === 'short'
                    ? 'bg-[#ffff00] text-slate-950 border border-slate-900'
                    : 'text-slate-700 hover:bg-slate-100'
                }`}
              >
                عربي (الدريف)
              </button>
              <button
                type="button"
                onClick={() => setNameMode('full')}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  nameMode === 'full'
                    ? 'bg-[#ffff00] text-slate-950 border border-slate-900'
                    : 'text-slate-700 hover:bg-slate-100'
                }`}
              >
                الاسم الكامل
              </button>
              <button
                type="button"
                onClick={() => setNameMode('original')}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  nameMode === 'original'
                    ? 'bg-[#ffff00] text-slate-950 border border-slate-900'
                    : 'text-slate-700 hover:bg-slate-100'
                }`}
              >
                الأصلي (PDF)
              </button>
            </div>

            {/* Separate Water Checkbox */}
            <label className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg font-bold text-slate-800 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={separateWater}
                onChange={(e) => setSeparateWater(e.target.checked)}
                className="rounded border-slate-300 text-slate-900 focus:ring-0"
              />
              <span>فصل الماء أسفل الإجمالي</span>
            </label>

            {sortConfig.field !== 'default' && (
              <button
                type="button"
                onClick={() => setSortConfig({ field: 'default', direction: 'asc' })}
                className="px-2.5 py-1.5 bg-white border border-slate-200 hover:bg-slate-100 rounded-lg font-bold text-slate-700 cursor-pointer"
              >
                إرجاع الترتيب القياسي
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Import Toast */}
      {importToast && (
        <div className="px-4 py-2 bg-emerald-50 border-b border-emerald-200 text-emerald-900 text-xs font-bold flex items-center gap-2">
          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>{importToast}</span>
        </div>
      )}

      {/* Main 4-Column Sheet Table (Matches Screenshot Layout) */}
      <div
        className={`overflow-auto bg-white p-3 sm:p-4 ${
          isFullscreen ? 'flex-1' : 'max-h-[640px]'
        }`}
        dir="rtl"
      >
        <div className="max-w-4xl mx-auto">
          <table className={`w-full border-collapse border-2 border-slate-950 ${tableFontClass}`}>
            <thead className="sticky top-0 z-20">
              <tr className="bg-[#ffff00] text-slate-950 font-black">
                <th
                  onClick={() => handleSort('code')}
                  className="w-36 px-3 py-2 text-center border border-slate-950 cursor-pointer select-none hover:bg-yellow-300 transition-colors"
                >
                  <div className="inline-flex items-center justify-center gap-1.5">
                    <span>كود الصنف</span>
                    {renderSortIcon('code')}
                  </div>
                </th>

                <th
                  onClick={() => handleSort('driveCode')}
                  className="w-32 px-3 py-2 text-center border border-slate-950 cursor-pointer select-none hover:bg-yellow-300 transition-colors"
                >
                  <div className="inline-flex items-center justify-center gap-1.5">
                    <span>كود درايف</span>
                    {renderSortIcon('driveCode')}
                  </div>
                </th>

                <th
                  onClick={() => handleSort('name')}
                  className="px-4 py-2 text-center border border-slate-950 cursor-pointer select-none hover:bg-yellow-300 transition-colors"
                >
                  <div className="inline-flex items-center justify-center gap-1.5">
                    <span>إسم الصنف</span>
                    {renderSortIcon('name')}
                  </div>
                </th>

                <th
                  onClick={() => handleSort('total')}
                  className="w-40 px-4 py-2 text-center border border-slate-950 cursor-pointer select-none hover:bg-yellow-300 transition-colors"
                >
                  <div className="inline-flex items-center justify-center gap-1.5">
                    <span>الاجمالى</span>
                    {renderSortIcon('total')}
                  </div>
                </th>
              </tr>
            </thead>

            <tbody>
              {displayedRows.length === 0 ? (
                <tr>
                  <td
                    colSpan={4}
                    className="px-4 py-8 text-center text-slate-500 font-bold border border-slate-950 bg-white"
                  >
                    لا توجد أصناف مطابقة للبحث الحالي.
                  </td>
                </tr>
              ) : separateWater && waterRows.length > 0 ? (
                <>
                  {regularRows.map((row) => renderDataRow(row, false))}

                  {/* Subtotal excluding water */}
                  <tr className="bg-amber-100 text-slate-950 font-black">
                    <td className="px-3 py-2 text-center border border-slate-950 font-mono tabular-nums">
                      -
                    </td>
                    <td className="px-3 py-2 text-center border border-slate-950 font-mono tabular-nums">
                      {regularRows.filter((r) => Boolean(r.driveCode)).length}/{regularRows.length}
                    </td>
                    <td className="px-4 py-2 text-center border border-slate-950">
                      الإجمالي (بدون الماء) ({displayUnitLabel})
                    </td>
                    <td
                      className="px-4 py-2 text-center border border-slate-950 font-mono tabular-nums"
                      dir="ltr"
                    >
                      {formatCleanNumber(totalWithoutWater, decimals)}
                    </td>
                  </tr>

                  {waterRows.map((row) => renderDataRow(row, true))}

                  {/* Grand Total including water */}
                  <tr className="bg-[#ffff00] text-slate-950 font-black">
                    <td className="px-3 py-2 text-center border border-slate-950 font-mono tabular-nums">
                      Σ
                    </td>
                    <td className="px-3 py-2 text-center border border-slate-950 font-mono tabular-nums">
                      {matchedCount}/{displayedRows.length}
                    </td>
                    <td className="px-4 py-2 text-center border border-slate-950">
                      الإجمالي الكلي (بالماء) ({displayUnitLabel})
                    </td>
                    <td
                      className="px-4 py-2 text-center border border-slate-950 font-mono tabular-nums"
                      dir="ltr"
                    >
                      {formatCleanNumber(totalWithWater, decimals)}
                    </td>
                  </tr>
                </>
              ) : (
                <>
                  {displayedRows.map((row) => renderDataRow(row, row.isWater))}

                  <tr className="bg-[#ffff00] text-slate-950 font-black">
                    <td className="px-3 py-2 text-center border border-slate-950 font-mono tabular-nums">
                      Σ
                    </td>
                    <td className="px-3 py-2 text-center border border-slate-950 font-mono tabular-nums">
                      {matchedCount}/{displayedRows.length}
                    </td>
                    <td className="px-4 py-2 text-center border border-slate-950">
                      الإجمالي الكلي ({displayUnitLabel})
                    </td>
                    <td
                      className="px-4 py-2 text-center border border-slate-950 font-mono tabular-nums"
                      dir="ltr"
                    >
                      {formatCleanNumber(totalWithWater, decimals)}
                    </td>
                  </tr>
                </>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Footer Note */}
      <div className="px-4 py-2.5 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs font-semibold text-slate-600">
        <span>
          💡 انقر على أي خلية في عمود «كود درايف» لتعديله وحفظه تلقائياً، أو انقر على أي رأس عمود للترتيب.
        </span>
        <span className="font-mono font-bold text-slate-800">
          {displayedRows.length} صنف
        </span>
      </div>
    </div>
  );
};
