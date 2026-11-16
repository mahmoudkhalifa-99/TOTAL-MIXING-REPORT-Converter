import React, { useMemo, useState } from 'react';
import {
  Search,
  Download,
  Copy,
  Check,
  Printer,
  FileSpreadsheet,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Columns3,
  Maximize2,
  Minimize2,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import { SheetData } from '../types/converter';
import {
  calculateSadatControlReport,
  SadatFormulaColumn,
  SadatMatrixRow,
  SADAT_BATCH_CODES_STORAGE_KEY,
  SADAT_META_STORAGE_KEY,
} from '../utils/sadatControlCalculator';
import { exportToCsv, copyTableToClipboard } from '../utils/excelExport';
import { exportSadatControlStyledExcel } from '../utils/styledExcelExport';

interface SadatControlReportViewProps {
  sheet: SheetData;
  fileName?: string;
  documentTitle?: string;
  tableFontClass?: string;
  fontPx?: number;
  canIncreaseFont?: boolean;
  canDecreaseFont?: boolean;
  onIncreaseFont?: () => void;
  onDecreaseFont?: () => void;
  onResetFont?: () => void;
}

type SortTarget =
  | { kind: 'code' }
  | { kind: 'driveCode' }
  | { kind: 'name' }
  | { kind: 'formula'; key: string }
  | { kind: 'fatteningTotal' }
  | { kind: 'fishTotal' }
  | { kind: 'duckTotal' }
  | { kind: 'petsTotal' }
  | { kind: 'grandTotal' };

function loadStoredJson<T>(key: string, fallback: T): T {
  try {
    const raw = typeof localStorage !== 'undefined' ? localStorage.getItem(key) : null;
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

export const SadatControlReportView: React.FC<SadatControlReportViewProps> = ({
  sheet,
  fileName = 'sadat_control_report.xlsx',
  documentTitle,
  tableFontClass = 'text-[11px] sm:text-xs',
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
  const [showDriveCode, setShowDriveCode] = useState<boolean>(false);
  const [useBilingualNames, setUseBilingualNames] = useState<boolean>(true);
  const [separateWater, setSeparateWater] = useState<boolean>(true);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);
  const [copied, setCopied] = useState(false);

  const [sortState, setSortState] = useState<{
    target: SortTarget;
    direction: 'asc' | 'desc';
  }>({ target: { kind: 'grandTotal' }, direction: 'desc' });

  const divisor = useTons && baseIsKg ? 1000 : 1;
  const decimals = useTons && baseIsKg ? 6 : 3;
  const grandTotalDecimals = useTons && baseIsKg ? 4 : 3;

  const report = useMemo(
    () => calculateSadatControlReport(sheet, metric, divisor, documentTitle),
    [sheet, metric, divisor, documentTitle]
  );

  // Ordered formula groups (each active group followed by its total column)
  const allOrderedFormulas = useMemo(
    () => [
      ...report.fatteningFormulas,
      ...report.fishFormulas,
      ...report.duckFormulas,
      ...report.petsFormulas,
    ],
    [report]
  );

  const [batchCodeOverrides, setBatchCodeOverrides] = useState<Record<string, string>>(() =>
    loadStoredJson<Record<string, string>>(SADAT_BATCH_CODES_STORAGE_KEY, {})
  );
  const [editingBatchKey, setEditingBatchKey] = useState<string | null>(null);
  const [editingBatchVal, setEditingBatchVal] = useState('');

  const [meta, setMeta] = useState<{ title: string; date: string; subLabel: string }>(() =>
    loadStoredJson(SADAT_META_STORAGE_KEY, {
      title: `تقرير كنترول السادات ${new Date().getFullYear()}`,
      date: '',
      subLabel: 'supledger',
    })
  );
  const [editingMetaField, setEditingMetaField] = useState<'title' | 'date' | 'subLabel' | null>(
    null
  );
  const [editingMetaVal, setEditingMetaVal] = useState('');

  const saveMetaField = (field: 'title' | 'date' | 'subLabel') => {
    const updated = { ...meta, [field]: editingMetaVal.trim() };
    setMeta(updated);
    try {
      localStorage.setItem(SADAT_META_STORAGE_KEY, JSON.stringify(updated));
    } catch {
      // ignore storage errors
    }
    setEditingMetaField(null);
  };

  const saveBatchCode = (formulaKey: string) => {
    const cleaned = editingBatchVal.replace(/,/g, '').trim();
    const updated = { ...batchCodeOverrides, [formulaKey]: cleaned };
    setBatchCodeOverrides(updated);
    try {
      localStorage.setItem(SADAT_BATCH_CODES_STORAGE_KEY, JSON.stringify(updated));
    } catch {
      // ignore storage errors
    }
    setEditingBatchKey(null);
  };

  const getRowDisplayName = (row: SadatMatrixRow) =>
    useBilingualNames ? row.bilingualName : row.name;

  const displayedRows = useMemo(() => {
    let items = [...report.rows];

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      items = items.filter(
        (r) =>
          r.name.toLowerCase().includes(q) ||
          r.bilingualName.toLowerCase().includes(q) ||
          r.code.toLowerCase().includes(q) ||
          r.driveCode.toLowerCase().includes(q)
      );
    }

    const dir = sortState.direction === 'asc' ? 1 : -1;
    items.sort((a, b) => {
      const t = sortState.target;
      if (t.kind === 'code' || t.kind === 'driveCode') {
        const va = t.kind === 'code' ? a.code : a.driveCode;
        const vb = t.kind === 'code' ? b.code : b.driveCode;
        const na = parseFloat(va);
        const nb = parseFloat(vb);
        if (!isNaN(na) && !isNaN(nb)) return (na - nb) * dir;
        return va.localeCompare(vb, 'ar') * dir;
      }
      if (t.kind === 'name') {
        return getRowDisplayName(a).localeCompare(getRowDisplayName(b), 'ar') * dir;
      }
      if (t.kind === 'formula') {
        const va = a.formulaValues[t.key] || 0;
        const vb = b.formulaValues[t.key] || 0;
        return (va - vb) * dir;
      }
      return (a[t.kind] - b[t.kind]) * dir;
    });

    return items;
  }, [report.rows, searchQuery, sortState, useBilingualNames]);

  const regularRows = useMemo(
    () => (separateWater ? displayedRows.filter((r) => !r.isWater) : displayedRows),
    [displayedRows, separateWater]
  );

  const waterRows = useMemo(
    () => (separateWater ? displayedRows.filter((r) => r.isWater) : []),
    [displayedRows, separateWater]
  );

  const computeSums = (rowsList: SadatMatrixRow[]) => {
    const formulaSums: Record<string, number> = {};
    for (const f of allOrderedFormulas) formulaSums[f.key] = 0;
    let fatteningTotal = 0;
    let fishTotal = 0;
    let duckTotal = 0;
    let petsTotal = 0;
    let grandTotal = 0;

    for (const r of rowsList) {
      for (const f of allOrderedFormulas) {
        formulaSums[f.key] += r.formulaValues[f.key] || 0;
      }
      fatteningTotal += r.fatteningTotal;
      fishTotal += r.fishTotal;
      duckTotal += r.duckTotal;
      petsTotal += r.petsTotal;
      grandTotal += r.grandTotal;
    }

    return { formulaSums, fatteningTotal, fishTotal, duckTotal, petsTotal, grandTotal };
  };

  const sumsWithoutWater = useMemo(
    () => computeSums(regularRows),
    [regularRows, allOrderedFormulas]
  );
  const sumsWithWater = useMemo(
    () => computeSums(displayedRows),
    [displayedRows, allOrderedFormulas]
  );

  const toggleSort = (target: SortTarget) => {
    setSortState((prev) => {
      const same =
        prev.target.kind === target.kind &&
        (prev.target.kind !== 'formula' ||
          (target.kind === 'formula' && prev.target.key === target.key));
      if (same) {
        return { target, direction: prev.direction === 'desc' ? 'asc' : 'desc' };
      }
      const isText =
        target.kind === 'name' || target.kind === 'code' || target.kind === 'driveCode';
      return { target, direction: isText ? 'asc' : 'desc' };
    });
  };

  const renderSortBadge = (target: SortTarget) => {
    const same =
      sortState.target.kind === target.kind &&
      (sortState.target.kind !== 'formula' ||
        (target.kind === 'formula' && sortState.target.key === target.key));
    if (!same) {
      return <ArrowUpDown className="w-3 h-3 opacity-45 shrink-0" />;
    }
    return sortState.direction === 'asc' ? (
      <ArrowUp className="w-3 h-3 text-amber-300 shrink-0" />
    ) : (
      <ArrowDown className="w-3 h-3 text-amber-300 shrink-0" />
    );
  };

  const fmt = (val: number, dec: number = decimals) => {
    if (Math.abs(val) < 1e-9) return '-';
    return val.toFixed(dec);
  };

  const effectiveDate = meta.date || report.detectedDate;

  // Build tabular SheetData for CSV, Clipboard, and XLSX Export
  const buildExportSheetData = (): SheetData => {
    const headers: string[] = ['كود الصنف'];
    if (showDriveCode) headers.push('كود الدريف');
    headers.push('اسم الصنف');

    for (const f of report.fatteningFormulas) headers.push(f.label);
    if (report.hasFattening) headers.push('اجمالي التسمين');

    for (const f of report.fishFormulas) headers.push(f.label);
    if (report.hasFish) headers.push('اجمالي سمك');

    for (const f of report.duckFormulas) headers.push(f.label);
    if (report.hasDuck) headers.push('اجمالي البط');

    for (const f of report.petsFormulas) headers.push(f.label);
    if (report.hasPets) headers.push('اجمالي الاليف');

    headers.push('الإجمالي العام');

    const mapRow = (r: SadatMatrixRow): (string | number)[] => {
      const arr: (string | number)[] = [r.code || ''];
      if (showDriveCode) arr.push(r.driveCode || '');
      arr.push(getRowDisplayName(r));

      for (const f of report.fatteningFormulas) {
        const v = r.formulaValues[f.key] || 0;
        arr.push(Math.abs(v) < 1e-9 ? '-' : Number(v.toFixed(decimals)));
      }
      if (report.hasFattening) {
        arr.push(Math.abs(r.fatteningTotal) < 1e-9 ? '-' : Number(r.fatteningTotal.toFixed(decimals)));
      }

      for (const f of report.fishFormulas) {
        const v = r.formulaValues[f.key] || 0;
        arr.push(Math.abs(v) < 1e-9 ? '-' : Number(v.toFixed(decimals)));
      }
      if (report.hasFish) {
        arr.push(Math.abs(r.fishTotal) < 1e-9 ? '-' : Number(r.fishTotal.toFixed(decimals)));
      }

      for (const f of report.duckFormulas) {
        const v = r.formulaValues[f.key] || 0;
        arr.push(Math.abs(v) < 1e-9 ? '-' : Number(v.toFixed(decimals)));
      }
      if (report.hasDuck) {
        arr.push(Math.abs(r.duckTotal) < 1e-9 ? '-' : Number(r.duckTotal.toFixed(decimals)));
      }

      for (const f of report.petsFormulas) {
        const v = r.formulaValues[f.key] || 0;
        arr.push(Math.abs(v) < 1e-9 ? '-' : Number(v.toFixed(decimals)));
      }
      if (report.hasPets) {
        arr.push(Math.abs(r.petsTotal) < 1e-9 ? '-' : Number(r.petsTotal.toFixed(decimals)));
      }

      arr.push(Number(r.grandTotal.toFixed(grandTotalDecimals)));
      return arr;
    };

    const mapSumRow = (
      label: string,
      sums: ReturnType<typeof computeSums>
    ): (string | number)[] => {
      const arr: (string | number)[] = [''];
      if (showDriveCode) arr.push('');
      arr.push(label);

      for (const f of report.fatteningFormulas) {
        arr.push(Number((sums.formulaSums[f.key] || 0).toFixed(decimals)));
      }
      if (report.hasFattening) arr.push(Number(sums.fatteningTotal.toFixed(decimals)));

      for (const f of report.fishFormulas) {
        arr.push(Number((sums.formulaSums[f.key] || 0).toFixed(decimals)));
      }
      if (report.hasFish) arr.push(Number(sums.fishTotal.toFixed(decimals)));

      for (const f of report.duckFormulas) {
        arr.push(Number((sums.formulaSums[f.key] || 0).toFixed(decimals)));
      }
      if (report.hasDuck) arr.push(Number(sums.duckTotal.toFixed(decimals)));

      for (const f of report.petsFormulas) {
        arr.push(Number((sums.formulaSums[f.key] || 0).toFixed(decimals)));
      }
      if (report.hasPets) arr.push(Number(sums.petsTotal.toFixed(decimals)));

      arr.push(Number(sums.grandTotal.toFixed(grandTotalDecimals)));
      return arr;
    };

    const rows: (string | number)[][] = regularRows.map(mapRow);
    let totalsRow: (string | number)[];

    if (waterRows.length > 0) {
      rows.push(mapSumRow('الإجمالي (بدون الماء)', sumsWithoutWater));
      for (const w of waterRows) rows.push(mapRow(w));
      totalsRow = mapSumRow('الإجمالي الكلي (بالماء)', sumsWithWater);
    } else {
      totalsRow = mapSumRow('الإجمالي الكلي', sumsWithWater);
    }

    return {
      sheetName: 'تقرير_كنترول_السادات',
      description: meta.title,
      isRtl: true,
      headers,
      rows,
      totalsRow,
    };
  };

  const buildSupledgerRow = (): string[] => {
    const cleanSup = (f: SadatFormulaColumn) =>
      (batchCodeOverrides[f.key] ?? f.productCode ?? '').replace(/,/g, '').trim();

    const row: string[] = [meta.subLabel || 'supledger'];
    if (showDriveCode) row.push('');
    row.push('');

    for (const f of report.fatteningFormulas) row.push(cleanSup(f));
    if (report.hasFattening) row.push('');

    for (const f of report.fishFormulas) row.push(cleanSup(f));
    if (report.hasFish) row.push('');

    for (const f of report.duckFormulas) row.push(cleanSup(f));
    if (report.hasDuck) row.push('');

    for (const f of report.petsFormulas) row.push(cleanSup(f));
    if (report.hasPets) row.push('');

    row.push(effectiveDate);
    return row;
  };

  const handleDownloadSadatExcel = async () => {
    await exportSadatControlStyledExcel({
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
    });
  };

  const handleDownloadCsv = () => {
    const baseName = fileName.replace(/\.(xlsx|pdf|csv)$/i, '');
    const sheetData = buildExportSheetData();
    const withSupledger: SheetData = {
      ...sheetData,
      headers: buildSupledgerRow(),
      rows: [sheetData.headers, ...sheetData.rows],
    };
    exportToCsv(withSupledger, `${baseName}_كنترول_السادات.csv`);
  };

  const handleCopy = async () => {
    const sheetData = buildExportSheetData();
    const withSupledger: SheetData = {
      ...sheetData,
      headers: buildSupledgerRow(),
      rows: [sheetData.headers, ...sheetData.rows],
    };
    const ok = await copyTableToClipboard(withSupledger);
    if (ok) {
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    }
  };

  const formatSupledgerDisplay = (val: string) => {
    return (val || '').replace(/,/g, '').trim();
  };

  const renderBatchCell = (f: SadatFormulaColumn) => {
    const batchVal = batchCodeOverrides[f.key] ?? f.productCode ?? '';
    const isEditing = editingBatchKey === f.key;
    return (
      <th
        key={`batch_${f.key}`}
        onClick={() => {
          if (!isEditing) {
            setEditingBatchKey(f.key);
            setEditingBatchVal(batchVal);
          }
        }}
        className="px-1.5 py-1 border-b-2 border-l border-slate-900 text-center font-mono tabular-nums bg-white cursor-pointer hover:bg-amber-50 whitespace-nowrap"
        title="انقر لإدخال أو تعديل رقم الـ supledger الخاص بهذه التركيبة"
      >
        {isEditing ? (
          <input
            type="text"
            autoFocus
            value={editingBatchVal}
            onChange={(e) => setEditingBatchVal(e.target.value)}
            onBlur={() => saveBatchCode(f.key)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') saveBatchCode(f.key);
              if (e.key === 'Escape') setEditingBatchKey(null);
            }}
            onClick={(e) => e.stopPropagation()}
            placeholder="11336"
            className="w-16 px-1 py-0.5 text-center font-mono text-[11px] border border-indigo-500 rounded bg-white"
          />
        ) : batchVal ? (
          formatSupledgerDisplay(batchVal)
        ) : (
          <span className="text-slate-300 font-normal text-[10px]">+ supledger</span>
        )}
      </th>
    );
  };

  const renderFormulaHeader = (f: SadatFormulaColumn, colorClass: string) => (
    <th
      key={f.key}
      onClick={() => toggleSort({ kind: 'formula', key: f.key })}
      className={`px-1.5 py-2 border-b-2 border-l border-slate-600/80 text-center font-extrabold ${colorClass} min-w-[72px] max-w-[130px] bg-[#002060] cursor-pointer hover:bg-[#002b80]`}
      title={`التركيبة: ${f.label} (${f.itemType}) - انقر للترتيب`}
    >
      <div className="flex items-center justify-center gap-1 leading-tight whitespace-normal break-words">
        <span>{f.label}</span>
        {renderSortBadge({ kind: 'formula', key: f.key })}
      </div>
    </th>
  );

  const displayUnit = useTons && baseIsKg ? 'طن' : report.unit;
  const fixedMetaColCount = showDriveCode ? 3 : 2;
  const totalColumnCount =
    fixedMetaColCount +
    allOrderedFormulas.length +
    (report.hasFattening ? 1 : 0) +
    (report.hasFish ? 1 : 0) +
    (report.hasDuck ? 1 : 0) +
    (report.hasPets ? 1 : 0) +
    1;

  const renderMatrixDataRow = (row: SadatMatrixRow, keyStr: string, isWaterHighlight: boolean) => {
    const baseCellBg = isWaterHighlight ? 'bg-sky-50' : 'bg-white';
    return (
      <tr
        key={keyStr}
        className={`${
          isWaterHighlight ? 'hover:bg-sky-100/70 border-t border-sky-300' : 'hover:bg-slate-100/70'
        } transition-colors`}
      >
        <td
          className={`px-2 py-1 border-b border-l border-slate-400 font-mono text-center font-bold text-slate-900 ${baseCellBg} tabular-nums whitespace-nowrap`}
        >
          {row.code || ''}
        </td>

        {showDriveCode && (
          <td
            className={`px-2 py-1 border-b border-l border-slate-400 font-mono text-center font-extrabold text-indigo-900 ${
              isWaterHighlight ? 'bg-sky-100/70' : 'bg-indigo-50/40'
            } tabular-nums whitespace-nowrap`}
          >
            {row.driveCode || ''}
          </td>
        )}

        <td
          className={`px-2.5 py-1 border-b border-l border-slate-400 font-bold ${
            isWaterHighlight ? 'text-sky-950' : 'text-slate-900'
          } ${baseCellBg} whitespace-nowrap`}
        >
          {getRowDisplayName(row)}
        </td>

        {/* تسمين */}
        {report.fatteningFormulas.map((f) => (
          <td
            key={f.key}
            dir="ltr"
            className={`px-2 py-1 border-b border-l border-slate-400 font-mono tabular-nums text-left font-bold text-slate-950 ${baseCellBg}`}
          >
            {fmt(row.formulaValues[f.key] || 0)}
          </td>
        ))}
        {report.hasFattening && (
          <td
            dir="ltr"
            className="px-2 py-1 border-b border-l border-slate-500 font-mono tabular-nums text-left font-extrabold text-slate-950 bg-[#ffe699]"
          >
            {fmt(row.fatteningTotal)}
          </td>
        )}

        {/* سمك */}
        {report.fishFormulas.map((f) => (
          <td
            key={f.key}
            dir="ltr"
            className={`px-2 py-1 border-b border-l border-slate-400 font-mono tabular-nums text-left font-bold text-slate-950 ${baseCellBg}`}
          >
            {fmt(row.formulaValues[f.key] || 0)}
          </td>
        ))}
        {report.hasFish && (
          <td
            dir="ltr"
            className="px-2 py-1 border-b border-l border-slate-500 font-mono tabular-nums text-left font-extrabold text-slate-950 bg-[#f4b084]"
          >
            {fmt(row.fishTotal)}
          </td>
        )}

        {/* بط */}
        {report.duckFormulas.map((f) => (
          <td
            key={f.key}
            dir="ltr"
            className={`px-2 py-1 border-b border-l border-slate-400 font-mono tabular-nums text-left font-bold text-slate-950 ${baseCellBg}`}
          >
            {fmt(row.formulaValues[f.key] || 0)}
          </td>
        ))}
        {report.hasDuck && (
          <td
            dir="ltr"
            className="px-2 py-1 border-b border-l border-slate-500 font-mono tabular-nums text-left font-extrabold text-slate-950 bg-[#b4c6e7]"
          >
            {fmt(row.duckTotal)}
          </td>
        )}

        {/* أليف */}
        {report.petsFormulas.map((f) => (
          <td
            key={f.key}
            dir="ltr"
            className={`px-2 py-1 border-b border-l border-slate-400 font-mono tabular-nums text-left font-bold text-slate-950 ${baseCellBg}`}
          >
            {fmt(row.formulaValues[f.key] || 0)}
          </td>
        ))}
        {report.hasPets && (
          <td
            dir="ltr"
            className="px-2 py-1 border-b border-l border-slate-500 font-mono tabular-nums text-left font-extrabold text-slate-950 bg-[#e2d4f0]"
          >
            {fmt(row.petsTotal)}
          </td>
        )}

        {/* الإجمالي العام */}
        <td
          dir="ltr"
          className="px-2.5 py-1 border-b border-slate-400 font-mono tabular-nums text-left font-extrabold text-slate-950 bg-[#a9d08e]"
        >
          {row.grandTotal.toFixed(grandTotalDecimals)}
        </td>
      </tr>
    );
  };

  const renderMatrixTotalRow = (
    label: string,
    countVal: number,
    sums: ReturnType<typeof computeSums>,
    isFinalGrandTotal: boolean
  ) => {
    const rowBg = isFinalGrandTotal ? 'bg-[#ffd966]' : 'bg-amber-100';
    const borderTop = isFinalGrandTotal ? 'border-t-2 border-slate-900' : 'border-t-2 border-slate-800';
    return (
      <tr className={`${rowBg} ${borderTop} font-black text-slate-950`}>
        <td className="px-2 py-2 border-b border-l border-slate-600 text-center font-mono">
          {countVal}
        </td>
        {showDriveCode && <td className="px-2 py-2 border-b border-l border-slate-600 text-center" />}
        <td className="px-2.5 py-2 border-b border-l border-slate-600 whitespace-nowrap">{label}</td>

        {/* تسمين */}
        {report.fatteningFormulas.map((f) => (
          <td
            key={`sum_${label}_${f.key}`}
            dir="ltr"
            className="px-2 py-2 border-b border-l border-slate-600 font-mono tabular-nums text-left"
          >
            {fmt(sums.formulaSums[f.key] || 0)}
          </td>
        ))}
        {report.hasFattening && (
          <td
            dir="ltr"
            className="px-2 py-2 border-b border-l border-slate-600 font-mono tabular-nums text-left bg-[#ffd966]"
          >
            {fmt(sums.fatteningTotal)}
          </td>
        )}

        {/* سمك */}
        {report.fishFormulas.map((f) => (
          <td
            key={`sum_${label}_${f.key}`}
            dir="ltr"
            className="px-2 py-2 border-b border-l border-slate-600 font-mono tabular-nums text-left"
          >
            {fmt(sums.formulaSums[f.key] || 0)}
          </td>
        ))}
        {report.hasFish && (
          <td
            dir="ltr"
            className="px-2 py-2 border-b border-l border-slate-600 font-mono tabular-nums text-left bg-[#e89968]"
          >
            {fmt(sums.fishTotal)}
          </td>
        )}

        {/* بط */}
        {report.duckFormulas.map((f) => (
          <td
            key={`sum_${label}_${f.key}`}
            dir="ltr"
            className="px-2 py-2 border-b border-l border-slate-600 font-mono tabular-nums text-left"
          >
            {fmt(sums.formulaSums[f.key] || 0)}
          </td>
        ))}
        {report.hasDuck && (
          <td
            dir="ltr"
            className="px-2 py-2 border-b border-l border-slate-600 font-mono tabular-nums text-left bg-[#8faadc]"
          >
            {fmt(sums.duckTotal)}
          </td>
        )}

        {/* أليف */}
        {report.petsFormulas.map((f) => (
          <td
            key={`sum_${label}_${f.key}`}
            dir="ltr"
            className="px-2 py-2 border-b border-l border-slate-600 font-mono tabular-nums text-left"
          >
            {fmt(sums.formulaSums[f.key] || 0)}
          </td>
        ))}
        {report.hasPets && (
          <td
            dir="ltr"
            className="px-2 py-2 border-b border-l border-slate-600 font-mono tabular-nums text-left bg-[#c9b2e3]"
          >
            {fmt(sums.petsTotal)}
          </td>
        )}

        {/* الإجمالي العام */}
        <td
          dir="ltr"
          className="px-2.5 py-2 border-b border-slate-600 font-mono tabular-nums text-left bg-[#8ec26d]"
        >
          {sums.grandTotal.toFixed(grandTotalDecimals)}
        </td>
      </tr>
    );
  };

  return (
    <div
      className={
        isFullscreen
          ? 'fixed inset-0 z-50 bg-white overflow-y-auto p-3 space-y-3 animate-in fade-in duration-150'
          : 'space-y-3 animate-in fade-in duration-200'
      }
      dir="rtl"
    >
      {/* Top Action & Filter Controls */}
      <div className="px-2 sm:px-3 pt-3 flex flex-wrap items-center justify-between gap-2 print:hidden">
        <div className="relative flex-1 min-w-[200px] max-w-xs">
          <Search className="w-4 h-4 text-slate-400 absolute right-2.5 top-2.5" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="بحث بكود الصنف أو اسمه..."
            className="w-full pr-8 pl-3 py-1.5 border border-slate-200 rounded-lg text-xs focus:outline-hidden focus:border-[#002060] bg-white"
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
          {report.hasActualData && report.hasTargetData && (
            <div className="inline-flex rounded-lg border border-slate-200 bg-slate-100 p-0.5">
              <button
                onClick={() => setMetric('actual')}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  metric === 'actual' ? 'bg-white text-slate-900 shadow-2xs' : 'text-slate-600'
                }`}
              >
                الفعلي (Actual)
              </button>
              <button
                onClick={() => setMetric('target')}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  metric === 'target' ? 'bg-white text-slate-900 shadow-2xs' : 'text-slate-600'
                }`}
              >
                المستهدف (Target)
              </button>
            </div>
          )}

          {baseIsKg && (
            <div className="inline-flex rounded-lg border border-slate-200 bg-slate-100 p-0.5">
              <button
                onClick={() => setUseTons(true)}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  useTons ? 'bg-[#002060] text-amber-300 shadow-2xs' : 'text-slate-600'
                }`}
              >
                بالطن (÷ 1000)
              </button>
              <button
                onClick={() => setUseTons(false)}
                className={`px-2.5 py-1 rounded-md font-bold transition-colors cursor-pointer ${
                  !useTons ? 'bg-[#002060] text-amber-300 shadow-2xs' : 'text-slate-600'
                }`}
              >
                كجم
              </button>
            </div>
          )}

          <label className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white font-semibold text-slate-700 cursor-pointer select-none">
            <input
              type="checkbox"
              checked={useBilingualNames}
              onChange={(e) => setUseBilingualNames(e.target.checked)}
              className="w-3.5 h-3.5 accent-[#002060]"
            />
            <span>عربي + إنجليزي</span>
          </label>

          <label className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white font-semibold text-slate-700 cursor-pointer select-none">
            <input
              type="checkbox"
              checked={separateWater}
              onChange={(e) => setSeparateWater(e.target.checked)}
              className="w-3.5 h-3.5 accent-[#002060]"
            />
            <span>فصل الماء أسفل الإجمالي</span>
          </label>

          <button
            onClick={() => setShowDriveCode((v) => !v)}
            className={`px-2.5 py-1.5 rounded-lg border font-semibold transition-colors flex items-center gap-1 cursor-pointer ${
              showDriveCode
                ? 'bg-indigo-50 text-indigo-900 border-indigo-300'
                : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
            }`}
          >
            <Columns3 className="w-3.5 h-3.5" />
            <span>كود الدريف</span>
          </button>

          <button
            onClick={handleCopy}
            className="px-2.5 py-1.5 font-semibold bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 rounded-lg transition-colors flex items-center gap-1 cursor-pointer"
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
            className="px-2.5 py-1.5 font-semibold bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 rounded-lg transition-colors flex items-center gap-1 cursor-pointer"
          >
            <Download className="w-3.5 h-3.5 text-slate-500" />
            <span>CSV</span>
          </button>

          <button
            onClick={() => window.print()}
            className="px-2.5 py-1.5 font-semibold bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 rounded-lg transition-colors flex items-center gap-1 cursor-pointer"
          >
            <Printer className="w-3.5 h-3.5 text-slate-500" />
            <span>طباعة</span>
          </button>

          <button
            onClick={() => setIsFullscreen((v) => !v)}
            className={`px-2.5 py-1.5 font-semibold border rounded-lg transition-colors flex items-center gap-1 cursor-pointer ${
              isFullscreen
                ? 'bg-[#002060] text-white border-[#002060]'
                : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
            }`}
            title={isFullscreen ? 'تصغير من وضع ملء الشاشة' : 'عرض بملء الشاشة الكاملة'}
          >
            {isFullscreen ? (
              <>
                <Minimize2 className="w-3.5 h-3.5" />
                <span>خروج من ملء الشاشة</span>
              </>
            ) : (
              <>
                <Maximize2 className="w-3.5 h-3.5 text-slate-500" />
                <span>ملء الشاشة</span>
              </>
            )}
          </button>

          {isFullscreen && onIncreaseFont && onDecreaseFont && (
            <div
              className="inline-flex items-center bg-white border border-slate-200 rounded-lg p-0.5"
              title="تكبير أو تصغير حجم الخط"
            >
              <button
                type="button"
                onClick={onDecreaseFont}
                disabled={!canDecreaseFont}
                className="px-2 py-1 font-bold text-slate-700 hover:bg-slate-100 rounded-md disabled:opacity-35 disabled:cursor-not-allowed cursor-pointer flex items-center gap-1"
                title="تصغير الخط"
              >
                <ZoomOut className="w-3.5 h-3.5 text-slate-600" />
                <span>A-</span>
              </button>
              <button
                type="button"
                onClick={onResetFont}
                className="px-2 py-1 font-mono font-bold text-slate-800 hover:bg-slate-50 border-x border-slate-200 cursor-pointer tabular-nums"
                title="إرجاع حجم الخط الافتراضي"
              >
                {fontPx ?? 13}px
              </button>
              <button
                type="button"
                onClick={onIncreaseFont}
                disabled={!canIncreaseFont}
                className="px-2 py-1 font-bold text-slate-700 hover:bg-slate-100 rounded-md disabled:opacity-35 disabled:cursor-not-allowed cursor-pointer flex items-center gap-1"
                title="تكبير الخط"
              >
                <ZoomIn className="w-3.5 h-3.5 text-emerald-700" />
                <span>A+</span>
              </button>
            </div>
          )}

          <button
            onClick={handleDownloadSadatExcel}
            className="px-3 py-1.5 font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg shadow-xs transition-all flex items-center gap-1.5 cursor-pointer"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>تصدير تقرير السادات (.xlsx)</span>
          </button>
        </div>
      </div>

      {/* Quick Category Totals Overview Bar (always visible without horizontal scrolling) */}
      <div className="px-2 sm:px-3 flex flex-wrap items-center gap-2 text-xs">
        {report.hasFattening && (
          <div className="px-3 py-1.5 rounded-lg border border-amber-300 bg-[#fff2cc] text-slate-900 font-bold flex items-center gap-1.5">
            <span>🌾 إجمالي التسمين:</span>
            <span className="font-mono tabular-nums" dir="ltr">
              {sumsWithWater.fatteningTotal.toFixed(grandTotalDecimals)}
            </span>
            <span className="text-[11px] text-slate-600">{displayUnit}</span>
          </div>
        )}
        {report.hasFish && (
          <div className="px-3 py-1.5 rounded-lg border border-orange-300 bg-[#fce4d6] text-slate-900 font-bold flex items-center gap-1.5">
            <span>🐟 إجمالي سمك:</span>
            <span className="font-mono tabular-nums" dir="ltr">
              {sumsWithWater.fishTotal.toFixed(grandTotalDecimals)}
            </span>
            <span className="text-[11px] text-slate-600">{displayUnit}</span>
          </div>
        )}
        {report.hasDuck && (
          <div className="px-3 py-1.5 rounded-lg border border-sky-300 bg-[#d9e1f2] text-slate-900 font-bold flex items-center gap-1.5">
            <span>🦆 إجمالي البط:</span>
            <span className="font-mono tabular-nums" dir="ltr">
              {sumsWithWater.duckTotal.toFixed(grandTotalDecimals)}
            </span>
            <span className="text-[11px] text-slate-600">{displayUnit}</span>
          </div>
        )}
        {report.hasPets && (
          <div className="px-3 py-1.5 rounded-lg border border-purple-300 bg-[#ede4f7] text-slate-900 font-bold flex items-center gap-1.5">
            <span>🐾 إجمالي الأليف:</span>
            <span className="font-mono tabular-nums" dir="ltr">
              {sumsWithWater.petsTotal.toFixed(grandTotalDecimals)}
            </span>
            <span className="text-[11px] text-slate-600">{displayUnit}</span>
          </div>
        )}
        <div className="px-3 py-1.5 rounded-lg border border-emerald-400 bg-[#d9ead3] text-emerald-950 font-extrabold flex items-center gap-1.5">
          <span>⭐ الإجمالي العام:</span>
          <span className="font-mono tabular-nums" dir="ltr">
            {sumsWithWater.grandTotal.toFixed(grandTotalDecimals)}
          </span>
          <span className="text-[11px] text-emerald-800">{displayUnit}</span>
        </div>
      </div>

      {/* Main Excel-Style Sadat Control Sheet */}
      <div className="px-2 sm:px-3 pb-4">
        <div className="border-2 border-slate-900 rounded-md overflow-hidden shadow-sm bg-white">
          {/* Top Yellow Header Banner matching image_9da537b6.png */}
          <div className="bg-[#ffff00] border-b-2 border-slate-900 px-3 py-2 flex flex-wrap items-center justify-between gap-3">
            <div className="bg-white border border-slate-500 px-3.5 py-1 text-center text-[11px] font-extrabold text-slate-900 leading-snug shadow-2xs">
              <div>شركة الدقهلية للدواجن</div>
              <div>مصنع أعلاف السادات</div>
              <div>إدارة المخازن</div>
              <div>مخزن الخامات</div>
            </div>

            <div className="text-center flex-1 min-w-[200px]">
              {editingMetaField === 'title' ? (
                <input
                  type="text"
                  autoFocus
                  value={editingMetaVal}
                  onChange={(e) => setEditingMetaVal(e.target.value)}
                  onBlur={() => saveMetaField('title')}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') saveMetaField('title');
                    if (e.key === 'Escape') setEditingMetaField(null);
                  }}
                  className="px-3 py-1 text-lg sm:text-2xl font-black text-center bg-white border border-slate-800 rounded w-full max-w-md"
                />
              ) : (
                <h2
                  onClick={() => {
                    setEditingMetaField('title');
                    setEditingMetaVal(meta.title);
                  }}
                  title="انقر لتعديل عنوان التقرير"
                  className="text-xl sm:text-2xl font-black text-slate-950 tracking-tight cursor-pointer hover:underline decoration-slate-700/50"
                >
                  {meta.title}
                </h2>
              )}
              <div className="text-[11px] font-bold text-slate-800 mt-0.5">
                الوحدة: {displayUnit} · عدد التركيبات: {allOrderedFormulas.length} · عدد الخامات: {displayedRows.length}
              </div>
            </div>

            <div className="bg-white border border-slate-400 px-4 py-2 flex items-center gap-2 shadow-2xs">
              <img src="/logo.png" alt="الدقهلية للدواجن - Dakahlia Poultry" className="h-10 w-auto object-contain" />
            </div>
          </div>

          {/* Scrollable Matrix Table */}
          <div className={`overflow-x-auto ${isFullscreen ? 'max-h-[calc(100vh-210px)]' : 'max-h-[78vh]'}`}>
            <table className={`w-full text-right border-separate border-spacing-0 ${tableFontClass}`}>
              <thead className="sticky top-0 z-30 select-none">
                {/* Sub-header White Reference Row (supledger | batch codes | date) */}
                <tr className="bg-white text-slate-900 font-bold text-[0.9em]">
                  <th
                    colSpan={fixedMetaColCount}
                    onClick={() => {
                      setEditingMetaField('subLabel');
                      setEditingMetaVal(meta.subLabel);
                    }}
                    className="px-2.5 py-1 border-b-2 border-l border-slate-900 text-center font-mono bg-white cursor-pointer hover:bg-slate-50"
                    title="انقر لتعديل التسمية"
                  >
                    {editingMetaField === 'subLabel' ? (
                      <input
                        type="text"
                        autoFocus
                        value={editingMetaVal}
                        onChange={(e) => setEditingMetaVal(e.target.value)}
                        onBlur={() => saveMetaField('subLabel')}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') saveMetaField('subLabel');
                          if (e.key === 'Escape') setEditingMetaField(null);
                        }}
                        onClick={(e) => e.stopPropagation()}
                        className="w-28 px-1.5 py-0.5 text-center border border-slate-500 rounded bg-white"
                      />
                    ) : (
                      meta.subLabel || 'supledger'
                    )}
                  </th>

                  {report.fatteningFormulas.map(renderBatchCell)}
                  {report.hasFattening && (
                    <th className="px-1.5 py-1 border-b-2 border-l border-slate-900 bg-white" />
                  )}

                  {report.fishFormulas.map(renderBatchCell)}
                  {report.hasFish && (
                    <th className="px-1.5 py-1 border-b-2 border-l border-slate-900 bg-white" />
                  )}

                  {report.duckFormulas.map(renderBatchCell)}
                  {report.hasDuck && (
                    <th className="px-1.5 py-1 border-b-2 border-l border-slate-900 bg-white" />
                  )}

                  {report.petsFormulas.map(renderBatchCell)}
                  {report.hasPets && (
                    <th className="px-1.5 py-1 border-b-2 border-l border-slate-900 bg-white" />
                  )}

                  {/* Editable Report Date above الإجمالي العام */}
                  <th
                    onClick={() => {
                      setEditingMetaField('date');
                      setEditingMetaVal(effectiveDate);
                    }}
                    className="px-2 py-1 border-b-2 border-slate-900 text-center font-mono tabular-nums bg-white cursor-pointer hover:bg-slate-50 whitespace-nowrap"
                    title="انقر لتعديل تاريخ التقرير"
                  >
                    {editingMetaField === 'date' ? (
                      <input
                        type="text"
                        autoFocus
                        value={editingMetaVal}
                        onChange={(e) => setEditingMetaVal(e.target.value)}
                        onBlur={() => saveMetaField('date')}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') saveMetaField('date');
                          if (e.key === 'Escape') setEditingMetaField(null);
                        }}
                        onClick={(e) => e.stopPropagation()}
                        className="w-24 px-1 py-0.5 text-center border border-slate-500 rounded bg-white"
                      />
                    ) : (
                      effectiveDate
                    )}
                  </th>
                </tr>

                {/* Main Dark Navy Header Row */}
                <tr className="bg-[#002060] text-white">
                  <th
                    onClick={() => toggleSort({ kind: 'code' })}
                    className="px-2 py-2 border-b-2 border-l border-slate-600/80 text-center font-extrabold bg-[#002060] whitespace-nowrap cursor-pointer hover:bg-[#002b80]"
                  >
                    <div className="flex items-center justify-center gap-1">
                      <span>كود الصنف</span>
                      {renderSortBadge({ kind: 'code' })}
                    </div>
                  </th>

                  {showDriveCode && (
                    <th
                      onClick={() => toggleSort({ kind: 'driveCode' })}
                      className="px-2 py-2 border-b-2 border-l border-slate-600/80 text-center font-extrabold text-indigo-200 bg-[#002060] whitespace-nowrap cursor-pointer hover:bg-[#002b80]"
                    >
                      <div className="flex items-center justify-center gap-1">
                        <span>كود الدريف</span>
                        {renderSortBadge({ kind: 'driveCode' })}
                      </div>
                    </th>
                  )}

                  <th
                    onClick={() => toggleSort({ kind: 'name' })}
                    className="px-2 py-2 border-b-2 border-l border-slate-600/80 font-extrabold min-w-[155px] bg-[#002060] whitespace-nowrap cursor-pointer hover:bg-[#002b80]"
                  >
                    <div className="flex items-center justify-between gap-2">
                      <span>اسم الصنف</span>
                      {renderSortBadge({ kind: 'name' })}
                    </div>
                  </th>

                  {/* تسمين formulas + اجمالي التسمين */}
                  {report.fatteningFormulas.map((f) => renderFormulaHeader(f, 'text-[#ffff00]'))}
                  {report.hasFattening && (
                    <th
                      onClick={() => toggleSort({ kind: 'fatteningTotal' })}
                      className="px-2 py-2 border-b-2 border-l border-slate-600/80 text-center font-extrabold text-[#ffff00] bg-[#002060] whitespace-nowrap cursor-pointer hover:bg-[#002b80]"
                    >
                      <div className="flex items-center justify-center gap-1">
                        <span>اجمالي التسمين</span>
                        {renderSortBadge({ kind: 'fatteningTotal' })}
                      </div>
                    </th>
                  )}

                  {/* سمك formulas + اجمالي سمك */}
                  {report.fishFormulas.map((f) => renderFormulaHeader(f, 'text-sky-200'))}
                  {report.hasFish && (
                    <th
                      onClick={() => toggleSort({ kind: 'fishTotal' })}
                      className="px-2 py-2 border-b-2 border-l border-slate-600/80 text-center font-extrabold text-amber-200 bg-[#002060] whitespace-nowrap cursor-pointer hover:bg-[#002b80]"
                    >
                      <div className="flex items-center justify-center gap-1">
                        <span>اجمالي سمك</span>
                        {renderSortBadge({ kind: 'fishTotal' })}
                      </div>
                    </th>
                  )}

                  {/* بط formulas + اجمالي البط (if exists) */}
                  {report.duckFormulas.map((f) => renderFormulaHeader(f, 'text-teal-200'))}
                  {report.hasDuck && (
                    <th
                      onClick={() => toggleSort({ kind: 'duckTotal' })}
                      className="px-2 py-2 border-b-2 border-l border-slate-600/80 text-center font-extrabold text-teal-200 bg-[#002060] whitespace-nowrap cursor-pointer hover:bg-[#002b80]"
                    >
                      <div className="flex items-center justify-center gap-1">
                        <span>اجمالي البط</span>
                        {renderSortBadge({ kind: 'duckTotal' })}
                      </div>
                    </th>
                  )}

                  {/* أليف formulas + اجمالي الاليف (if exists) */}
                  {report.petsFormulas.map((f) => renderFormulaHeader(f, 'text-purple-200'))}
                  {report.hasPets && (
                    <th
                      onClick={() => toggleSort({ kind: 'petsTotal' })}
                      className="px-2 py-2 border-b-2 border-l border-slate-600/80 text-center font-extrabold text-purple-200 bg-[#002060] whitespace-nowrap cursor-pointer hover:bg-[#002b80]"
                    >
                      <div className="flex items-center justify-center gap-1">
                        <span>اجمالي الاليف</span>
                        {renderSortBadge({ kind: 'petsTotal' })}
                      </div>
                    </th>
                  )}

                  {/* الإجمالي العام */}
                  <th
                    onClick={() => toggleSort({ kind: 'grandTotal' })}
                    className="px-2.5 py-2 border-b-2 border-slate-900 text-center font-extrabold text-emerald-200 bg-[#002060] whitespace-nowrap cursor-pointer hover:bg-[#002b80]"
                  >
                    <div className="flex items-center justify-center gap-1">
                      <span>الإجمالي العام</span>
                      {renderSortBadge({ kind: 'grandTotal' })}
                    </div>
                  </th>
                </tr>
              </thead>

              <tbody>
                {displayedRows.length === 0 ? (
                  <tr>
                    <td
                      colSpan={totalColumnCount}
                      className="py-10 text-center text-slate-500 text-sm"
                    >
                      لا توجد أصناف مطابقة للبحث.
                    </td>
                  </tr>
                ) : (
                  <>
                    {regularRows.map((row, idx) =>
                      renderMatrixDataRow(row, `${row.code}_${row.name}_${idx}`, false)
                    )}

                    {waterRows.length > 0 &&
                      renderMatrixTotalRow(
                        `الإجمالي (بدون الماء) (${displayUnit})`,
                        regularRows.length,
                        sumsWithoutWater,
                        false
                      )}

                    {waterRows.map((row, idx) =>
                      renderMatrixDataRow(row, `water_${row.code}_${idx}`, true)
                    )}

                    {renderMatrixTotalRow(
                      waterRows.length > 0
                        ? `الإجمالي الكلي (بالماء) (${displayUnit})`
                        : `الإجمالي الكلي (${displayUnit})`,
                      displayedRows.length,
                      sumsWithWater,
                      true
                    )}
                  </>
                )}
              </tbody>
            </table>
          </div>
        </div>

        <p className="mt-2 text-[11px] text-slate-500 print:hidden">
          💡 يتم سحب رقم الـ <strong>supledger</strong> تلقائياً من ترويسة كل تركيبة في ملف الكنترول (مثل <code>11336</code>)، ويمكنك أيضاً النقر على أي خلية في صف <code>supledger</code> لتعديله أو إدخاله يدوياً. كما يظهر كل نوع بعمود إجماليه المستقل (<strong>اجمالي التسمين</strong>، <strong>اجمالي سمك</strong>، <strong>اجمالي البط</strong>، <strong>اجمالي الاليف</strong> عند وجودها).
        </p>
      </div>
    </div>
  );
};
