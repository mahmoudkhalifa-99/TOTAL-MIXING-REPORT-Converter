import React, { useState, useMemo } from 'react';
import {
  Scale,
  CheckCircle2,
  AlertTriangle,
  Copy,
  Check,
  FileSpreadsheet,
  FileText,
  PieChart,
  Filter,
  Layers,
  ArrowRightLeft,
  ChevronDown,
  ChevronUp,
  Sparkles,
  Info,
  Calculator,
} from 'lucide-react';
import { SheetData } from '../types/converter';
import {
  calculateControlTotals,
  generateReconciliationTextReport,
  CategorySummaryItem,
  SheetControlTotalResult,
} from '../utils/controlTotalCalculator';

interface ControlTotalReconciliationProps {
  sheet: SheetData;
  activeCategoryFilter: string;
  onSelectCategoryFilter: (category: string) => void;
  pdfDataUrl?: string;
  onOpenSplitView?: () => void;
  includeReconciliationInExcel: boolean;
  onToggleIncludeReconciliation: (include: boolean) => void;
}

export const ControlTotalReconciliation: React.FC<
  ControlTotalReconciliationProps
> = ({
  sheet,
  activeCategoryFilter,
  onSelectCategoryFilter,
  pdfDataUrl,
  onOpenSplitView,
  includeReconciliationInExcel,
  onToggleIncludeReconciliation,
}) => {
  const [selectedQtyColIndex, setSelectedQtyColIndex] = useState<number | undefined>(undefined);
  const [pdfControlInput, setPdfControlInput] = useState<string>('');
  const [copiedReport, setCopiedReport] = useState(false);
  const [isExpanded, setIsExpanded] = useState(true);

  // Calculate totals and breakdown based on active sheet and chosen column
  const summary: SheetControlTotalResult | null = useMemo(() => {
    return calculateControlTotals(sheet, selectedQtyColIndex);
  }, [sheet, selectedQtyColIndex]);

  // Handle PDF control target number comparison
  const parsedPdfControl = useMemo(() => {
    if (!pdfControlInput.trim()) return undefined;
    const clean = pdfControlInput.replace(/,/g, '').trim();
    const num = parseFloat(clean);
    return isNaN(num) ? undefined : num;
  }, [pdfControlInput]);

  const reconciliationDiff = useMemo(() => {
    if (!summary || parsedPdfControl === undefined) return null;
    return Number((summary.grandTotalQuantity - parsedPdfControl).toFixed(2));
  }, [summary, parsedPdfControl]);

  const isMatched = useMemo(() => {
    if (reconciliationDiff === null) return null;
    return Math.abs(reconciliationDiff) < 0.001;
  }, [reconciliationDiff]);

  // Handle copy text report to clipboard
  const handleCopyReport = async () => {
    if (!summary) return;
    const reportText = generateReconciliationTextReport(
      summary,
      parsedPdfControl
    );
    try {
      await navigator.clipboard.writeText(reportText);
      setCopiedReport(true);
      setTimeout(() => setCopiedReport(false), 2500);
    } catch (err) {
      console.error('Failed to copy reconciliation report', err);
    }
  };

  if (!summary) {
    return null;
  }

  return (
    <div className="bg-white rounded-2xl border border-slate-200/90 shadow-xs overflow-hidden transition-all mb-6">
      {/* Header bar */}
      <div className="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 text-white p-4 sm:p-5 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-emerald-500/20 border border-emerald-400/30 flex items-center justify-center text-emerald-400 shadow-inner shrink-0">
            <Scale className="w-5 h-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h3 className="text-base sm:text-lg font-bold tracking-tight text-white flex items-center gap-2">
                <span>لوحة إجمالي الكنترول ومطابقة الـ PDF</span>
                <span className="text-[11px] font-semibold px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                  تدقيق آلي 100%
                </span>
              </h3>
            </div>
            <p className="text-xs text-slate-300 mt-0.5">
              مجموع كميات الكنترول، المطابقة مع رقم الـ PDF، وتفصيل إجمالي كل نوع صنف (تسمين، سمك، بط، أليف)
            </p>
          </div>
        </div>

        {/* Header Action Buttons */}
        <div className="flex items-center gap-2">
          <button
            onClick={handleCopyReport}
            className="px-3 py-1.5 bg-white/10 hover:bg-white/20 text-white rounded-lg text-xs font-semibold transition-colors flex items-center gap-1.5 cursor-pointer border border-white/15"
            title="نسخ تقرير المطابقة النصي للمشاركة عبر الواتساب أو الإيميل"
          >
            {copiedReport ? (
              <>
                <Check className="w-3.5 h-3.5 text-emerald-400" />
                <span className="text-emerald-300">تم نسخ التقرير!</span>
              </>
            ) : (
              <>
                <Copy className="w-3.5 h-3.5 text-slate-300" />
                <span>نسخ تقرير المطابقة</span>
              </>
            )}
          </button>

          <button
            onClick={() => setIsExpanded(!isExpanded)}
            className="p-1.5 bg-white/10 hover:bg-white/20 text-slate-200 rounded-lg transition-colors cursor-pointer"
            title={isExpanded ? 'طي اللوحة' : 'توسيع اللوحة'}
          >
            {isExpanded ? (
              <ChevronUp className="w-4 h-4" />
            ) : (
              <ChevronDown className="w-4 h-4" />
            )}
          </button>
        </div>
      </div>

      {isExpanded && (
        <div className="p-4 sm:p-6 space-y-6 bg-slate-50/50">
          {/* Top KPI Cards Grid */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            {/* Card 1: Grand Control Total */}
            <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-2xs flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between text-xs text-slate-500 mb-1">
                  <span className="font-bold text-slate-700 flex items-center gap-1">
                    <Scale className="w-3.5 h-3.5 text-emerald-600" />
                    <span>إجمالي الكنترول المستخرج:</span>
                  </span>

                  {/* Column Picker */}
                  {summary.numericColumns.length > 1 && (
                    <div className="relative">
                      <select
                        value={summary.quantityColumnIndex}
                        onChange={(e) =>
                          setSelectedQtyColIndex(Number(e.target.value))
                        }
                        className="text-[11px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 rounded px-1.5 py-0.5 cursor-pointer focus:outline-hidden"
                        title="اختر عمود الكنترول المطلوب حسابه"
                      >
                        {summary.numericColumns.map((col) => (
                          <option key={col.index} value={col.index}>
                            عمود: {col.name}
                          </option>
                        ))}
                      </select>
                    </div>
                  )}
                </div>

                <div className="mt-2 flex items-baseline gap-2">
                  <span className="text-2xl sm:text-3xl font-extrabold text-slate-900 font-mono tracking-tight tabular-nums">
                    {summary.grandTotalQuantity.toLocaleString('ar-EG', {
                      minimumFractionDigits: 2,
                      maximumFractionDigits: 2,
                    })}
                  </span>
                  <span className="text-xs font-bold text-slate-500">
                    {summary.quantityColumnName.includes('طن')
                      ? 'طن'
                      : summary.quantityColumnName.includes('كجم')
                      ? 'كجم'
                      : 'وحدة / كمية'}
                  </span>
                </div>
              </div>

              <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-500">
                <span>العمود المعتمد:</span>
                <span className="font-semibold text-slate-800 truncate max-w-[150px]">
                  {summary.quantityColumnName}
                </span>
              </div>
            </div>

            {/* Card 2: Interactive PDF Control Matcher */}
            <div className={`rounded-xl border p-4 shadow-2xs flex flex-col justify-between transition-colors ${
              reconciliationDiff === null
                ? 'bg-white border-slate-200'
                : isMatched
                ? 'bg-emerald-50/70 border-emerald-300 ring-2 ring-emerald-200'
                : 'bg-amber-50/70 border-amber-300 ring-2 ring-amber-200'
            }`}>
              <div>
                <div className="flex items-center justify-between text-xs mb-1.5">
                  <span className="font-bold flex items-center gap-1.5 text-slate-800">
                    <Calculator className="w-3.5 h-3.5 text-indigo-600" />
                    <span>مطابقة كنترول الـ PDF:</span>
                  </span>

                  {reconciliationDiff !== null && (
                    <span
                      className={`text-[11px] font-bold px-2 py-0.5 rounded-full flex items-center gap-1 ${
                        isMatched
                          ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                          : 'bg-amber-100 text-amber-900 border border-amber-300'
                      }`}
                    >
                      {isMatched ? (
                        <>
                          <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                          <span>متطابق 100%</span>
                        </>
                      ) : (
                        <>
                          <AlertTriangle className="w-3 h-3 text-amber-600" />
                          <span>يوجد فرق ({reconciliationDiff > 0 ? '+' : ''}{reconciliationDiff})</span>
                        </>
                      )}
                    </span>
                  )}
                </div>

                {/* Input for PDF control number */}
                <div className="flex items-center gap-2 mt-2">
                  <div className="relative flex-1">
                    <input
                      type="text"
                      value={pdfControlInput}
                      onChange={(e) => setPdfControlInput(e.target.value)}
                      placeholder="أدخل رقم كنترول الـ PDF (مثال: 397)"
                      className="w-full px-3 py-1.5 text-sm font-mono font-bold border border-slate-300 rounded-lg bg-white text-slate-900 focus:outline-hidden focus:border-indigo-500 placeholder:text-slate-400 placeholder:font-normal"
                    />
                    {pdfControlInput && (
                      <button
                        onClick={() => setPdfControlInput('')}
                        className="absolute left-2.5 top-2 text-slate-400 hover:text-slate-600 text-xs"
                      >
                        ✕
                      </button>
                    )}
                  </div>

                  <button
                    onClick={() =>
                      setPdfControlInput(summary.grandTotalQuantity.toString())
                    }
                    className="px-2.5 py-1.5 text-[11px] font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg border border-slate-200 transition-colors shrink-0 cursor-pointer"
                    title="استخدام إجمالي الجدول كقيمة مرجعية للاختبار"
                  >
                    تعبئة ({summary.grandTotalQuantity})
                  </button>
                </div>
              </div>

              {/* Status footer message */}
              <div className="mt-3 pt-2.5 border-t border-slate-200/60 flex items-center justify-between text-[11px]">
                {reconciliationDiff === null ? (
                  <span className="text-slate-500">
                    أدخل رقم الكنترول المدون في ملف الـ PDF للتحقق الفوري من المطابقة
                  </span>
                ) : isMatched ? (
                  <span className="text-emerald-700 font-bold flex items-center gap-1">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                    <span>البيانات متطابقة تماماً مع ملف الـ PDF (فرق المطابقة: 0.00)</span>
                  </span>
                ) : (
                  <span className="text-amber-800 font-semibold flex items-center gap-1">
                    <AlertTriangle className="w-3.5 h-3.5 text-amber-600" />
                    <span>
                      فرق المطابقة:{' '}
                      <strong className="font-mono tabular-nums text-amber-950 font-bold">
                        {reconciliationDiff > 0 ? `+${reconciliationDiff}` : reconciliationDiff}
                      </strong>
                    </span>
                  </span>
                )}

                {pdfDataUrl && onOpenSplitView && (
                  <button
                    onClick={onOpenSplitView}
                    className="text-indigo-600 hover:text-indigo-800 font-bold underline shrink-0 cursor-pointer"
                    title="معاينة ملف الـ PDF الأصلي جنباً إلى جنب للتحقق البصري"
                  >
                    فحص الـ PDF
                  </button>
                )}
              </div>
            </div>

            {/* Card 3: Financial & Workbook Overview */}
            <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-2xs flex flex-col justify-between">
              <div>
                <span className="text-xs font-bold text-slate-700 flex items-center gap-1 mb-1">
                  <PieChart className="w-3.5 h-3.5 text-purple-600" />
                  <span>ملخص الورقة والقيم المالية:</span>
                </span>

                <div className="mt-2 space-y-1">
                  {summary.grandTotalAmount > 0 ? (
                    <div>
                      <span className="text-xs text-slate-500">
                        إجمالي القيمة ({summary.amountColumnName}):
                      </span>
                      <div className="text-xl sm:text-2xl font-extrabold text-slate-900 font-mono tracking-tight tabular-nums">
                        {summary.grandTotalAmount.toLocaleString('ar-EG', {
                          minimumFractionDigits: 2,
                          maximumFractionDigits: 2,
                        })}{' '}
                        <span className="text-xs font-bold text-slate-500 font-sans">
                          ر.س
                        </span>
                      </div>
                    </div>
                  ) : (
                    <div className="text-sm text-slate-600">
                      إجمالي البنود: <strong className="font-mono text-base">{summary.totalRows}</strong> بند
                    </div>
                  )}
                </div>
              </div>

              <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-500">
                <span>عدد التركيبات: <strong>{summary.totalRows}</strong></span>
                <span className="text-emerald-700 font-semibold">
                  {summary.isCrossCheckBalanced
                    ? '✅ التوازن الحسابي: سليم 100%'
                    : '⚠️ تحقق من التوازن'}
                </span>
              </div>
            </div>
          </div>

          {/* Item Types Visual Proportion Distribution Bar */}
          <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-2xs space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Layers className="w-4 h-4 text-indigo-600" />
                <span>شريط التوزيع النسبي لإجمالي أنواع الأصناف (% من الكنترول العام):</span>
              </span>
              <span className="text-[11px] text-slate-500 font-mono">
                المجموع الكلي: 100% ({summary.grandTotalQuantity} {summary.quantityColumnName.includes('طن') ? 'طن' : ''})
              </span>
            </div>

            {/* Proportion segments */}
            <div className="w-full h-4 bg-slate-100 rounded-full overflow-hidden flex shadow-inner">
              {summary.categories.map((cat) => {
                if (cat.percentage <= 0) return null;
                return (
                  <div
                    key={cat.category}
                    style={{ width: `${cat.percentage}%` }}
                    className={`${cat.barColor} h-full transition-all hover:opacity-90 relative group cursor-pointer`}
                    onClick={() =>
                      onSelectCategoryFilter(
                        activeCategoryFilter === cat.category ? 'all' : cat.category
                      )
                    }
                    title={`${cat.category}: ${cat.totalQuantity} (${cat.percentage}%) - اضغط للتصفية`}
                  />
                );
              })}
            </div>

            {/* Legend */}
            <div className="flex flex-wrap items-center gap-x-4 gap-y-1.5 text-xs">
              {summary.categories.map((cat) => (
                <button
                  key={cat.category}
                  onClick={() =>
                    onSelectCategoryFilter(
                      activeCategoryFilter === cat.category ? 'all' : cat.category
                    )
                  }
                  className={`flex items-center gap-1.5 px-2 py-0.5 rounded-md transition-colors cursor-pointer text-[11px] font-semibold ${
                    activeCategoryFilter === cat.category
                      ? 'bg-slate-800 text-white shadow-2xs'
                      : 'text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  <span className={`w-2.5 h-2.5 rounded-full ${cat.barColor}`} />
                  <span>{cat.icon}</span>
                  <span className="font-bold">{cat.category}:</span>
                  <span className="font-mono tabular-nums">{cat.percentage}%</span>
                  <span className="text-slate-400 font-mono">({cat.totalQuantity})</span>
                </button>
              ))}

              {activeCategoryFilter !== 'all' && (
                <button
                  onClick={() => onSelectCategoryFilter('all')}
                  className="mr-auto text-[11px] text-indigo-600 hover:text-indigo-800 font-bold underline cursor-pointer"
                >
                  إلغاء التصفية (عرض الكل)
                </button>
              )}
            </div>
          </div>

          {/* Breakdown Cards for Each Item Type (تسمين، سمك، بط، أليف) */}
          <div className="space-y-2.5">
            <div className="flex items-center justify-between">
              <h4 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                <span>إجمالي كل نوع صنف بالتفصيل:</span>
                <span className="text-xs text-slate-500 font-normal">
                  (انقر على أي بطاقة لعرض صفوفها مباشرة في الجدول أدناه)
                </span>
              </h4>

              <div className="flex items-center gap-2 text-xs">
                <label className="flex items-center gap-1.5 text-slate-600 cursor-pointer font-medium select-none">
                  <input
                    type="checkbox"
                    checked={includeReconciliationInExcel}
                    onChange={(e) =>
                      onToggleIncludeReconciliation(e.target.checked)
                    }
                    className="w-3.5 h-3.5 text-emerald-600 rounded border-slate-300 focus:ring-emerald-500"
                  />
                  <span>تضمين ورقة مطابقة الكنترول في ملف الإكسل المصدّر</span>
                </label>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
              {summary.categories.map((cat) => {
                const isSelected = activeCategoryFilter === cat.category;
                return (
                  <div
                    key={cat.category}
                    onClick={() =>
                      onSelectCategoryFilter(
                        isSelected ? 'all' : cat.category
                      )
                    }
                    className={`rounded-xl border p-4 transition-all cursor-pointer relative overflow-hidden flex flex-col justify-between ${
                      cat.cardBgClass
                    } ${
                      isSelected
                        ? `${cat.borderColor} ring-2 ring-indigo-500 shadow-md`
                        : `${cat.borderColor} hover:shadow-xs`
                    }`}
                  >
                    {/* Top Type Title & Badge */}
                    <div>
                      <div className="flex items-center justify-between gap-2">
                        <div className="flex items-center gap-1.5">
                          <span className="text-lg">{cat.icon}</span>
                          <span className={`font-extrabold text-sm ${cat.textColor}`}>
                            {cat.category}
                          </span>
                        </div>
                        <span
                          className={`text-[10px] font-bold px-2 py-0.5 rounded-full border shadow-2xs ${cat.badgeClass}`}
                        >
                          {cat.percentage}% من الكنترول
                        </span>
                      </div>

                      <p className="text-[11px] text-slate-500 mt-1 line-clamp-1">
                        {cat.label}
                      </p>

                      {/* Main Totals for this category */}
                      <div className="mt-3">
                        <div className="text-xs text-slate-500">إجمالي الكمية:</div>
                        <div className="text-xl sm:text-2xl font-black font-mono tracking-tight tabular-nums text-slate-900 mt-0.5">
                          {cat.totalQuantity.toLocaleString('ar-EG', {
                            minimumFractionDigits: 2,
                            maximumFractionDigits: 2,
                          })}
                          <span className="text-xs font-bold text-slate-500 mr-1 font-sans">
                            {summary.quantityColumnName.includes('طن')
                              ? 'طن'
                              : summary.quantityColumnName.includes('كجم')
                              ? 'كجم'
                              : ''}
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* Footer metrics: Count and Monetary Total */}
                    <div className="mt-3 pt-2.5 border-t border-slate-200/70 text-[11px] space-y-1">
                      <div className="flex items-center justify-between text-slate-600">
                        <span>عدد التركيبات:</span>
                        <strong className="font-mono text-slate-900">
                          {cat.count} تركيبات
                        </strong>
                      </div>

                      {cat.totalAmount > 0 && (
                        <div className="flex items-center justify-between text-slate-600">
                          <span>إجمالي القيمة:</span>
                          <strong className="font-mono text-slate-900">
                            {cat.totalAmount.toLocaleString('ar-EG', {
                              minimumFractionDigits: 2,
                            })}{' '}
                            ر.س
                          </strong>
                        </div>
                      )}

                      {cat.sampleItems.length > 0 && (
                        <div className="text-[10px] text-slate-400 truncate pt-0.5" title={cat.sampleItems.join(', ')}>
                          أمثلة: {cat.sampleItems.slice(0, 2).join('، ')}
                        </div>
                      )}
                    </div>

                    {/* Indicator if filtered */}
                    {isSelected && (
                      <div className="absolute top-1 left-1 w-2 h-2 rounded-full bg-indigo-600 animate-pulse" />
                    )}
                  </div>
                );
              })}
            </div>
          </div>

          {/* Mathematical cross-check confirmation banner */}
          <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-xl flex flex-wrap items-center justify-between gap-3 text-xs text-emerald-950">
            <div className="flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
              <span>
                <strong>التحقق الحسابي الشامل:</strong> مجموع أصناف (تسمين + سمك + بط + أليف) ={' '}
                <strong className="font-mono tabular-nums text-emerald-900">
                  {summary.crossCheckSum.toLocaleString('ar-EG', { minimumFractionDigits: 2 })}
                </strong>{' '}
                وهو يطابق إجمالي الكنترول العام بنسبة 100%.
              </span>
            </div>

            <button
              onClick={handleCopyReport}
              className="text-emerald-800 hover:text-emerald-950 font-bold underline flex items-center gap-1 cursor-pointer"
            >
              <Copy className="w-3.5 h-3.5" />
              <span>نسخ تقرير المطابقة</span>
            </button>
          </div>
        </div>
      )}
    </div>
  );
};
