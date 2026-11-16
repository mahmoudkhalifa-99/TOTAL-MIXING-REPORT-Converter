import React, { useState } from 'react';
import {
  PieChart,
  Copy,
  Check,
  TrendingUp,
  Scale,
  PlusCircle,
  ChevronDown,
  ChevronUp,
  Sparkles,
  Layers,
  ArrowRight,
} from 'lucide-react';
import {
  OverallTotalsSummary,
  formatWeight,
  generateTotalsReportText,
} from '../utils/categoryTotals';
import { ItemType } from '../utils/itemClassifier';

interface CategoryTotalsBarProps {
  summary: OverallTotalsSummary;
  docTitle?: string;
  activeCategoryFilter: string;
  onSelectCategoryFilter: (category: string) => void;
  onAppendTotalsToSheet?: () => void;
  isTotalsRowAppended?: boolean;
}

export const CategoryTotalsBar: React.FC<CategoryTotalsBarProps> = ({
  summary,
  docTitle = '',
  activeCategoryFilter,
  onSelectCategoryFilter,
  onAppendTotalsToSheet,
  isTotalsRowAppended = false,
}) => {
  const [copied, setCopied] = useState(false);
  const [isExpanded, setIsExpanded] = useState(true);

  const { categories, grandTotal, weightUnit, targetColName, actualColName } = summary;

  const handleCopyReport = async () => {
    const text = generateTotalsReportText(summary, docTitle);
    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    } catch (err) {
      console.error('Failed to copy totals', err);
    }
  };

  const categoryOrder: ItemType[] = ['سمك', 'تسمين', 'بط', 'اليف'];

  return (
    <div className="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 text-white rounded-xl shadow-md border border-slate-700/60 overflow-hidden my-3 animate-in fade-in duration-200">
      {/* Executive Header Bar */}
      <div className="px-4 py-3 sm:px-5 flex flex-wrap items-center justify-between gap-3 border-b border-slate-700/50">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-emerald-500/20 border border-emerald-400/30 flex items-center justify-center text-emerald-400 shrink-0">
            <Scale className="w-4 h-4" />
          </div>
          <div>
            <h4 className="text-sm font-bold tracking-tight text-white flex items-center gap-2">
              <span>إجماليات الأنواع الأربعة (سمك • تسمين • بط • أليف) والإجمالي الكلي</span>
              <span className="text-[10px] bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 px-2 py-0.5 rounded-full font-mono">
                {grandTotal.count} صنف
              </span>
            </h4>
            <p className="text-[11px] text-slate-300 mt-0.5">
              مجموع الأوزان الفعلية والمستهدفة مع احتساب صافي الفروقات ونسب المساهمة
            </p>
          </div>
        </div>

        {/* Action Buttons */}
        <div className="flex items-center gap-2">
          {onAppendTotalsToSheet && (
            <button
              onClick={onAppendTotalsToSheet}
              className={`px-3 py-1.5 text-xs font-semibold rounded-lg border transition-all flex items-center gap-1.5 cursor-pointer shadow-xs ${
                isTotalsRowAppended
                  ? 'bg-emerald-600/30 border-emerald-400/50 text-emerald-200'
                  : 'bg-slate-700/80 hover:bg-slate-700 border-slate-600 text-slate-200'
              }`}
              title="إضافة صفوف مجاميع الأنواع الأربعة في أسفل جدول الإكسل"
            >
              <PlusCircle className="w-3.5 h-3.5 text-emerald-400" />
              <span>{isTotalsRowAppended ? 'تمت إضافة صفوف الإجماليات بالجدول' : 'تضمين الإجماليات بأسفل الجدول'}</span>
            </button>
          )}

          <button
            onClick={handleCopyReport}
            className="px-3 py-1.5 text-xs font-semibold bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg transition-all flex items-center gap-1.5 cursor-pointer shadow-xs"
            title="نسخ تقرير الإجماليات كنص منسق للواتساب أو البريد"
          >
            {copied ? (
              <>
                <Check className="w-3.5 h-3.5 text-white" />
                <span>تم نسخ التقرير!</span>
              </>
            ) : (
              <>
                <Copy className="w-3.5 h-3.5 text-emerald-200" />
                <span>نسخ تقرير الإجماليات</span>
              </>
            )}
          </button>

          <button
            onClick={() => setIsExpanded(!isExpanded)}
            className="p-1.5 text-slate-400 hover:text-white rounded-lg hover:bg-slate-700/50 transition-colors cursor-pointer"
            title={isExpanded ? 'طي لوحة الإجماليات' : 'توسيع لوحة الإجماليات'}
          >
            {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {/* Expanded Totals Cards Grid */}
      {isExpanded && (
        <div className="p-4 sm:p-5 space-y-4">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3.5">
            {/* 1. Fish Card */}
            {categoryOrder.map((catKey) => {
              const cat = categories[catKey];
              const isSelected = activeCategoryFilter === catKey;

              const bgCard =
                catKey === 'سمك'
                  ? 'bg-gradient-to-b from-sky-950/80 to-slate-900 border-sky-500/40'
                  : catKey === 'تسمين'
                  ? 'bg-gradient-to-b from-amber-950/80 to-slate-900 border-amber-500/40'
                  : catKey === 'بط'
                  ? 'bg-gradient-to-b from-teal-950/80 to-slate-900 border-teal-500/40'
                  : 'bg-gradient-to-b from-purple-950/80 to-slate-900 border-purple-500/40';

              const accentColor =
                catKey === 'سمك'
                  ? 'text-sky-300'
                  : catKey === 'تسمين'
                  ? 'text-amber-300'
                  : catKey === 'بط'
                  ? 'text-teal-300'
                  : 'text-purple-300';

              const barColor =
                catKey === 'سمك'
                  ? 'bg-sky-500'
                  : catKey === 'تسمين'
                  ? 'bg-amber-500'
                  : catKey === 'بط'
                  ? 'bg-teal-500'
                  : 'bg-purple-500';

              return (
                <div
                  key={catKey}
                  onClick={() =>
                    onSelectCategoryFilter(isSelected ? 'all' : catKey)
                  }
                  className={`rounded-xl border p-3.5 transition-all cursor-pointer flex flex-col justify-between relative group ${bgCard} ${
                    isSelected
                      ? 'ring-2 ring-emerald-400 shadow-lg scale-[1.02]'
                      : 'hover:border-slate-500 hover:scale-[1.01]'
                  }`}
                  title={`انقر لتصفية الجدول وعرض أصناف ${cat.label} فقط`}
                >
                  {/* Card Header */}
                  <div>
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-xs font-bold flex items-center gap-1.5 text-white">
                        <span className="text-base">{cat.emoji}</span>
                        <span>إجمالي {catKey}</span>
                      </span>
                      <span className="text-[11px] font-mono font-semibold px-2 py-0.5 rounded-full bg-slate-800/90 text-slate-300 border border-slate-700">
                        {cat.count} صنف
                      </span>
                    </div>

                    {/* Actual Total */}
                    <div className="mt-1">
                      <div className="text-[10px] text-slate-400">
                        الفعلي ({actualColName || 'Actual'}):
                      </div>
                      <div className={`text-base sm:text-lg font-bold font-mono tracking-tight ${accentColor}`}>
                        {formatWeight(cat.actualSum, weightUnit)}
                      </div>
                    </div>

                    {/* Target Total */}
                    {cat.hasTargetData && (
                      <div className="mt-1.5 flex items-center justify-between text-[11px] border-t border-slate-700/60 pt-1.5 text-slate-300">
                        <span className="text-slate-400">المستهدف:</span>
                        <span className="font-mono font-semibold">
                          {formatWeight(cat.targetSum, weightUnit)}
                        </span>
                      </div>
                    )}

                    {/* Variance Diff */}
                    {cat.hasTargetData && (
                      <div className="mt-1 flex items-center justify-between text-[11px] font-mono">
                        <span className="text-slate-400">الفرق:</span>
                        <span
                          className={`font-bold ${
                            cat.variance >= 0 ? 'text-emerald-400' : 'text-rose-400'
                          }`}
                        >
                          {cat.variance >= 0 ? '+' : ''}
                          {formatWeight(cat.variance, weightUnit)}
                        </span>
                      </div>
                    )}
                  </div>

                  {/* Progress & Percentage of Grand Total */}
                  <div className="mt-3 pt-2 border-t border-slate-700/50">
                    <div className="flex items-center justify-between text-[10px] text-slate-300 mb-1">
                      <span>نسبة المساهمة:</span>
                      <span className="font-mono font-bold text-white">
                        {cat.actualPercentOfGrand.toFixed(1)}%
                      </span>
                    </div>
                    <div className="w-full bg-slate-800 rounded-full h-1.5 overflow-hidden">
                      <div
                        className={`h-full rounded-full transition-all duration-500 ${barColor}`}
                        style={{ width: `${Math.min(cat.actualPercentOfGrand, 100)}%` }}
                      />
                    </div>
                  </div>

                  {isSelected && (
                    <div className="absolute top-2 left-2 w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
                  )}
                </div>
              );
            })}

            {/* 5. Grand Total Card (الإجمالي الكلي للأربعة) */}
            <div
              onClick={() => onSelectCategoryFilter('all')}
              className={`rounded-xl border p-3.5 transition-all cursor-pointer flex flex-col justify-between bg-gradient-to-b from-emerald-950 via-slate-900 to-slate-900 border-emerald-500/50 ${
                activeCategoryFilter === 'all'
                  ? 'ring-2 ring-emerald-400 shadow-xl'
                  : 'hover:border-emerald-400/80 hover:scale-[1.01]'
              }`}
              title="الإجمالي الكلي لجميع الأصناف الأربعة - انقر لعرض الكل في الجدول"
            >
              <div>
                <div className="flex items-center justify-between mb-2">
                  <span className="text-xs font-bold flex items-center gap-1.5 text-emerald-300">
                    <span className="text-base">⭐</span>
                    <span>الإجمالي الكلي للأربعة</span>
                  </span>
                  <span className="text-[11px] font-mono font-bold px-2 py-0.5 rounded-full bg-emerald-900/60 text-emerald-300 border border-emerald-600/50">
                    {grandTotal.count} صنف (100%)
                  </span>
                </div>

                {/* Grand Actual Total */}
                <div className="mt-1">
                  <div className="text-[10px] text-emerald-400/80">
                    إجمالي الفعلي العام ({actualColName || 'Total'}):
                  </div>
                  <div className="text-base sm:text-xl font-extrabold font-mono text-emerald-300 tracking-tight">
                    {formatWeight(grandTotal.actualSum, weightUnit)}
                  </div>
                </div>

                {/* Grand Target Total */}
                {grandTotal.targetSum > 0 && (
                  <div className="mt-1.5 flex items-center justify-between text-[11px] border-t border-emerald-900/40 pt-1.5 text-slate-300">
                    <span className="text-slate-400">إجمالي المستهدف:</span>
                    <span className="font-mono font-semibold">
                      {formatWeight(grandTotal.targetSum, weightUnit)}
                    </span>
                  </div>
                )}

                {/* Net Variance */}
                {grandTotal.targetSum > 0 && (
                  <div className="mt-1 flex items-center justify-between text-[11px] font-mono">
                    <span className="text-slate-400">صافي الفرق العام:</span>
                    <span
                      className={`font-bold ${
                        grandTotal.variance >= 0 ? 'text-emerald-400' : 'text-rose-400'
                      }`}
                    >
                      {grandTotal.variance >= 0 ? '+' : ''}
                      {formatWeight(grandTotal.variance, weightUnit)}
                    </span>
                  </div>
                )}
              </div>

              {/* Status footer */}
              <div className="mt-3 pt-2 border-t border-emerald-900/40 flex items-center justify-between text-[10px] text-emerald-400">
                <span>تغطية الأصناف:</span>
                <span className="font-bold flex items-center gap-1">
                  <span>مكتمل 100%</span>
                  <Check className="w-3 h-3" />
                </span>
              </div>
            </div>
          </div>

          {/* Quick Filter notice */}
          {activeCategoryFilter !== 'all' && (
            <div className="bg-slate-800/80 border border-slate-700 rounded-lg px-3 py-1.5 flex items-center justify-between text-xs text-slate-300">
              <span className="flex items-center gap-1.5">
                <span>يتم الآن تصفية الجدول لعرض أصناف:</span>
                <strong className="text-emerald-300">"{activeCategoryFilter}"</strong>
                <span className="text-slate-400">({categories[activeCategoryFilter as ItemType]?.count || 0} صنف)</span>
              </span>
              <button
                onClick={() => onSelectCategoryFilter('all')}
                className="text-emerald-400 hover:text-emerald-300 text-xs font-bold underline cursor-pointer"
              >
                إلغاء التصفية وعرض الكل ({grandTotal.count})
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
