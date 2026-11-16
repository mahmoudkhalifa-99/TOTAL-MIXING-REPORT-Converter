import React from 'react';
import { FileSpreadsheet, History, HelpCircle, Sparkles, Trash2 } from 'lucide-react';

interface HeaderProps {
  onOpenHistory: () => void;
  onOpenHelp: () => void;
  onOpenFeatures?: () => void;
  onNewConversion?: () => void;
  onClearData?: () => void;
  hasActiveData?: boolean;
  activeSection?: 'main' | 'pdf_to_excel';
  onSelectSection?: (section: 'main' | 'pdf_to_excel') => void;
  historyCount: number;
  fullWidth?: boolean;
}

export const Header: React.FC<HeaderProps> = ({
  onOpenHistory,
  onOpenHelp,
  onOpenFeatures,
  onNewConversion,
  onClearData,
  hasActiveData = false,
  activeSection = 'main',
  onSelectSection,
  historyCount,
  fullWidth = false,
}) => {
  return (
    <header className="bg-white border-b border-slate-200 sticky top-0 z-30">
      <div
        className={`${
          fullWidth ? 'w-full px-3 sm:px-6' : 'max-w-7xl mx-auto px-4 sm:px-6 lg:px-8'
        } h-16 flex items-center justify-between gap-8`}
      >
        {/* Zone 1: Brand Wordmark */}
        <div className="flex items-center gap-3 shrink-0">
          <img
            src="/icon-192.png"
            alt="الدقهلية للدواجن"
            className="w-10 h-10 object-contain shrink-0"
          />
          <div className="flex flex-col">
            <span className="text-base font-bold text-slate-900 tracking-tight whitespace-nowrap">
              TOTAL MIXING REPORT Converter
            </span>
            <span className="text-xs text-slate-500 hidden sm:inline">
              تحويل ذكي فوري للجداول والبيانات
            </span>
          </div>
        </div>

        {/* Zone 2: Navigation Links */}
        <nav className="hidden md:flex items-center gap-6 text-sm font-medium text-slate-600">
          {onNewConversion ? (
            <button
              onClick={onNewConversion}
              className="hover:text-emerald-700 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              تحويل المستندات
            </button>
          ) : (
            <a
              href="#converter"
              className="hover:text-emerald-700 transition-colors whitespace-nowrap shrink-0"
            >
              تحويل المستندات
            </a>
          )}
          {onOpenFeatures ? (
            <button
              onClick={onOpenFeatures}
              className="hover:text-emerald-700 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              المميزات والضمانات
            </button>
          ) : (
            <a
              href="#features"
              className="hover:text-emerald-700 transition-colors whitespace-nowrap shrink-0"
            >
              المميزات والضمانات
            </a>
          )}
          <button
            onClick={onOpenHelp}
            className="hover:text-emerald-700 transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer"
          >
            <HelpCircle className="w-4 h-4 text-slate-400" />
            <span>كيف يعمل؟</span>
          </button>
        </nav>

        {/* Zone 3: Primary Action / History */}
        <div className="flex items-center gap-3 shrink-0">
          {onSelectSection && (
            <button
              onClick={() =>
                onSelectSection(activeSection === 'pdf_to_excel' ? 'main' : 'pdf_to_excel')
              }
              className={`px-3.5 py-2 text-xs font-bold rounded-lg border transition-colors flex items-center gap-1.5 cursor-pointer whitespace-nowrap shrink-0 ${
                activeSection === 'pdf_to_excel'
                  ? 'bg-slate-900 text-white border-slate-900'
                  : 'bg-emerald-50 text-emerald-800 border-emerald-200 hover:bg-emerald-100'
              }`}
              title="فتح وحدة PDF to Excel Converter (تحويل مفتوح المصدر بدون API)"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>PDF to Excel Converter</span>
            </button>
          )}

          <button
            onClick={onOpenHistory}
            className="px-3.5 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors flex items-center gap-2 cursor-pointer whitespace-nowrap shrink-0"
            title="سجل التحويلات السابقة"
          >
            <History className="w-4 h-4 text-slate-600" />
            <span className="hidden sm:inline">سجل الملفات</span>
            {historyCount > 0 && (
              <span className="text-xs font-mono font-bold text-emerald-700">
                ({historyCount})
              </span>
            )}
          </button>

          {onClearData && hasActiveData && (
            <button
              type="button"
              onClick={onClearData}
              className="px-3.5 py-2 text-xs font-bold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-lg transition-colors flex items-center gap-1.5 cursor-pointer whitespace-nowrap shrink-0"
              title="مسح البيانات الحالية والسجل المحفوظ"
            >
              <Trash2 className="w-3.5 h-3.5 text-rose-600" />
              <span>مسح البيانات</span>
            </button>
          )}

          {onNewConversion ? (
            <button
              onClick={onNewConversion}
              className="px-4 py-2 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg transition-colors shadow-xs flex items-center gap-1.5 whitespace-nowrap shrink-0 cursor-pointer"
              title="تحويل أو رفع ملف جديد"
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>تحويل ملف جديد</span>
            </button>
          ) : (
            <a
              href="#converter"
              className="px-4 py-2 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg transition-colors shadow-xs flex items-center gap-1.5 whitespace-nowrap shrink-0"
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>تحويل ملف جديد</span>
            </a>
          )}
        </div>
      </div>
    </header>
  );
};
