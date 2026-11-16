import React from 'react';
import { X, History, FileSpreadsheet, Download, Trash2, ArrowRight } from 'lucide-react';
import { HistoryItem } from '../types/converter';
import { exportToExcel } from '../utils/excelExport';

interface HistoryModalProps {
  isOpen: boolean;
  onClose: () => void;
  items: HistoryItem[];
  onSelectItem: (item: HistoryItem) => void;
  onClearHistory: () => void;
  onDeleteItem: (id: string) => void;
}

export const HistoryModal: React.FC<HistoryModalProps> = ({
  isOpen,
  onClose,
  items,
  onSelectItem,
  onClearHistory,
  onDeleteItem,
}) => {
  if (!isOpen) return null;

  const formatDate = (timestamp: number) => {
    return new Date(timestamp).toLocaleDateString('ar-SA', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-xs">
      <div className="bg-white rounded-xl shadow-xl border border-slate-200 w-full max-w-2xl max-h-[calc(var(--app-vh,1vh)*85)] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        {/* Header */}
        <div className="p-4 sm:p-5 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <History className="w-5 h-5 text-emerald-600" />
            <h3 className="text-base font-bold text-slate-900">
              سجل الملفات والتحويلات السابقة
            </h3>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-600 rounded-md hover:bg-slate-100 transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Content List */}
        <div className="p-4 sm:p-5 overflow-y-auto flex-1 space-y-3">
          {items.length === 0 ? (
            <div className="text-center py-12 text-slate-400 space-y-2">
              <FileSpreadsheet className="w-12 h-12 mx-auto stroke-1 text-slate-300" />
              <p className="text-sm font-semibold text-slate-600">
                لا توجد ملفات محفوظة في السجل حتى الآن
              </p>
              <p className="text-xs text-slate-400">
                الملفات التي تقوم بتحويلها ستظهر هنا تلقائياً لسرعة الوصول إليها وتنزيلها
              </p>
            </div>
          ) : (
            items.map((item) => (
              <div
                key={item.id}
                className="p-3.5 border border-slate-200 hover:border-emerald-300 rounded-lg bg-slate-50/50 hover:bg-emerald-50/20 transition-all flex items-center justify-between gap-4"
              >
                <div className="flex items-start gap-3 min-w-0">
                  <div className="w-9 h-9 rounded-md bg-emerald-100 text-emerald-800 flex items-center justify-center shrink-0 mt-0.5">
                    <FileSpreadsheet className="w-4 h-4" />
                  </div>
                  <div className="min-w-0">
                    <h4 className="text-xs sm:text-sm font-bold text-slate-900 truncate">
                      {item.documentTitle || item.fileName}
                    </h4>
                    <div className="flex items-center gap-2 text-[11px] text-slate-500 mt-1">
                      <span>{item.fileName}</span>
                      <span aria-hidden="true">·</span>
                      <span className="font-mono tabular-nums">{item.sheetCount} ورقة</span>
                      <span aria-hidden="true">·</span>
                      <span className="font-mono tabular-nums">{item.rowCount} صف</span>
                      <span aria-hidden="true">·</span>
                      <span>{formatDate(item.timestamp)}</span>
                    </div>
                  </div>
                </div>

                <div className="flex items-center gap-1.5 shrink-0">
                  <button
                    onClick={() => {
                      onSelectItem(item);
                      onClose();
                    }}
                    className="px-3 py-1.5 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-md transition-colors cursor-pointer flex items-center gap-1"
                    title="فتح وعرض في الجدول التفاعلي"
                  >
                    <span>فتح</span>
                    <ArrowRight className="w-3.5 h-3.5 rtl:rotate-180" />
                  </button>

                  <button
                    onClick={() => {
                      const baseName = item.fileName.replace(/\.[^/.]+$/, '');
                      exportToExcel(item.data.sheets, `${baseName}.xlsx`, true);
                    }}
                    className="p-1.5 text-slate-600 hover:text-emerald-700 bg-white border border-slate-200 hover:bg-slate-50 rounded-md transition-colors cursor-pointer"
                    title="تنزيل إكسل مباشر"
                  >
                    <Download className="w-4 h-4" />
                  </button>

                  <button
                    onClick={() => onDeleteItem(item.id)}
                    className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-md transition-colors cursor-pointer"
                    title="حذف من السجل"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>
            ))
          )}
        </div>

        {/* Footer */}
        {items.length > 0 && (
          <div className="p-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs">
            <span className="text-slate-500 font-mono tabular-nums">
              إجمالي العناصر المحفوظة: {items.length}
            </span>
            <button
              onClick={onClearHistory}
              className="text-rose-600 hover:text-rose-800 font-semibold cursor-pointer"
            >
              مسح السجل بالكامل
            </button>
          </div>
        )}
      </div>
    </div>
  );
};
