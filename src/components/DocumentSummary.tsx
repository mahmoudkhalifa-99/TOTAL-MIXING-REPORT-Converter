import React, { useState } from 'react';
import {
  FileText,
  CheckCircle2,
  AlertTriangle,
  Info,
  Calendar,
  Sparkles,
  ChevronDown,
  ChevronUp,
} from 'lucide-react';
import { ExtractedDocument } from '../types/converter';

interface DocumentSummaryProps {
  document: ExtractedDocument;
}

export const DocumentSummary: React.FC<DocumentSummaryProps> = ({ document }) => {
  const [isExpanded, setIsExpanded] = useState(false);

  if (!document.documentSummary && (!document.notes || document.notes.length === 0)) {
    return null;
  }

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden transition-all">
      <div
        onClick={() => setIsExpanded(!isExpanded)}
        className="p-3.5 sm:px-4 flex items-center justify-between cursor-pointer hover:bg-slate-50/80 transition-colors"
      >
        <div className="flex items-center gap-2">
          <Sparkles className="w-4 h-4 text-emerald-600 shrink-0" />
          <h4 className="text-xs sm:text-sm font-bold text-slate-800">
            ملخص المستند وملاحظات الاستخراج
          </h4>
          {document.detectedLanguage && (
            <span className="text-[11px] text-slate-500 bg-slate-100 px-2 py-0.5 rounded font-mono">
              {document.detectedLanguage}
            </span>
          )}
        </div>
        <div className="flex items-center gap-2 text-xs text-slate-500">
          <span className="text-[11px] font-medium hidden sm:inline">
            {isExpanded ? 'طي الملخص' : 'عرض تفاصيل الملخص والملاحظات'}
          </span>
          {isExpanded ? (
            <ChevronUp className="w-4 h-4 text-slate-500" />
          ) : (
            <ChevronDown className="w-4 h-4 text-slate-500" />
          )}
        </div>
      </div>

      {isExpanded && (
        <div className="px-4 pb-4 pt-1 border-t border-slate-100 space-y-3 animate-in fade-in duration-150">
          {document.documentSummary && (
            <p className="text-xs text-slate-700 leading-relaxed bg-slate-50 p-3 rounded-lg border border-slate-100">
              {document.documentSummary}
            </p>
          )}

          {document.notes && document.notes.length > 0 && (
            <div className="space-y-1.5 pt-1">
              <span className="text-[11px] font-bold text-slate-500">
                ملاحظات وتنبيهات الاستخراج:
              </span>
              <div className="space-y-1">
                {document.notes.map((note, index) => (
                  <div
                    key={index}
                    className="flex items-start gap-2 text-xs text-slate-600"
                  >
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0 mt-0.5" />
                    <span>{note}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

