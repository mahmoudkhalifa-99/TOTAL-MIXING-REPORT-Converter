import React from 'react';
import { X, CheckCircle2, FileSpreadsheet, Sparkles, HelpCircle } from 'lucide-react';

interface HelpModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const HelpModal: React.FC<HelpModalProps> = ({ isOpen, onClose }) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-xs">
      <div className="bg-white rounded-xl shadow-xl border border-slate-200 w-full max-w-xl max-h-[calc(var(--app-vh,1vh)*85)] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        <div className="p-4 sm:p-5 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <HelpCircle className="w-5 h-5 text-emerald-600" />
            <h3 className="text-base font-bold text-slate-900">
              دليل الاستخدام والأسئلة الشائعة
            </h3>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-600 rounded-md hover:bg-slate-100 transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="p-5 sm:p-6 overflow-y-auto space-y-5 text-xs sm:text-sm text-slate-700 leading-relaxed">
          <div className="space-y-2">
            <h4 className="font-bold text-slate-900 text-sm flex items-center gap-2">
              <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center text-xs font-bold">1</span>
              <span>ما هي أنواع ملفات الـ PDF المدعومة؟</span>
            </h4>
            <p className="text-slate-600 pr-8">
              يدعم التطبيق جميع أنواع ملفات الـ PDF الرقمية، والملفات الممسوحة ضوئياً (Scanned PDF)، وحتى لقطات الشاشة وصور الجداول (PNG / JPG).
            </p>
          </div>

          <div className="space-y-2">
            <h4 className="font-bold text-slate-900 text-sm flex items-center gap-2">
              <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center text-xs font-bold">2</span>
              <span>كيف يتعامل التطبيق مع اللغة العربية واتجاه اليمين لليسار (RTL)؟</span>
            </h4>
            <p className="text-slate-600 pr-8">
              التطبيق مجهز بنظام ذكي يكتشف اللغة العربية تلقائياً، ويضبط أوراق العمل في ملف الإكسل الناتج لتكون باتجاه اليمين لليسار (Right-to-Left)، مع الحفاظ التام على تسلسل الحروف والأرقام دون أي عكس أو تشويه.
            </p>
          </div>

          <div className="space-y-2">
            <h4 className="font-bold text-slate-900 text-sm flex items-center gap-2">
              <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center text-xs font-bold">3</span>
              <span>هل الأرقام في ملف الإكسل قابلة للعمليات الحسابية والمعادلات؟</span>
            </h4>
            <p className="text-slate-600 pr-8">
              نعم، يقوم التطبيق بتنظيف الأرقام والأسعار من الرموز الزائدة وتحويلها إلى قيم رقمية حقيقية (Numeric Cells)، مما يتيح لك استخدام دوال مثل <code className="bg-slate-100 px-1 rounded font-mono">=SUM()</code> و <code className="bg-slate-100 px-1 rounded font-mono">=AVERAGE()</code> مباشرة في برنامج إكسل.
            </p>
          </div>

          <div className="space-y-2">
            <h4 className="font-bold text-slate-900 text-sm flex items-center gap-2">
              <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center text-xs font-bold">4</span>
              <span>هل يمكنني تعديل البيانات قبل تنزيل الملف؟</span>
            </h4>
            <p className="text-slate-600 pr-8">
              بالتأكيد! بعد اكتمال التحويل، يظهر جدول تفاعلي كامل يمكنك النقر نقراً مزدوجاً على أي خلية لتعديلها، أو تعديل أسماء الأعمدة، أو إضافة صفوف جديدة قبل الضغط على زر تحميل ملف الإكسل.
            </p>
          </div>

          <div className="space-y-2">
            <h4 className="font-bold text-slate-900 text-sm flex items-center gap-2">
              <span className="w-6 h-6 rounded-full bg-indigo-100 text-indigo-700 flex items-center justify-center text-xs font-bold">5</span>
              <span>كيف يعمل عمود "نوع الصنف" والتصنيف التلقائي للتركيبات؟</span>
            </h4>
            <div className="text-slate-600 pr-8 space-y-1.5">
              <p>يتم إضافة عمود "نوع الصنف" وتصنيف كل صف آلياً وفقاً لاسم التركيبة:</p>
              <ul className="list-disc list-inside space-y-1 text-xs">
                <li><strong className="text-sky-700">🐟 سمك:</strong> إذا كان اسم التركيبة يحتوي على كلمة <code className="bg-slate-100 px-1 py-0.5 rounded font-mono font-bold">Tilapia</code>.</li>
                <li><strong className="text-amber-700">🌾 تسمين:</strong> إذا كان اسم التركيبة يحتوي على أي من الرموز <code className="bg-slate-100 px-1 py-0.5 rounded font-mono font-bold">GR</code> أو <code className="bg-slate-100 px-1 py-0.5 rounded font-mono font-bold">FN</code> أو <code className="bg-slate-100 px-1 py-0.5 rounded font-mono font-bold">ST</code> أو <code className="bg-slate-100 px-1 py-0.5 rounded font-mono font-bold">SS</code>.</li>
                <li><strong className="text-teal-700">🦆 بط:</strong> إذا كان اسم التركيبة يحتوي على كلمة <code className="bg-slate-100 px-1 py-0.5 rounded font-mono font-bold">Duck</code>.</li>
                <li><strong className="text-purple-700">🐾 اليف:</strong> لجميع التركيبات والأصناف الأخرى.</li>
              </ul>
              <p className="text-[11px] text-slate-500 mt-1">يمكنك أيضاً تصفية الجدول بنقرة واحدة حسب نوع الصنف، أو تعديل أي تصنيف مباشرة من القائمة المنسدلة للخلية.</p>
            </div>
          </div>
        </div>

        <div className="p-4 bg-slate-50 border-t border-slate-200 text-right">
          <button
            onClick={onClose}
            className="px-4 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer"
          >
            فهمت ذلك، إغلاق
          </button>
        </div>
      </div>
    </div>
  );
};
