import React from 'react';
import {
  X,
  FileSpreadsheet,
  Cpu,
  Layers,
  Calculator,
  ShieldCheck,
  Eye,
  Tag,
  CheckCircle2,
} from 'lucide-react';

interface FeaturesModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const FeaturesModal: React.FC<FeaturesModalProps> = ({ isOpen, onClose }) => {
  if (!isOpen) return null;

  const features = [
    {
      icon: Cpu,
      title: 'استخراج ذكي بالذكاء الاصطناعي',
      description:
        'يتعرف تلقائياً على بنية الجداول حتى غير المخططة أو الممسوحة ضوئياً، ويستخرج البيانات بنقاء تام.',
    },
    {
      icon: Layers,
      title: 'دعم الأوراق المتعددة (Multi-Sheet)',
      description:
        'يقوم بتوزيع الجداول المتباينة داخل مستند الـ PDF إلى صفحات إكسل مستقلة ومنظمة داخل نفس المصنف.',
    },
    {
      icon: Calculator,
      title: 'أرقام جاهزة للمعادلات الحسابية',
      description:
        'تحويل تلقائي للأرقام والمبالغ المالية إلى قيم رقمية صحيحة لتشغيل دوال الجمع والضرب والـ SUM فوراً.',
    },
    {
      icon: Eye,
      title: 'معاينة ومقارنة جنباً إلى جنب',
      description:
        'قارن ملف الـ PDF الأصلي مع جدول الإكسل المستخرج للتأكد من مطابقة الأرقام قبل التنزيل.',
    },
    {
      icon: ShieldCheck,
      title: 'تحرير وتعديل فوري في المتصفح',
      description:
        'انقر نقراً مزدوجاً على أي خلية لتصحيح أو تعديل أي قيمة، وأضف صفوفاً جديدة بنقرة واحدة.',
    },
    {
      icon: Tag,
      title: 'تصنيف تلقائي لـ "نوع الصنف"',
      description:
        'إضافة وتصنيف عمود "نوع الصنف" آلياً: سمك (Tilapia)، تسمين (GR - FN - ST - SS)، بط (Duck)، وأليف لباقي الخلطات.',
    },
    {
      icon: FileSpreadsheet,
      title: 'دعم كامل لصيغ Excel و CSV و TSV',
      description:
        'تنزيل مباشر لصيغة .xlsx المتوافقة مع Microsoft Excel و Google Sheets، مع دعم النسخ المباشر للحافظة.',
    },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-xs">
      <div className="bg-white rounded-xl shadow-xl border border-slate-200 w-full max-w-3xl max-h-[calc(var(--app-vh,1vh)*85)] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        <div className="p-4 sm:p-5 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-5 h-5 text-emerald-600" />
            <h3 className="text-base font-bold text-slate-900">
              مميزات وضمانات محول PDF إلى Excel
            </h3>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-600 rounded-md hover:bg-slate-100 transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="p-5 sm:p-6 overflow-y-auto space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {features.map((feat, idx) => {
              const Icon = feat.icon;
              return (
                <div
                  key={idx}
                  className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 hover:bg-white hover:border-emerald-300 transition-all"
                >
                  <div className="w-8 h-8 rounded-lg bg-emerald-100 text-emerald-700 flex items-center justify-center mb-2.5">
                    <Icon className="w-4 h-4" />
                  </div>
                  <h4 className="text-xs sm:text-sm font-bold text-slate-900 mb-1">
                    {feat.title}
                  </h4>
                  <p className="text-xs text-slate-600 leading-relaxed">
                    {feat.description}
                  </p>
                </div>
              );
            })}
          </div>
        </div>

        <div className="p-4 bg-slate-50 border-t border-slate-200 text-right">
          <button
            onClick={onClose}
            className="px-4 py-2 text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg transition-colors cursor-pointer"
          >
            إغلاق
          </button>
        </div>
      </div>
    </div>
  );
};
