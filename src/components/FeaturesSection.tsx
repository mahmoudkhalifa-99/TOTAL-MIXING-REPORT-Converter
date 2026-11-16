import React from 'react';
import {
  FileSpreadsheet,
  Cpu,
  Layers,
  Calculator,
  ShieldCheck,
  Eye,
  Tag,
} from 'lucide-react';

export const FeaturesSection: React.FC = () => {
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
    <section id="features" className="py-12 border-t border-slate-200">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center max-w-2xl mx-auto mb-10">
          <h3 className="text-xl sm:text-2xl font-bold text-slate-900 tracking-tight">
            مميزات محول PDF إلى Excel المتقدم
          </h3>
          <p className="mt-2 text-xs sm:text-sm text-slate-600">
            مصمم لتوفير ساعات من إعادة كتابة البيانات يدوياً وتحويل أي مستند إلى جدول إكسل منسق وجاهز للعمل.
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {features.map((feat, idx) => {
            const Icon = feat.icon;
            return (
              <div
                key={idx}
                className="p-5 rounded-xl border border-slate-200 bg-white hover:border-emerald-200 transition-colors"
              >
                <div className="w-9 h-9 rounded-lg bg-emerald-50 text-emerald-700 flex items-center justify-center mb-3">
                  <Icon className="w-5 h-5" />
                </div>
                <h4 className="text-sm font-bold text-slate-900 mb-1.5">
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
    </section>
  );
};
