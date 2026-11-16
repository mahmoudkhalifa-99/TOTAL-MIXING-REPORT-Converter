import React, { useRef, useState, useEffect } from 'react';
import {
  UploadCloud,
  FileText,
  SlidersHorizontal,
  ChevronDown,
  ChevronUp,
  X,
  Play,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  Clock,
  Sparkles,
  Scale,
} from 'lucide-react';
import { ConversionOptions } from '../types/converter';
import { SAMPLE_DOCUMENTS, SampleDocItem } from '../utils/sampleDocuments';

interface DropzoneProps {
  onFileSelected: (file: File, options: ConversionOptions) => void;
  onSelectSample: (sample: SampleDocItem) => void;
  isLoading: boolean;
  loadingStep: string;
}

export const Dropzone: React.FC<DropzoneProps> = ({
  onFileSelected,
  onSelectSample,
  isLoading,
  loadingStep,
}) => {
  const [dragActive, setDragActive] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [showSamples, setShowSamples] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [elapsedSeconds, setElapsedSeconds] = useState(0);

  useEffect(() => {
    let interval: any;
    if (isLoading) {
      setElapsedSeconds(0);
      interval = setInterval(() => {
        setElapsedSeconds((prev) => prev + 1);
      }, 1000);
    } else {
      setElapsedSeconds(0);
    }
    return () => clearInterval(interval);
  }, [isLoading]);

  // Conversion options
  const [options, setOptions] = useState<ConversionOptions>({
    extractionMode: 'auto',
    customInstructions: '',
    parseNumbers: true,
    autoDetectRTL: true,
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  };

  const validateAndSetFile = (file: File) => {
    setErrorMessage(null);
    const lowerName = file.name.toLowerCase();
    const isPdf = file.type === 'application/pdf' || lowerName.endsWith('.pdf');
    const isSheet =
      lowerName.endsWith('.xlsx') ||
      lowerName.endsWith('.xls') ||
      lowerName.endsWith('.csv');

    if (!isPdf && !isSheet) {
      setErrorMessage('الملف غير مدعوم. يرجى اختيار ملف PDF نصي أو ملف Excel / CSV.');
      return;
    }

    // 25MB safety limit
    if (file.size > 25 * 1024 * 1024) {
      setErrorMessage('حجم الملف كبير جداً. الحد الأقصى للملف هو 25 ميجابايت.');
      return;
    }

    setSelectedFile(file);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      validateAndSetFile(e.dataTransfer.files[0]);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      validateAndSetFile(e.target.files[0]);
    }
  };

  const handleStartConversion = () => {
    if (!selectedFile) return;
    onFileSelected(selectedFile, options);
  };

  const formatFileSize = (bytes: number) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  return (
    <div id="converter" className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
      <div className="p-6 md:p-8 space-y-6">
        {/* Drag and Drop Box */}
        <div
          onDragEnter={handleDrag}
          onDragLeave={handleDrag}
          onDragOver={handleDrag}
          onDrop={handleDrop}
          onClick={() => !selectedFile && !isLoading && fileInputRef.current?.click()}
          className={`relative border-2 border-dashed rounded-xl p-8 sm:p-12 text-center transition-all ${
            dragActive
              ? 'border-emerald-500 bg-emerald-50/60'
              : selectedFile
              ? 'border-emerald-300 bg-emerald-50/30'
              : 'border-slate-300 hover:border-emerald-400 hover:bg-slate-50/60 cursor-pointer'
          }`}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".pdf,.xlsx,.xls,.csv"
            onChange={handleChange}
            className="hidden"
            disabled={isLoading}
          />

          {!selectedFile ? (
            <div className="space-y-4">
              <div className="w-16 h-16 mx-auto rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center">
                <UploadCloud className="w-8 h-8" />
              </div>
              <div>
                <p className="text-base font-semibold text-slate-800">
                  اسحب وأفلت ملف الـ PDF هنا، أو{' '}
                  <span className="text-emerald-700 underline underline-offset-4 font-bold">
                    تصفح من جهازك
                  </span>
                </p>
                <div className="mt-2 flex items-center justify-center gap-2 text-xs text-slate-500 flex-wrap">
                  <span>تحويل محلي فوري بدون API أو ذكاء اصطناعي</span>
                  <span aria-hidden="true">·</span>
                  <span>يدعم PDF النصي و Excel/CSV حتى 25 ميجابايت</span>
                  <span aria-hidden="true">·</span>
                  <span>دعم كامل للجداول متعددة الصفحات</span>
                </div>
              </div>
            </div>
          ) : (
            <div className="space-y-4" onClick={(e) => e.stopPropagation()}>
              <div className="flex items-center justify-center gap-3">
                <div className="w-12 h-12 rounded-lg bg-emerald-600 text-white flex items-center justify-center">
                  <FileText className="w-6 h-6" />
                </div>
                <div className="text-right">
                  <h4 className="text-base font-bold text-slate-900 truncate max-w-sm sm:max-w-md">
                    {selectedFile.name}
                  </h4>
                  <div className="flex items-center gap-2 text-xs text-slate-500 mt-0.5">
                    <span className="font-mono tabular-nums">{formatFileSize(selectedFile.size)}</span>
                    <span aria-hidden="true">·</span>
                    <span>جاهز للتحويل</span>
                  </div>
                </div>
                {!isLoading && (
                  <button
                    onClick={() => setSelectedFile(null)}
                    className="p-1.5 text-slate-400 hover:text-slate-600 rounded-md hover:bg-slate-200 transition-colors mr-3"
                    title="إزالة الملف"
                  >
                    <X className="w-5 h-5" />
                  </button>
                )}
              </div>

              {/* Conversion Button */}
              {!isLoading ? (
                <div className="pt-2 flex flex-wrap items-center justify-center gap-3">
                  <button
                    onClick={handleStartConversion}
                    className="px-6 py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-sm rounded-lg shadow-sm hover:shadow-md transition-all flex items-center gap-2 cursor-pointer"
                  >
                    <FileSpreadsheet className="w-5 h-5" />
                    <span>تحويل المستند إلى Excel الآن</span>
                  </button>
                  <button
                    onClick={() => fileInputRef.current?.click()}
                    className="px-4 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold text-sm rounded-lg transition-colors cursor-pointer"
                  >
                    اختيار ملف آخر
                  </button>
                </div>
              ) : (
                <div className="pt-4 max-w-md mx-auto space-y-3">
                  <div className="flex items-center justify-between text-xs font-semibold text-emerald-800">
                    <div className="flex items-center gap-2">
                      <div className="w-3.5 h-3.5 rounded-full border-2 border-emerald-600 border-t-transparent animate-spin" />
                      <span>{loadingStep || 'جاري استخراج الجداول...'}</span>
                    </div>
                    <div className="flex items-center gap-1 font-mono tabular-nums text-emerald-900 bg-emerald-100 px-2 py-0.5 rounded text-[11px]">
                      <Clock className="w-3 h-3" />
                      <span>{elapsedSeconds} ثانية</span>
                    </div>
                  </div>
                  <div className="w-full bg-slate-200 rounded-full h-2 overflow-hidden">
                    <div
                      className="bg-emerald-600 h-2 rounded-full transition-all duration-300"
                      style={{
                        width: `${Math.min(25 + elapsedSeconds * 12, 95)}%`,
                      }}
                    />
                  </div>
                  <p className="text-xs text-slate-500 text-center">
                    {elapsedSeconds < 4
                      ? 'قراءة الملف والتعرف على بنية الجداول...'
                      : elapsedSeconds < 8
                      ? 'استخراج الأعمدة والبيانات والصفوف بدقة فائقة...'
                      : 'تنسيق الخلايا وأوراق العمل الحسابية للإكسل...'}
                  </p>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Error notification if any */}
        {errorMessage && (
          <div className="p-4 bg-rose-50 border border-rose-200 rounded-lg flex items-start gap-3 text-rose-800 text-sm">
            <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
            <div className="flex-1">
              <span className="font-semibold">تنبيه: </span>
              {errorMessage}
            </div>
            <button
              onClick={() => setErrorMessage(null)}
              className="text-rose-500 hover:text-rose-700"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* Control Total Reconciliation Helper Note */}
        <div className="p-3 bg-gradient-to-r from-indigo-50 via-sky-50 to-emerald-50 border border-indigo-200/80 rounded-xl flex flex-wrap items-center justify-between gap-3 text-xs text-indigo-950">
          <div className="flex items-center gap-2">
            <Scale className="w-4 h-4 text-indigo-600 shrink-0" />
            <span>
              <strong>تريد معرفة إجمالي الكنترول ومطابقته مع الـ PDF وإجمالي كل صنف؟</strong> يقوم النظام بحساب الكنترول العام وتفصيل كميات ونسب (تسمين، سمك، بط، أليف) مع المطابقة الفورية.
            </span>
          </div>
          <button
            type="button"
            onClick={() => onSelectSample(SAMPLE_DOCUMENTS[0])}
            className="text-xs font-bold text-indigo-700 hover:text-indigo-950 underline cursor-pointer shrink-0"
          >
            معاينة بالنموذج الآن ←
          </button>
        </div>

        {/* Toggleable Action Buttons (Samples & Advanced Options - hidden until clicked) */}
        <div className="pt-2 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center flex-wrap gap-2.5">
            {/* Quick Button: Control Total & PDF Matching Feed Sample */}
            <button
              type="button"
              onClick={() => onSelectSample(SAMPLE_DOCUMENTS[0])}
              className="px-3.5 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-2 cursor-pointer border bg-indigo-50 text-indigo-900 border-indigo-200 hover:bg-indigo-100 shadow-2xs"
              title="تحميل جدول خلطات الأعلاف مباشرة لمعاينة إجمالي الكنترول ومطابقة الـ PDF وتفصيل كل صنف (تسمين، سمك، بط، أليف)"
            >
              <Scale className="w-3.5 h-3.5 text-indigo-600" />
              <span>معاينة إجمالي الكنترول ومطابقة الأصناف</span>
            </button>

            {/* Button 1: Toggle Sample Documents */}
            <button
              type="button"
              onClick={() => {
                setShowSamples(!showSamples);
                if (!showSamples) setShowAdvanced(false);
              }}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-2 cursor-pointer border ${
                showSamples
                  ? 'bg-emerald-600 text-white border-emerald-600 shadow-2xs'
                  : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50 hover:border-slate-300'
              }`}
              title="عرض النماذج التجريبية الجاهزة لتجربة التحويل بنقرة واحدة"
            >
              <Sparkles className={`w-3.5 h-3.5 ${showSamples ? 'text-white' : 'text-emerald-600'}`} />
              <span>نماذج تجريبية جاهزة</span>
              <span className={`text-[11px] font-mono px-1.5 py-0.2 rounded ${showSamples ? 'bg-emerald-700 text-white' : 'bg-slate-100 text-slate-600'}`}>
                {SAMPLE_DOCUMENTS.length}
              </span>
              {showSamples ? (
                <ChevronUp className="w-3.5 h-3.5" />
              ) : (
                <ChevronDown className="w-3.5 h-3.5 text-slate-400" />
              )}
            </button>

            {/* Button 2: Toggle Advanced Options */}
            <button
              type="button"
              onClick={() => {
                setShowAdvanced(!showAdvanced);
                if (!showAdvanced) setShowSamples(false);
              }}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-2 cursor-pointer border ${
                showAdvanced
                  ? 'bg-slate-800 text-white border-slate-800 shadow-2xs'
                  : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50 hover:border-slate-300'
              }`}
              title="تعديل نمط الاستخراج وإضافة تعليمات مخصصة"
            >
              <SlidersHorizontal className={`w-3.5 h-3.5 ${showAdvanced ? 'text-white' : 'text-slate-500'}`} />
              <span>خيارات استخراج متقدمة</span>
              {showAdvanced ? (
                <ChevronUp className="w-3.5 h-3.5" />
              ) : (
                <ChevronDown className="w-3.5 h-3.5 text-slate-400" />
              )}
            </button>
          </div>
        </div>

        {/* Expandable Advanced Options (ONLY visible when toggled) */}
        {showAdvanced && (
          <div className="border border-slate-200 rounded-xl overflow-hidden p-4 sm:p-5 bg-white space-y-4 text-sm animate-in fade-in duration-200">
            <div className="flex items-center justify-between pb-2 border-b border-slate-100">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <SlidersHorizontal className="w-4 h-4 text-slate-600" />
                <span>خيارات الاستخراج المتقدمة والتعليمات المخصصة:</span>
              </span>
              <button
                type="button"
                onClick={() => setShowAdvanced(false)}
                className="text-slate-400 hover:text-slate-600 text-xs font-bold p-1 cursor-pointer"
              >
                ✕ إغلاق
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1.5">
                  نمط الاستخراج الرئيسي
                </label>
                <select
                  value={options.extractionMode}
                  onChange={(e) =>
                    setOptions({
                      ...options,
                      extractionMode: e.target.value as any,
                    })
                  }
                  className="w-full px-3 py-2 text-xs sm:text-sm border border-slate-300 rounded-lg focus:outline-hidden focus:border-emerald-500 bg-white"
                >
                  <option value="auto">كشف تلقائي ذكي (مستحسن لكافة المستندات)</option>
                  <option value="mixing_reports">تقارير خلطات ومصانع أعلاف (استخراج 100% لكافة التركيبات بلا استثناء)</option>
                  <option value="tables_only">استخراج الجداول الصريحة فقط وتجاهل النصوص الجانبية</option>
                  <option value="financial">فواتير وكشوف حسابات مالية (تركيز عالي على الأرقام والضرائب)</option>
                  <option value="raw">استخراج موسع شامل لكافة الكيانات المنظمة</option>
                </select>

              </div>

              <div className="space-y-2 pt-1">
                <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-slate-700">
                  <input
                    type="checkbox"
                    checked={options.parseNumbers}
                    onChange={(e) =>
                      setOptions({ ...options, parseNumbers: e.target.checked })
                    }
                    className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500"
                  />
                  <span>تحويل الأرقام والأسعار تلقائياً إلى قيم عددية لتمكين معادلات إكسل</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-slate-700">
                  <input
                    type="checkbox"
                    checked={options.autoDetectRTL}
                    onChange={(e) =>
                      setOptions({ ...options, autoDetectRTL: e.target.checked })
                    }
                    className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500"
                  />
                  <span>تهيئة اتجاه ورقة العمل من اليمين لليسار (RTL) تلقائياً للمحتوى العربي</span>
                </label>
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1.5">
                ملاحظات إضافية للاستخراج المحلي:
              </label>
              <textarea
                value={options.customInstructions}
                onChange={(e) =>
                  setOptions({ ...options, customInstructions: e.target.value })
                }
                placeholder="مثال: استخرج فقط جدول المبيعات دون بيانات العميل..."
                rows={2}
                className="w-full px-3 py-2 text-xs sm:text-sm border border-slate-300 rounded-lg focus:outline-hidden focus:border-emerald-500"
              />
            </div>
          </div>
        )}

        {/* Expandable Quick Sample Document Selector (ONLY visible when toggled) */}
        {showSamples && (
          <div id="samples" className="border border-emerald-200 bg-emerald-50/20 rounded-xl p-4 sm:p-5 space-y-3 animate-in fade-in duration-200">
            <div className="flex items-center justify-between pb-2 border-b border-emerald-100">
              <div className="flex items-center gap-2">
                <Sparkles className="w-4 h-4 text-emerald-600" />
                <span className="text-xs font-bold text-slate-800">
                  نماذج جاهزة للمعاينة والتجربة الفورية (اختر نموذجاً للتحميل المباشر):
                </span>
              </div>
              <button
                type="button"
                onClick={() => setShowSamples(false)}
                className="text-slate-400 hover:text-slate-600 text-xs font-bold p-1 cursor-pointer"
              >
                ✕ إغلاق
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3 pt-1">
              {SAMPLE_DOCUMENTS.map((sample) => (
                <button
                  key={sample.id}
                  onClick={() => onSelectSample(sample)}
                  disabled={isLoading}
                  className="p-3 text-right bg-white hover:bg-emerald-50 border border-slate-200 hover:border-emerald-300 rounded-lg transition-all group cursor-pointer disabled:opacity-50 shadow-2xs"
                >
                  <div className="flex items-center justify-between mb-1.5">
                    <span className="text-xs font-semibold text-emerald-800">
                      {sample.category}
                    </span>
                    <Play className="w-3.5 h-3.5 text-slate-400 group-hover:text-emerald-600 transition-colors" />
                  </div>
                  <h5 className="text-xs font-bold text-slate-900 group-hover:text-emerald-900 line-clamp-1">
                    {sample.name}
                  </h5>
                  <p className="mt-1 text-[11px] text-slate-500 line-clamp-2 leading-relaxed">
                    {sample.description}
                  </p>
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
