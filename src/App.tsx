import React, { useState, useEffect } from 'react';
import { Header } from './components/Header';
import { Dropzone } from './components/Dropzone';
import { ExcelGrid } from './components/ExcelGrid';
import { FeaturesSection } from './components/FeaturesSection';
import { FeaturesModal } from './components/FeaturesModal';
import { HistoryModal } from './components/HistoryModal';
import { HelpModal } from './components/HelpModal';
import {
  ExtractedDocument,
  ConversionOptions,
  HistoryItem,
} from './types/converter';
import { SampleDocItem, SAMPLE_DOCUMENTS } from './utils/sampleDocuments';
import { propagateFormulaNames } from './utils/formulaPropagator';
import { ensureItemTypeColumn } from './utils/itemClassifier';
import { consolidateDuplicateCodeColumns } from './utils/codeSplitter';
import {
  extractDocumentFromPdfLocal,
  extractDocumentFromSpreadsheetLocal,
} from './utils/localPdfExtractor';
import { PdfToExcelConverter } from './modules/pdfToExcelConverter/PdfToExcelConverter';
import {
  AlertCircle,
  X,
  CheckCircle2,
  RefreshCw,
  ShieldCheck,
  ChevronDown,
  ChevronUp,
  FilePlus,
} from 'lucide-react';

const STORAGE_KEY = 'pdf_to_excel_history_v1';

export default function App() {
  const [currentDocument, setCurrentDocument] = useState<ExtractedDocument | null>(null);
  const [currentFileName, setCurrentFileName] = useState<string>('');
  const [pdfDataUrl, setPdfDataUrl] = useState<string | undefined>(undefined);
  const [showDropzone, setShowDropzone] = useState<boolean>(true);

  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [loadingStep, setLoadingStep] = useState<string>('');
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [successToast, setSuccessToast] = useState<string | null>(null);

  const [history, setHistory] = useState<HistoryItem[]>([]);
  const [isHistoryOpen, setIsHistoryOpen] = useState<boolean>(false);
  const [isHelpOpen, setIsHelpOpen] = useState<boolean>(false);
  const [isFeaturesModalOpen, setIsFeaturesModalOpen] = useState<boolean>(false);
  const [showFeatures, setShowFeatures] = useState<boolean>(false);
  const [activeSection, setActiveSection] = useState<'main' | 'pdf_to_excel'>('main');


  // Load history on initial mount
  useEffect(() => {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) {
        const parsed: HistoryItem[] = JSON.parse(stored);
        const cleaned = parsed.map((item) => ({
          ...item,
          data: {
            ...item.data,
            sheets: Array.isArray(item.data?.sheets)
              ? item.data.sheets.map((s) => consolidateDuplicateCodeColumns(s))
              : [],
          },
        }));
        setHistory(cleaned);
      }
    } catch (e) {
      console.error('Failed to load history from localStorage', e);
    }
  }, []);

  // Save history on changes (gracefully handles localStorage QuotaExceededError for large PDFs)
  const saveHistory = (newItems: HistoryItem[]) => {
    setHistory(newItems);
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(newItems));
    } catch {
      try {
        const stripped = newItems.map((item) => ({ ...item, pdfDataUrl: undefined }));
        localStorage.setItem(STORAGE_KEY, JSON.stringify(stripped));
      } catch (e2) {
        console.error('Failed to save history to localStorage', e2);
      }
    }
  };

  const [lastAttempt, setLastAttempt] = useState<{ file: File; options: ConversionOptions } | null>(null);

  const resolveMimeType = (file: File) => {
    const rawType = (file.type || '').toLowerCase();
    if (rawType === 'image/jpg') return 'image/jpeg';
    if (rawType) return rawType;
    const name = file.name.toLowerCase();
    if (name.endsWith('.png')) return 'image/png';
    if (name.endsWith('.jpg') || name.endsWith('.jpeg')) return 'image/jpeg';
    if (name.endsWith('.webp')) return 'image/webp';
    return 'application/pdf';
  };

  // Convert uploaded file 100% locally without AI or external API
  const handleFileSelected = async (file: File, options: ConversionOptions) => {
    setLastAttempt({ file, options });
    setCurrentDocument(null);
    setCurrentFileName('');
    setPdfDataUrl(undefined);
    setErrorMessage(null);
    setSuccessToast(null);
    setIsLoading(true);
    setLoadingStep('جاري قراءة وتجهيز الملف محلياً...');

    try {
      const lowerName = file.name.toLowerCase();
      const resolvedMime = resolveMimeType(file);
      const isPdf = resolvedMime === 'application/pdf' || lowerName.endsWith('.pdf');
      const isSpreadsheet =
        lowerName.endsWith('.xlsx') ||
        lowerName.endsWith('.xls') ||
        lowerName.endsWith('.csv');

      if (!isPdf && !isSpreadsheet) {
        throw new Error(
          'التحويل المحلي بدون API يدعم ملفات الـ PDF النصية أو ملفات Excel / CSV مباشرة (لأن الصور الممسوحة تحتاج محرك OCR أو API).'
        );
      }

      // Generate Data URL for PDF split-screen preview when reasonable size
      let dataUrl: string | undefined;
      if (isPdf && file.size < 15 * 1024 * 1024) {
        dataUrl = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result as string);
          reader.onerror = reject;
          reader.readAsDataURL(file);
        });
        setPdfDataUrl(dataUrl);
      }
      setCurrentFileName(file.name);

      const extracted = isPdf
        ? await extractDocumentFromPdfLocal(file, options, (step) => setLoadingStep(step))
        : await extractDocumentFromSpreadsheetLocal(file, options, (step) => setLoadingStep(step));

      if (!extracted || !Array.isArray(extracted.sheets) || extracted.sheets.length === 0) {
        throw new Error('لم يتم العثور على جداول أو بيانات منظمة داخل هذا الملف.');
      }

      setCurrentDocument(extracted);
      setShowDropzone(false); // Hide dropzone after conversion

      // Save to history (keep pdfDataUrl only when reasonably sized so localStorage stays healthy)
      const totalRows = extracted.sheets.reduce(
        (acc, s) => acc + (s.rows?.length || 0),
        0
      );
      const historyItem: HistoryItem = {
        id: 'hist_' + Date.now(),
        fileName: file.name,
        fileSize: file.size,
        timestamp: Date.now(),
        documentTitle: extracted.documentTitle || file.name,
        sheetCount: extracted.sheets.length,
        rowCount: totalRows,
        data: extracted,
        pdfDataUrl: isPdf && dataUrl && dataUrl.length < 1_500_000 ? dataUrl : undefined,
      };

      const updatedHistory = [historyItem, ...history.slice(0, 19)];
      saveHistory(updatedHistory);

      setSuccessToast(`تم تحويل الملف محلياً بنجاح بدون API! تم استخراج ${extracted.sheets.length} أوراق عمل بإجمالي ${totalRows} صف.`);
      setTimeout(() => setSuccessToast(null), 5000);

      // Scroll to result view
      setTimeout(() => {
        const resultEl = document.getElementById('result-section');
        if (resultEl) {
          resultEl.scrollIntoView({ behavior: 'smooth' });
        }
      }, 200);
    } catch (err: any) {
      const rawMsg = String(err?.message || '');
      console.warn('Local conversion warning:', rawMsg || err);
      setErrorMessage(
        rawMsg || 'حدث خطأ أثناء قراءة الملف محلياً. تأكد من أن ملف الـ PDF صالح ويحتوي على نصوص وجداول.'
      );
    } finally {
      setIsLoading(false);
      setLoadingStep('');
    }
  };

  // Load sample document
  const handleSelectSample = (sample: SampleDocItem) => {
    setCurrentDocument(sample.data);
    setCurrentFileName(`${sample.name}.pdf`);
    setPdfDataUrl(undefined);
    setErrorMessage(null);
    setShowDropzone(false);

    setSuccessToast(`تم تحميل نموذج "${sample.name}". يمكنك استعراضه وتعديله وتنزيله كملف Excel.`);
    setTimeout(() => setSuccessToast(null), 4000);

    setTimeout(() => {
      const resultEl = document.getElementById('result-section');
      if (resultEl) {
        resultEl.scrollIntoView({ behavior: 'smooth' });
      }
    }, 150);
  };

  // Open item from history
  const handleSelectHistoryItem = (item: HistoryItem) => {
    const sanitizedSheets = item.data.sheets.map((s) => {
      const consolidated = consolidateDuplicateCodeColumns(s);
      const formulaRes = propagateFormulaNames(consolidated, item.data);
      return ensureItemTypeColumn(formulaRes.sheet, item.data).sheet;
    });
    setActiveSection('main');
    setCurrentDocument({ ...item.data, sheets: sanitizedSheets });
    setCurrentFileName(item.fileName);
    setPdfDataUrl(item.pdfDataUrl);
    setErrorMessage(null);
    setShowDropzone(false);
  };

  const handleClearHistory = () => {
    saveHistory([]);
  };

  const handleDeleteHistoryItem = (id: string) => {
    saveHistory(history.filter((h) => h.id !== id));
  };

  const handleStartNewConversion = () => {
    setActiveSection('main');
    setCurrentDocument(null);
    setCurrentFileName('');
    setPdfDataUrl(undefined);
    setErrorMessage(null);
    setSuccessToast(null);
    setShowDropzone(true);
    setTimeout(() => {
      document.getElementById('converter')?.scrollIntoView({ behavior: 'smooth' });
    }, 50);
  };

  const handleClearCurrentData = () => {
    setCurrentDocument(null);
    setCurrentFileName('');
    setPdfDataUrl(undefined);
    setLastAttempt(null);
    setErrorMessage(null);
    setShowDropzone(true);

    // Also clear cached history and saved report batch/meta inputs so no stale data remains
    saveHistory([]);
    try {
      localStorage.removeItem('sadat_control_batch_codes_v1');
      localStorage.removeItem('sadat_control_meta_v1');
    } catch {
      // ignore storage errors
    }

    setSuccessToast('تم مسح جميع البيانات والسجل بنجاح. يمكنك رفع ملف جديد الآن.');
    setTimeout(() => setSuccessToast(null), 4000);
    setTimeout(() => {
      document.getElementById('converter')?.scrollIntoView({ behavior: 'smooth' });
    }, 50);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col font-sans" dir="rtl">
      {/* Top Bar */}
      <Header
        onOpenHistory={() => setIsHistoryOpen(true)}
        onOpenHelp={() => setIsHelpOpen(true)}
        onOpenFeatures={() => {
          if (currentDocument) {
            setIsFeaturesModalOpen(true);
          } else {
            setShowFeatures((prev) => {
              const nextVal = !prev;
              if (nextVal) {
                setTimeout(() => {
                  document.getElementById('features')?.scrollIntoView({ behavior: 'smooth' });
                }, 100);
              }
              return nextVal;
            });
          }
        }}
        onNewConversion={handleStartNewConversion}
        onClearData={handleClearCurrentData}
        hasActiveData={Boolean(currentDocument) || history.length > 0}
        activeSection={activeSection}
        onSelectSection={(sec) => setActiveSection(sec)}
        historyCount={history.length}
        fullWidth={Boolean(currentDocument) || activeSection === 'pdf_to_excel'}
      />

      {/* Main Container - Full screen width when viewing extracted spreadsheet */}
      <main
        className={
          currentDocument || activeSection === 'pdf_to_excel'
            ? 'flex-1 w-full max-w-none mx-auto px-2 sm:px-4 py-4 space-y-4'
            : 'flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6'
        }
      >
        {activeSection === 'pdf_to_excel' ? (
          <PdfToExcelConverter />
        ) : (
        <>
        {/* Success Toast */}
        {successToast && (
          <div className="p-4 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center justify-between gap-3 text-emerald-900 text-xs sm:text-sm animate-in fade-in duration-200">
            <div className="flex items-center gap-2">
              <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0" />
              <span className="font-semibold">{successToast}</span>
            </div>
            <button
              onClick={() => setSuccessToast(null)}
              className="text-emerald-700 hover:text-emerald-900"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* Global Error Banner */}
        {errorMessage && (
          <div className="p-4 bg-rose-50 border border-rose-200 rounded-xl flex items-start justify-between gap-3 text-rose-900 text-xs sm:text-sm animate-in fade-in duration-200">
            <div className="flex items-start gap-2.5">
              <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
              <div>
                <h5 className="font-bold">تنبيه أثناء معالجة الملف:</h5>
                <p className="mt-0.5 text-rose-800 leading-relaxed">{errorMessage}</p>
                {lastAttempt && (
                  <button
                    onClick={() => handleFileSelected(lastAttempt.file, lastAttempt.options)}
                    disabled={isLoading}
                    className="mt-3 px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg font-bold text-xs flex items-center gap-2 transition-all shadow-xs cursor-pointer disabled:opacity-50"
                  >
                    <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} />
                    <span>إعادة المحاولة الآن (Retry)</span>
                  </button>
                )}
              </div>
            </div>
            <button
              onClick={() => setErrorMessage(null)}
              className="text-rose-600 hover:text-rose-800 p-1 rounded-md"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* File Dropzone & Converter Controls - Hidden after conversion when document is loaded */}
        {(!currentDocument || showDropzone) && (
          <div id="converter" className="space-y-2 animate-in fade-in duration-200">
            {currentDocument && (
              <div className="flex items-center justify-between pb-1">
                <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                  <FilePlus className="w-4 h-4 text-emerald-600" />
                  <span>تحويل ملف جديد:</span>
                </span>
                <button
                  onClick={() => setShowDropzone(false)}
                  className="text-xs text-slate-500 hover:text-slate-800 font-semibold px-2 py-1 rounded-md hover:bg-slate-100 transition-colors cursor-pointer"
                  title="إخفاء منطقة الرفع والعودة للجدول الحالي"
                >
                  ✕ إخفاء منطقة الرفع والعودة للجدول
                </button>
              </div>
            )}
            <Dropzone
              onFileSelected={handleFileSelected}
              onSelectSample={handleSelectSample}
              isLoading={isLoading}
              loadingStep={loadingStep}
            />
          </div>
        )}

        {/* Result Area (Interactive Excel Spreadsheet) */}
        {currentDocument && (
          <div id="result-section">
            {/* Editable Spreadsheet Grid */}
            <ExcelGrid
              key={`${currentFileName}_${currentDocument.sheets?.[0]?.rows?.length || 0}`}
              document={currentDocument}
              fileName={currentFileName || 'document.xlsx'}
              onUpdateDocument={(updated) => setCurrentDocument(updated)}
              onClearData={handleClearCurrentData}
              pdfDataUrl={pdfDataUrl}
            />
          </div>
        )}

        {/* Features & Guarantees Section - COMPLETELY HIDDEN after conversion when document is active */}
        {!currentDocument && (
          <div id="features" className="pt-4 border-t border-slate-200 flex flex-col items-center">
            <button
              onClick={() => setShowFeatures(!showFeatures)}
              className={`px-5 py-2.5 rounded-xl border text-xs sm:text-sm font-bold transition-all flex items-center gap-2 cursor-pointer shadow-2xs ${
                showFeatures
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50 hover:border-slate-300'
              }`}
            >
              <ShieldCheck className="w-4 h-4 text-emerald-600" />
              <span>{showFeatures ? 'إخفاء مميزات وضمانات النظام' : 'عرض مميزات وضمانات المحول'}</span>
              {showFeatures ? (
                <ChevronUp className="w-4 h-4 text-emerald-700" />
              ) : (
                <ChevronDown className="w-4 h-4 text-slate-400" />
              )}
            </button>

            {showFeatures && (
              <div className="w-full mt-4 animate-in fade-in duration-200">
                <FeaturesSection />
              </div>
            )}
          </div>
        )}
        </>
        )}
      </main>

      {/* History Modal */}
      <HistoryModal
        isOpen={isHistoryOpen}
        onClose={() => setIsHistoryOpen(false)}
        items={history}
        onSelectItem={handleSelectHistoryItem}
        onClearHistory={handleClearHistory}
        onDeleteItem={handleDeleteHistoryItem}
      />

      {/* Help & FAQ Modal */}
      <HelpModal isOpen={isHelpOpen} onClose={() => setIsHelpOpen(false)} />

      {/* Features & Guarantees Modal */}
      <FeaturesModal
        isOpen={isFeaturesModalOpen}
        onClose={() => setIsFeaturesModalOpen(false)}
      />

      {/* Clean Footer */}
      <footer className="bg-white border-t border-slate-200 py-5 mt-8 text-center text-xs text-slate-500">
        <div
          className={`${
            currentDocument ? 'w-full px-4 sm:px-6' : 'max-w-7xl mx-auto px-4 sm:px-6 lg:px-8'
          } flex flex-col sm:flex-row items-center justify-between gap-4`}
        >
          <p>© {new Date().getFullYear()} TOTAL MIXING REPORT Converter - تحويل ومعالجة المستندات والجداول بدقة متناهية.</p>
          <div className="flex items-center gap-4 text-slate-600">
            <button
              onClick={() => setIsHelpOpen(true)}
              className="hover:text-emerald-700 transition-colors cursor-pointer"
            >
              الأسئلة الشائعة
            </button>
            <span aria-hidden="true">·</span>
            <button
              onClick={() => setIsHistoryOpen(true)}
              className="hover:text-emerald-700 transition-colors cursor-pointer"
            >
              سجل التحويلات
            </button>
            <span aria-hidden="true">·</span>
            <span>يدعم معايير Microsoft Excel و Google Sheets</span>
          </div>
        </div>
      </footer>
    </div>
  );
}
