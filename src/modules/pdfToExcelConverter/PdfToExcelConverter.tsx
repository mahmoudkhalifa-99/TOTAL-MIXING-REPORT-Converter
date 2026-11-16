import React, { useEffect, useMemo, useRef, useState } from 'react';
import type * as pdfjsLib from 'pdfjs-dist';
import {
  UploadCloud,
  FileText,
  FileSpreadsheet,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  ScanText,
  Layers,
  Server,
  Download,
  Plus,
  Columns3,
  ArrowLeftRight,
  Cpu,
  ShieldCheck,
  RefreshCw,
  Trash2,
} from 'lucide-react';
import {
  CamelotFlavor,
  ConversionProgress,
  ConversionResultSummary,
  ExportSheetStrategy,
  ExtractedTable,
  ExtractionEngineMode,
  OcrLanguageOption,
  PdfDocumentInspection,
  PythonBackendHealth,
} from './types';
import { inspectPdfDocument, parsePageRangeSpec } from './pdfAnalyzer';
import {
  evaluateTableDiagnostics,
  extractTablesFromPdfPage,
  mergeMultiPageExtractedTables,
} from './tableExtractor';
import { extractTablesWithOcr } from './ocrEngine';
import { exportTablesToXlsx } from './xlsxExporter';

function formatBytes(bytes: number): string {
  if (!bytes) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const i = Math.min(units.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
  return `${(bytes / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 2)} ${units[i]}`;
}

export const PdfToExcelConverter: React.FC = () => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const pdfDocRef = useRef<pdfjsLib.PDFDocumentProxy | null>(null);

  const [dragActive, setDragActive] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [inspection, setInspection] = useState<PdfDocumentInspection | null>(null);
  const [isInspecting, setIsInspecting] = useState(false);

  // Page selection
  const [selectedPages, setSelectedPages] = useState<number[]>([]);
  const [pageRangeInput, setPageRangeInput] = useState('all');

  // Engine & extraction options
  const [engineMode, setEngineMode] = useState<ExtractionEngineMode>('local_js');
  const [camelotFlavor, setCamelotFlavor] = useState<CamelotFlavor>('auto');
  const [mergeMultipage, setMergeMultipage] = useState(true);
  const [enableOcr, setEnableOcr] = useState(false);
  const [ocrLang, setOcrLang] = useState<OcrLanguageOption>('ara+eng');
  const [exportStrategy, setExportStrategy] = useState<ExportSheetStrategy>(
    'separate_sheets_per_table'
  );

  // Python FastAPI backend status
  const [pythonHealth, setPythonHealth] = useState<PythonBackendHealth>({
    available: false,
    camelot: false,
    pdfplumber: false,
  });

  // Progress, results & editing state
  const [isConverting, setIsConverting] = useState(false);
  const [progress, setProgress] = useState<ConversionProgress>({ percent: 0, stage: '' });
  const [resultSummary, setResultSummary] = useState<ConversionResultSummary | null>(null);
  const [tables, setTables] = useState<ExtractedTable[]>([]);
  const [activeTableIndex, setActiveTableIndex] = useState(0);
  const [editingCell, setEditingCell] = useState<{
    rowIdx: number;
    colIdx: number;
  } | null>(null);
  const [editingHeaderIdx, setEditingHeaderIdx] = useState<number | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [exportToast, setExportToast] = useState<string | null>(null);

  const checkPythonBackendHealth = async () => {
    try {
      const res = await fetch('/api/python-pdf/health');
      if (res.ok) {
        const data = await res.json();
        setPythonHealth({
          available: Boolean(data?.available),
          camelot: Boolean(data?.camelot),
          pdfplumber: Boolean(data?.pdfplumber),
        });
        return;
      }
    } catch {
      // ignore offline python check
    }
    setPythonHealth({ available: false, camelot: false, pdfplumber: false });
  };

  useEffect(() => {
    void checkPythonBackendHealth();
  }, []);

  const handleFileChosen = async (file: File) => {
    setErrorMsg(null);
    setResultSummary(null);
    setTables([]);
    setActiveTableIndex(0);

    if (file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf')) {
      setErrorMsg('يرجى اختيار ملف بصيغة PDF صالحة.');
      return;
    }

    setSelectedFile(file);
    setIsInspecting(true);
    setProgress({ percent: 5, stage: 'جاري فحص ملف الـ PDF واكتشاف نوع الصفحات...' });

    try {
      if (pdfDocRef.current) {
        const prevDoc = pdfDocRef.current as pdfjsLib.PDFDocumentProxy & {
          destroy?: () => Promise<void>;
        };
        if (typeof prevDoc.destroy === 'function') {
          await prevDoc.destroy().catch(() => {});
        } else if (typeof prevDoc.cleanup === 'function') {
          await prevDoc.cleanup().catch(() => {});
        }
        pdfDocRef.current = null;
      }

      const { inspection: info, pdfDoc } = await inspectPdfDocument(file, (pct, stage) => {
        setProgress({ percent: pct, stage });
      });

      pdfDocRef.current = pdfDoc;
      setInspection(info);

      const allPages = info.pages.map((p) => p.pageNumber);
      setSelectedPages(allPages);
      setPageRangeInput(allPages.length > 0 ? `1-${info.totalPages}` : 'all');

      // OCR remains strictly optional — off by default even if scanned pages exist
      setEnableOcr(false);
    } catch (err: any) {
      setErrorMsg(err?.message || 'تعذر قراءة ملف الـ PDF. تأكد من سلامة الملف وأنه غير تالف.');
      setInspection(null);
    } finally {
      setIsInspecting(false);
    }
  };

  const handleApplyPageRangeInput = (raw: string) => {
    setPageRangeInput(raw);
    if (!inspection) return;
    const parsed = parsePageRangeSpec(raw, inspection.totalPages);
    setSelectedPages(parsed);
  };

  const toggleSinglePage = (pNum: number) => {
    setSelectedPages((prev) => {
      const exists = prev.includes(pNum);
      const next = exists ? prev.filter((p) => p !== pNum) : [...prev, pNum].sort((a, b) => a - b);
      setPageRangeInput(next.join(', '));
      return next;
    });
  };

  const handleClearData = async () => {
    if (pdfDocRef.current) {
      const prevDoc = pdfDocRef.current as pdfjsLib.PDFDocumentProxy & {
        destroy?: () => Promise<void>;
      };
      if (typeof prevDoc.destroy === 'function') {
        await prevDoc.destroy().catch(() => {});
      } else if (typeof prevDoc.cleanup === 'function') {
        await prevDoc.cleanup().catch(() => {});
      }
      pdfDocRef.current = null;
    }
    setSelectedFile(null);
    setInspection(null);
    setSelectedPages([]);
    setPageRangeInput('all');
    setTables([]);
    setResultSummary(null);
    setActiveTableIndex(0);
    setEditingCell(null);
    setEditingHeaderIdx(null);
    setErrorMsg(null);
    setExportToast(null);
    setProgress({ percent: 0, stage: '' });
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleConvertPdf = async () => {
    if (!selectedFile || !inspection) return;
    if (selectedPages.length === 0) {
      setErrorMsg('يرجى تحديد صفحة واحدة على الأقل للتحويل.');
      return;
    }

    setErrorMsg(null);
    setIsConverting(true);
    setResultSummary(null);
    setTables([]);
    setActiveTableIndex(0);

    try {
      let extractedTables: ExtractedTable[] = [];
      const warnings: string[] = [];
      const errors: string[] = [];
      const skippedScannedPages: number[] = [];

      if (engineMode === 'python_fastapi') {
        setProgress({ percent: 15, stage: 'جاري إرسال الملف إلى خادم Python FastAPI (Camelot + pdfplumber)...' });
        const fileBase64 = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result as string);
          reader.onerror = reject;
          reader.readAsDataURL(selectedFile);
        });

        const res = await fetch('/api/python-pdf/extract', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fileBase64,
            fileName: selectedFile.name,
            pages: selectedPages.join(','),
            flavor: camelotFlavor,
            mergeMultipage,
          }),
        });

        const rawText = await res.text();
        let data: any = null;
        try {
          data = rawText ? JSON.parse(rawText) : null;
        } catch {
          data = null;
        }
        if (!res.ok) {
          if (res.status === 413) {
            throw new Error(
              'حجم ملف الـ PDF يتجاوز الحد المسموح لطلبات الخادم السحابي (4.5 ميجابايت على Vercel). يرجى استخدام محرك المتصفح المحلي (pdfjs-dist) الذي يعالج الملفات مباشرة بدون حد للحجم.'
            );
          }
          throw new Error(data?.error || 'تعذر استخراج الجداول عبر خادم Python FastAPI.');
        }

        extractedTables = Array.isArray(data?.tables) ? data.tables : [];
        if (extractedTables.length === 0) {
          warnings.push('لم يكتشف خادم Python FastAPI جداول صريحة في الصفحات المحددة.');
        }
      } else {
        // 100% open-source local JS engine: pdfjs-dist + optional Tesseract.js
        const pdfDoc = pdfDocRef.current;
        if (!pdfDoc) {
          throw new Error('لم يتم تحميل مستند الـ PDF بشكل صحيح. يرجى إعادة اختيار الملف.');
        }

        const pageInfoMap = new Map(inspection.pages.map((p) => [p.pageNumber, p]));
        const textPagesToProcess: number[] = [];
        const ocrPagesToProcess: number[] = [];

        for (const pNum of selectedPages) {
          const info = pageInfoMap.get(pNum);
          if (!info) continue;
          if (info.kind === 'scanned') {
            if (enableOcr) {
              ocrPagesToProcess.push(pNum);
            } else {
              skippedScannedPages.push(pNum);
            }
          } else if (info.kind === 'mixed') {
            textPagesToProcess.push(pNum);
            if (enableOcr && info.charCount < 60) {
              ocrPagesToProcess.push(pNum);
            }
          } else {
            textPagesToProcess.push(pNum);
          }
        }

        const totalWorkUnits = Math.max(1, textPagesToProcess.length + ocrPagesToProcess.length);
        let doneUnits = 0;

        for (const pNum of textPagesToProcess) {
          setProgress({
            percent: Math.round((doneUnits / totalWorkUnits) * 85),
            stage: `جاري استخراج الجداول والأعمدة من الصفحة النصية ${pNum}...`,
          });
          try {
            const pageTables = await extractTablesFromPdfPage(
              pdfDoc,
              pNum,
              extractedTables.length
            );
            extractedTables.push(...pageTables);
          } catch (pageErr: any) {
            errors.push(`تعذر استخراج الصفحة ${pNum}: ${pageErr?.message || 'خطأ غير معروف'}`);
          }
          doneUnits++;
        }

        if (ocrPagesToProcess.length > 0 && enableOcr) {
          try {
            const ocrTables = await extractTablesWithOcr(
              pdfDoc,
              ocrPagesToProcess,
              ocrLang,
              extractedTables.length,
              (ocrPct, ocrStage) => {
                const overall = Math.min(
                  95,
                  Math.round(
                    ((doneUnits + (ocrPct / 100) * ocrPagesToProcess.length) / totalWorkUnits) *
                      90
                  )
                );
                setProgress({ percent: overall, stage: ocrStage });
              }
            );
            extractedTables.push(...ocrTables);
          } catch (ocrErr: any) {
            errors.push(
              `تعذر تشغيل OCR على الصفحات الممسوحة (${ocrPagesToProcess.join(', ')}): ${
                ocrErr?.message || 'تحقق من الاتصال بتحميل حزمة اللغة أو أعد المحاولة.'
              }`
            );
          }
        }

        if (skippedScannedPages.length > 0) {
          warnings.push(
            `تم تخطي ${skippedScannedPages.length} صفحة ممسوحة ضوئياً (${skippedScannedPages.join(
              ', '
            )}) لأن خيار OCR الاختياري غير مفعّل.`
          );
        }

        if (mergeMultipage && extractedTables.length > 1) {
          setProgress({ percent: 94, stage: 'جاري ربط الجداول الممتدة عبر عدة صفحات...' });
          extractedTables = mergeMultiPageExtractedTables(extractedTables);
        }
      }

      if (extractedTables.length === 0) {
        warnings.push(
          'لم يتم اكتشاف جداول منظمة في الصفحات المختارة. إذا كان الملف ممسوحاً ضوئياً، فعّل خيار OCR الاختياري وأعد المحاولة.'
        );
      }

      extractedTables.forEach((tbl) => {
        tbl.warnings.forEach((w) => {
          if (!warnings.includes(w)) warnings.push(`${tbl.name}: ${w}`);
        });
      });

      const overallAccuracy =
        extractedTables.length > 0
          ? Math.round(
              extractedTables.reduce((sum, t) => sum + t.accuracyScore, 0) /
                extractedTables.length
            )
          : 0;

      if (extractedTables.length > 0 && overallAccuracy < 80) {
        warnings.unshift(
          `تنبيه انخفاض الدقة: متوسط دقة الاستخراج الإجمالي هو ${overallAccuracy}% — يُنصح بمراجعة الخلايا وتعديلها في الجدول التفاعلي قبل التصدير.`
        );
      }

      setTables(extractedTables);
      setResultSummary({
        tables: extractedTables,
        processedPages: selectedPages,
        skippedScannedPages,
        overallAccuracy,
        warnings,
        errors,
      });
      setProgress({ percent: 100, stage: 'اكتمل التحويل! يمكنك مراجعة البيانات وتعديلها وتصدير Excel.' });
    } catch (err: any) {
      setErrorMsg(err?.message || 'حدث خطأ غير متوقع أثناء استخراج الجداول.');
    } finally {
      setIsConverting(false);
    }
  };

  const activeTable = tables[activeTableIndex] || null;

  const updateActiveTable = (updater: (tbl: ExtractedTable) => ExtractedTable) => {
    setTables((prev) =>
      prev.map((t, idx) => (idx === activeTableIndex ? updater(t) : t))
    );
  };

  const handleCellEdit = (rowIdx: number, colIdx: number, value: string) => {
    updateActiveTable((tbl) => {
      const nextRows = tbl.rows.map((r, rI) => {
        if (rI !== rowIdx) return r;
        const copy = [...r];
        while (copy.length < tbl.headers.length) copy.push('');
        copy[colIdx] = value;
        return copy;
      });
      const diag = evaluateTableDiagnostics(tbl.headers, nextRows, tbl.accuracyScore);
      return {
        ...tbl,
        rows: nextRows,
        emptyCellRatio: diag.emptyCellRatio,
      };
    });
  };

  const handleHeaderEdit = (colIdx: number, value: string) => {
    updateActiveTable((tbl) => {
      const nextHeaders = [...tbl.headers];
      nextHeaders[colIdx] = value;
      return { ...tbl, headers: nextHeaders };
    });
  };

  const handleToggleTableDirection = () => {
    updateActiveTable((tbl) => ({
      ...tbl,
      isRtl: !tbl.isRtl,
      headers: [...tbl.headers].reverse(),
      rows: tbl.rows.map((r) => [...r].reverse()),
    }));
  };

  const handleAddRow = () => {
    updateActiveTable((tbl) => ({
      ...tbl,
      rows: [...tbl.rows, new Array(tbl.headers.length).fill('')],
    }));
  };

  const handleAddColumn = () => {
    updateActiveTable((tbl) => ({
      ...tbl,
      headers: [...tbl.headers, `عمود_${tbl.headers.length + 1}`],
      rows: tbl.rows.map((r) => [...r, '']),
    }));
  };

  const handleExportExcel = () => {
    try {
      exportTablesToXlsx(tables, selectedFile?.name || 'converted', exportStrategy);
      setExportToast('تم إنشاء وتنزيل ملف Excel (.xlsx) بنجاح!');
      setTimeout(() => setExportToast(null), 4000);
    } catch (err: any) {
      setErrorMsg(err?.message || 'تعذر تصدير ملف Excel.');
    }
  };

  const totalExtractedRows = useMemo(
    () => tables.reduce((sum, t) => sum + t.rows.length, 0),
    [tables]
  );

  return (
    <section className="space-y-6" dir="rtl">
      {/* Module Title & Open Source Guarantee Header */}
      <div className="bg-white rounded-xl border border-slate-200 p-5 sm:p-6 shadow-xs flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-start gap-3.5">
          <div className="w-12 h-12 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-xs">
            <FileSpreadsheet className="w-6 h-6" />
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h2 className="text-lg sm:text-xl font-extrabold text-slate-900">
                PDF to Excel Converter
              </h2>
              <span className="px-2.5 py-0.5 text-[11px] font-bold rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200">
                مفتوح المصدر · بدون AI أو API مدفوع
              </span>
            </div>
            <p className="text-xs sm:text-sm text-slate-600 mt-1">
              تحويل دقيق لملفات PDF النصية والممسوحة ضوئياً إلى Excel عبر{' '}
              <span className="font-mono font-bold text-slate-800">pdfjs-dist</span> و{' '}
              <span className="font-mono font-bold text-slate-800">SheetJS</span> و{' '}
              <span className="font-mono font-bold text-slate-800">Tesseract.js OCR</span> مع جاهزية الربط مع{' '}
              <span className="font-mono font-bold text-slate-800">Python FastAPI (Camelot + pdfplumber)</span>.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 text-xs font-bold text-slate-700 bg-slate-50 px-3 py-2 rounded-lg border border-slate-200">
          <ShieldCheck className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>حفظ تام للأرقام والتواريخ والخلايا الفارغة دون تخمين</span>
        </div>
      </div>

      {/* Upload Dropzone Card */}
      <div className="bg-white rounded-xl border border-slate-200 p-5 sm:p-6 shadow-xs space-y-5">
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setDragActive(true);
          }}
          onDragLeave={() => setDragActive(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragActive(false);
            if (e.dataTransfer.files?.[0]) {
              void handleFileChosen(e.dataTransfer.files[0]);
            }
          }}
          onClick={() => !isInspecting && !isConverting && fileInputRef.current?.click()}
          className={`border-2 border-dashed rounded-xl p-8 text-center transition-all cursor-pointer ${
            dragActive
              ? 'border-emerald-500 bg-emerald-50/70'
              : selectedFile
              ? 'border-emerald-300 bg-emerald-50/20'
              : 'border-slate-300 hover:border-emerald-400 hover:bg-slate-50'
          }`}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".pdf,application/pdf"
            className="hidden"
            onChange={(e) => {
              if (e.target.files?.[0]) {
                void handleFileChosen(e.target.files[0]);
              }
            }}
          />

          <div className="space-y-3">
            <div className="w-14 h-14 mx-auto rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center">
              <UploadCloud className="w-7 h-7" />
            </div>
            {selectedFile ? (
              <div className="space-y-1">
                <p className="text-sm sm:text-base font-extrabold text-slate-900">
                  {selectedFile.name}
                </p>
                <p className="text-xs font-semibold text-slate-600">
                  الحجم: <span className="font-mono">{formatBytes(selectedFile.size)}</span>
                  {inspection && (
                    <>
                      {' · '}
                      عدد الصفحات: <span className="font-mono font-bold">{inspection.totalPages}</span>
                    </>
                  )}
                  {' · '}انقر لاختيار ملف PDF آخر
                </p>
              </div>
            ) : (
              <div>
                <p className="text-sm sm:text-base font-bold text-slate-800">
                  اسحب وأفلت ملف الـ PDF هنا، أو{' '}
                  <span className="text-emerald-700 underline underline-offset-4">
                    اختر ملفاً من جهازك
                  </span>
                </p>
                <p className="text-xs text-slate-500 mt-1.5">
                  يدعم الملفات النصية والممسوحة ضوئياً متعددة الصفحات · العربية والإنجليزية
                </p>
              </div>
            )}
          </div>
        </div>

        {/* File Inspection Summary + Page Picker + Engine Options */}
        {inspection && (
          <div className="space-y-5 pt-2">
            {/* Scan Type Summary Badges */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <div className="p-3 rounded-lg bg-slate-50 border border-slate-200">
                <div className="text-xs font-semibold text-slate-500">نوع المستند المكتشف</div>
                <div className="text-sm font-extrabold text-slate-900 mt-1">
                  {inspection.documentKind === 'text'
                    ? 'ملف نصي (Selectable Text)'
                    : inspection.documentKind === 'scanned'
                    ? 'ملف ممسوح ضوئياً (Scanned)'
                    : 'مستند مختلط (نصوص + صور)'}
                </div>
              </div>
              <div className="p-3 rounded-lg bg-emerald-50/70 border border-emerald-200">
                <div className="text-xs font-semibold text-emerald-700">صفحات نصية</div>
                <div className="text-lg font-mono font-extrabold text-emerald-950 mt-0.5">
                  {inspection.textPagesCount}
                </div>
              </div>
              <div className="p-3 rounded-lg bg-amber-50/80 border border-amber-200">
                <div className="text-xs font-semibold text-amber-800">صفحات ممسوحة ضوئياً</div>
                <div className="text-lg font-mono font-extrabold text-amber-950 mt-0.5">
                  {inspection.scannedPagesCount}
                </div>
              </div>
              <div className="p-3 rounded-lg bg-sky-50/80 border border-sky-200">
                <div className="text-xs font-semibold text-sky-800">صفحات مختلطة</div>
                <div className="text-lg font-mono font-extrabold text-sky-950 mt-0.5">
                  {inspection.mixedPagesCount}
                </div>
              </div>
            </div>

            {/* Page Selection Controls */}
            <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/60 space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="text-xs sm:text-sm font-extrabold text-slate-900 flex items-center gap-1.5">
                  <FileText className="w-4 h-4 text-emerald-700" />
                  <span>اختيار الصفحات المطلوب تحويلها ({selectedPages.length} / {inspection.totalPages}):</span>
                </span>
                <div className="flex flex-wrap items-center gap-1.5 text-xs">
                  <button
                    type="button"
                    onClick={() => {
                      const all = inspection.pages.map((p) => p.pageNumber);
                      setSelectedPages(all);
                      setPageRangeInput(`1-${inspection.totalPages}`);
                    }}
                    className="px-2.5 py-1 bg-white hover:bg-slate-100 border border-slate-200 rounded-md font-bold cursor-pointer"
                  >
                    تحديد الكل
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      const txt = inspection.pages
                        .filter((p) => p.kind === 'text' || p.kind === 'mixed')
                        .map((p) => p.pageNumber);
                      setSelectedPages(txt);
                      setPageRangeInput(txt.join(', '));
                    }}
                    className="px-2.5 py-1 bg-white hover:bg-emerald-50 border border-slate-200 rounded-md font-bold cursor-pointer"
                  >
                    النصية فقط
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setSelectedPages([]);
                      setPageRangeInput('');
                    }}
                    className="px-2.5 py-1 bg-white hover:bg-rose-50 border border-slate-200 rounded-md font-bold cursor-pointer"
                  >
                    إلغاء التحديد
                  </button>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <label className="text-xs font-bold text-slate-700">نطاق الصفحات:</label>
                <input
                  type="text"
                  value={pageRangeInput}
                  onChange={(e) => handleApplyPageRangeInput(e.target.value)}
                  placeholder="مثال: 1-3, 5, 8"
                  className="px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-mono font-bold w-48 focus:outline-hidden focus:border-emerald-600"
                  dir="ltr"
                />
              </div>

              {/* Clickable per-page chips */}
              <div className="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto pt-1">
                {inspection.pages.map((p) => {
                  const isSelected = selectedPages.includes(p.pageNumber);
                  const kindLabel =
                    p.kind === 'scanned' ? 'ممسوحة' : p.kind === 'mixed' ? 'مختلطة' : 'نصية';
                  return (
                    <button
                      key={p.pageNumber}
                      type="button"
                      onClick={() => toggleSinglePage(p.pageNumber)}
                      className={`px-2.5 py-1 rounded-md text-xs font-bold border transition-colors cursor-pointer flex items-center gap-1 ${
                        isSelected
                          ? 'bg-emerald-600 text-white border-emerald-600'
                          : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                      }`}
                    >
                      <span className="font-mono">ص {p.pageNumber}</span>
                      <span
                        className={`text-[10px] px-1 rounded ${
                          isSelected
                            ? 'bg-emerald-700 text-emerald-50'
                            : p.kind === 'scanned'
                            ? 'bg-amber-100 text-amber-800'
                            : 'bg-slate-100 text-slate-600'
                        }`}
                      >
                        {kindLabel}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Engine & Optional OCR Settings */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              {/* Engine Selection */}
              <div className="p-4 rounded-xl border border-slate-200 bg-white space-y-3">
                <div className="flex items-center justify-between">
                  <span className="text-xs sm:text-sm font-extrabold text-slate-900 flex items-center gap-1.5">
                    <Cpu className="w-4 h-4 text-emerald-700" />
                    <span>محرك الاستخراج:</span>
                  </span>
                  <button
                    type="button"
                    onClick={() => void checkPythonBackendHealth()}
                    className="text-[11px] font-bold text-slate-600 hover:text-emerald-700 flex items-center gap-1 cursor-pointer"
                  >
                    <RefreshCw className="w-3 h-3" />
                    <span>فحص خادم Python</span>
                  </button>
                </div>

                <div className="space-y-2 text-xs">
                  <label className="flex items-start gap-2.5 p-2.5 rounded-lg border border-slate-200 cursor-pointer hover:bg-slate-50">
                    <input
                      type="radio"
                      name="pdfEngineMode"
                      checked={engineMode === 'local_js'}
                      onChange={() => setEngineMode('local_js')}
                      className="mt-0.5 text-emerald-600"
                    />
                    <div>
                      <div className="font-extrabold text-slate-900">
                        المحرك المحلي المدمج (pdfjs-dist + Tesseract.js)
                      </div>
                      <div className="text-slate-500 mt-0.5">
                        يعمل مباشرة في المتصفح مجاناً دون الحاجة لتشغيل خادم خارجي.
                      </div>
                    </div>
                  </label>

                  <label className="flex items-start gap-2.5 p-2.5 rounded-lg border border-slate-200 cursor-pointer hover:bg-slate-50">
                    <input
                      type="radio"
                      name="pdfEngineMode"
                      checked={engineMode === 'python_fastapi'}
                      onChange={() => setEngineMode('python_fastapi')}
                      className="mt-0.5 text-emerald-600"
                    />
                    <div className="flex-1">
                      <div className="flex items-center justify-between gap-2 flex-wrap">
                        <span className="font-extrabold text-slate-900 flex items-center gap-1">
                          <Server className="w-3.5 h-3.5 text-indigo-600" />
                          <span>خادم Python FastAPI (Camelot + pdfplumber)</span>
                        </span>
                        <span
                          className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                            pythonHealth.available
                              ? 'bg-emerald-100 text-emerald-800'
                              : 'bg-slate-100 text-slate-600'
                          }`}
                        >
                          {pythonHealth.available ? 'متصل وجاهز' : 'غير متصل (اختياري)'}
                        </span>
                      </div>
                      <div className="text-slate-500 mt-0.5">
                        مخصص للجداول المعقدة والمؤطرة عبر FastAPI في{' '}
                        <span className="font-mono">backend/pdf_extractor_api.py</span>.
                      </div>

                      {engineMode === 'python_fastapi' && (
                        <div className="mt-2 flex items-center gap-2">
                          <span className="font-bold text-slate-700">وضع Camelot:</span>
                          <select
                            value={camelotFlavor}
                            onChange={(e) => setCamelotFlavor(e.target.value as CamelotFlavor)}
                            className="px-2 py-1 border border-slate-300 rounded text-xs font-bold bg-white"
                          >
                            <option value="auto">تلقائي (Lattice ثم Stream)</option>
                            <option value="lattice">Lattice (للجداول ذات الحدود)</option>
                            <option value="stream">Stream (للجداول بدون خطوط)</option>
                          </select>
                        </div>
                      )}
                    </div>
                  </label>
                </div>
              </div>

              {/* Optional OCR & Multi-page Merging */}
              <div className="p-4 rounded-xl border border-slate-200 bg-white space-y-3">
                <span className="text-xs sm:text-sm font-extrabold text-slate-900 flex items-center gap-1.5">
                  <ScanText className="w-4 h-4 text-amber-600" />
                  <span>خيارات التعرف الضوئي OCR ودمج الصفحات:</span>
                </span>

                <label className="flex items-start gap-2.5 p-2.5 rounded-lg border border-amber-200 bg-amber-50/40 cursor-pointer text-xs">
                  <input
                    type="checkbox"
                    checked={enableOcr}
                    onChange={(e) => setEnableOcr(e.target.checked)}
                    className="mt-0.5 rounded border-slate-300 text-amber-600"
                  />
                  <div>
                    <div className="font-extrabold text-slate-900">
                      تفعيل OCR اختيارياً للصفحات الممسوحة ضوئياً (Tesseract.js)
                    </div>
                    <div className="text-slate-600 mt-0.5">
                      {inspection.scannedPagesCount > 0 || inspection.mixedPagesCount > 0
                        ? `تم اكتشاف ${
                            inspection.scannedPagesCount + inspection.mixedPagesCount
                          } صفحة ممسوحة/مختلطة؛ فعّل هذا الخيار فقط إذا أردت قراءتها عبر OCR.`
                        : 'الملف نصي بالكامل حالياً ولا يحتاج إلى OCR، ويمكنك تفعيله عند الحاجة.'}
                    </div>
                  </div>
                </label>

                {enableOcr && (
                  <div className="flex items-center gap-2 text-xs pr-2">
                    <label className="font-bold text-slate-700">لغة التعرف الضوئي (OCR):</label>
                    <select
                      value={ocrLang}
                      onChange={(e) => setOcrLang(e.target.value as OcrLanguageOption)}
                      className="px-2.5 py-1.5 border border-slate-300 rounded-md font-bold bg-white"
                    >
                      <option value="ara+eng">العربية + الإنجليزية (ara+eng)</option>
                      <option value="ara">العربية فقط (ara)</option>
                      <option value="eng">الإنجليزية فقط (eng)</option>
                    </select>
                  </div>
                )}

                <label className="flex items-center gap-2.5 p-2.5 rounded-lg border border-slate-200 cursor-pointer text-xs">
                  <input
                    type="checkbox"
                    checked={mergeMultipage}
                    onChange={(e) => setMergeMultipage(e.target.checked)}
                    className="rounded border-slate-300 text-emerald-600"
                  />
                  <span className="font-bold text-slate-800">
                    معالجة ودمج الجداول التي تمتد عبر أكثر من صفحة تلقائياً
                  </span>
                </label>
              </div>
            </div>

            {/* Convert Action Button */}
            <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
              <div className="flex flex-wrap items-center gap-2.5">
                <button
                  type="button"
                  onClick={handleConvertPdf}
                  disabled={isConverting || isInspecting || selectedPages.length === 0}
                  className="px-6 py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold text-sm rounded-xl shadow-xs transition-all flex items-center gap-2 cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <FileSpreadsheet className="w-5 h-5" />
                  <span>
                    {isConverting ? 'جاري استخراج الجداول...' : 'بدء تحويل PDF إلى Excel الآن'}
                  </span>
                </button>

                <button
                  type="button"
                  onClick={() => void handleClearData()}
                  disabled={isConverting || isInspecting}
                  className="px-4 py-3 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 font-bold text-sm rounded-xl transition-all flex items-center gap-2 cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
                  title="مسح الملف والجداول المستخرجة"
                >
                  <Trash2 className="w-4 h-4 text-rose-600" />
                  <span>مسح البيانات</span>
                </button>
              </div>

              {tables.length > 0 && (
                <span className="text-xs font-bold text-emerald-800 bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-200">
                  تم اكتشاف {tables.length} جدول · {totalExtractedRows} صف جاهز للمعاينة والتصدير
                </span>
              )}
            </div>
          </div>
        )}

        {/* Progress Indicator */}
        {(isInspecting || isConverting) && (
          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
            <div className="flex items-center justify-between text-xs font-bold text-slate-800">
              <span>{progress.stage || 'جاري المعالجة...'}</span>
              <span className="font-mono text-emerald-800">{progress.percent}%</span>
            </div>
            <progress
              value={progress.percent}
              max={100}
              className="w-full h-2.5 rounded-full overflow-hidden accent-emerald-600"
            />
          </div>
        )}

        {/* Errors Alert */}
        {errorMsg && (
          <div className="p-4 rounded-xl bg-rose-50 border border-rose-200 text-rose-900 text-xs sm:text-sm flex items-start gap-2.5">
            <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
            <div className="font-bold">{errorMsg}</div>
          </div>
        )}
      </div>

      {/* Extraction Diagnostics, Stats & Low Accuracy Warnings */}
      {resultSummary && (
        <div className="space-y-4">
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <div className="text-xs font-bold text-slate-500">الصفحات المعالجة</div>
              <div className="text-xl font-mono font-extrabold text-slate-900 mt-1">
                {resultSummary.processedPages.length}
              </div>
            </div>
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <div className="text-xs font-bold text-slate-500">الجداول المكتشفة</div>
              <div className="text-xl font-mono font-extrabold text-emerald-700 mt-1">
                {tables.length}
              </div>
            </div>
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <div className="text-xs font-bold text-slate-500">إجمالي الصفوف</div>
              <div className="text-xl font-mono font-extrabold text-slate-900 mt-1">
                {totalExtractedRows}
              </div>
            </div>
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
              <div className="text-xs font-bold text-slate-500">متوسط دقة الاستخراج</div>
              <div
                className={`text-xl font-mono font-extrabold mt-1 ${
                  resultSummary.overallAccuracy >= 80 ? 'text-emerald-700' : 'text-amber-600'
                }`}
              >
                {resultSummary.overallAccuracy}%
              </div>
            </div>
          </div>

          {resultSummary.warnings.length > 0 && (
            <div className="p-4 rounded-xl bg-amber-50 border border-amber-300 text-amber-950 text-xs space-y-1.5">
              <div className="font-extrabold flex items-center gap-1.5 text-amber-900">
                <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                <span>تنبيهات دقة الاستخراج والمراجعة:</span>
              </div>
              <ul className="list-disc pr-5 space-y-1 font-semibold">
                {resultSummary.warnings.map((w, idx) => (
                  <li key={idx}>{w}</li>
                ))}
              </ul>
            </div>
          )}

          {resultSummary.errors.length > 0 && (
            <div className="p-4 rounded-xl bg-rose-50 border border-rose-200 text-rose-900 text-xs space-y-1.5">
              <div className="font-extrabold flex items-center gap-1.5">
                <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
                <span>أخطاء جزئية أثناء استخراج بعض الصفحات:</span>
              </div>
              <ul className="list-disc pr-5 space-y-1 font-semibold">
                {resultSummary.errors.map((er, idx) => (
                  <li key={idx}>{er}</li>
                ))}
              </ul>
            </div>
          )}
        </div>
      )}

      {/* Interactive Table Preview, Editing & Excel Export */}
      {tables.length > 0 && activeTable && (
        <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
          {/* Export & Sheet Strategy Bar */}
          <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2">
              {tables.map((tbl, idx) => (
                <button
                  key={tbl.id}
                  type="button"
                  onClick={() => {
                    setActiveTableIndex(idx);
                    setEditingCell(null);
                    setEditingHeaderIdx(null);
                  }}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold border transition-all cursor-pointer flex items-center gap-1.5 ${
                    idx === activeTableIndex
                      ? 'bg-emerald-600 text-white border-emerald-600 shadow-2xs'
                      : 'bg-white text-slate-800 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  <Layers className="w-3.5 h-3.5" />
                  <span>{tbl.name}</span>
                  <span
                    className={`font-mono text-[10px] px-1.5 py-0.5 rounded ${
                      idx === activeTableIndex
                        ? 'bg-emerald-700 text-white'
                        : tbl.accuracyScore >= 80
                        ? 'bg-emerald-50 text-emerald-800'
                        : 'bg-amber-100 text-amber-800'
                    }`}
                  >
                    {tbl.accuracyScore}%
                  </span>
                </button>
              ))}
            </div>

            <div className="flex flex-wrap items-center gap-2 text-xs">
              <select
                value={exportStrategy}
                onChange={(e) => setExportStrategy(e.target.value as ExportSheetStrategy)}
                className="px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg font-bold text-slate-800"
                title="طريقة توزيع أوراق العمل داخل ملف Excel"
              >
                <option value="separate_sheets_per_table">Sheet منفصل لكل جدول</option>
                <option value="separate_sheets_per_page">Sheet منفصل لكل صفحة</option>
                <option value="single_merged_sheet">دمج كافة الجداول في Sheet واحد</option>
              </select>

              <button
                type="button"
                onClick={handleExportExcel}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold rounded-lg shadow-xs transition-all flex items-center gap-1.5 cursor-pointer"
              >
                <Download className="w-4 h-4" />
                <span>تصدير Excel (.xlsx)</span>
              </button>
            </div>
          </div>

          {exportToast && (
            <div className="px-4 py-2.5 bg-emerald-50 border-b border-emerald-200 text-emerald-900 text-xs font-bold flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-600" />
              <span>{exportToast}</span>
            </div>
          )}

          {/* Table Edit Controls */}
          <div className="px-4 py-2.5 bg-white border-b border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs">
            <div className="flex flex-wrap items-center gap-2 text-slate-600 font-bold">
              <span>
                المحرك المستخدم: <span className="font-mono text-slate-900">{activeTable.engine}</span>
              </span>
              <span>·</span>
              <span>
                الصفحات: <span className="font-mono text-slate-900">{activeTable.pages.join(', ')}</span>
              </span>
              <span>·</span>
              <span>
                دقة المحاذاة:{' '}
                <span
                  className={`font-mono ${
                    activeTable.accuracyScore >= 80 ? 'text-emerald-700' : 'text-amber-700'
                  }`}
                >
                  {activeTable.accuracyScore}%
                </span>
              </span>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={handleToggleTableDirection}
                className="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-md font-bold text-slate-800 flex items-center gap-1 cursor-pointer"
              >
                <ArrowLeftRight className="w-3.5 h-3.5 text-emerald-700" />
                <span>الاتجاه: {activeTable.isRtl ? 'RTL (يمين لليسار)' : 'LTR (يسار لليمين)'}</span>
              </button>
              <button
                type="button"
                onClick={handleAddRow}
                className="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-md font-bold text-slate-800 flex items-center gap-1 cursor-pointer"
              >
                <Plus className="w-3.5 h-3.5 text-emerald-700" />
                <span>إضافة صف</span>
              </button>
              <button
                type="button"
                onClick={handleAddColumn}
                className="px-2.5 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-md font-bold text-slate-800 flex items-center gap-1 cursor-pointer"
              >
                <Columns3 className="w-3.5 h-3.5 text-emerald-700" />
                <span>إضافة عمود</span>
              </button>
            </div>
          </div>

          {/* Interactive Editable Table */}
          <div className="overflow-auto max-h-[540px]" dir={activeTable.isRtl ? 'rtl' : 'ltr'}>
            <table className="w-full border-collapse text-xs sm:text-sm">
              <thead className="sticky top-0 z-10 bg-slate-100 border-b-2 border-slate-300">
                <tr>
                  <th className="w-12 px-2 py-2 text-center border-r border-slate-300 font-mono font-bold text-slate-600 bg-slate-200">
                    #
                  </th>
                  {activeTable.headers.map((header, colIdx) => (
                    <th
                      key={colIdx}
                      onClick={() => setEditingHeaderIdx(colIdx)}
                      className="px-3 py-2 text-right border-r border-slate-300 font-extrabold text-slate-900 cursor-pointer hover:bg-slate-200/70 min-w-[120px]"
                      title="انقر لتعديل عنوان العمود"
                    >
                      {editingHeaderIdx === colIdx ? (
                        <input
                          type="text"
                          autoFocus
                          value={header}
                          onChange={(e) => handleHeaderEdit(colIdx, e.target.value)}
                          onBlur={() => setEditingHeaderIdx(null)}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter' || e.key === 'Escape') {
                              setEditingHeaderIdx(null);
                            }
                          }}
                          className="w-full px-1.5 py-0.5 bg-white border border-emerald-500 rounded font-bold outline-hidden"
                        />
                      ) : (
                        header || `عمود_${colIdx + 1}`
                      )}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {activeTable.rows.map((row, rIdx) => (
                  <tr key={rIdx} className="hover:bg-emerald-50/40 border-b border-slate-200">
                    <td className="px-2 py-1.5 text-center border-r border-slate-200 font-mono font-bold text-slate-500 bg-slate-50 select-none">
                      {rIdx + 1}
                    </td>
                    {activeTable.headers.map((_, cIdx) => {
                      const val = row[cIdx] ?? '';
                      const isEditing =
                        editingCell?.rowIdx === rIdx && editingCell?.colIdx === cIdx;
                      const isEmpty = val.trim() === '';

                      return (
                        <td
                          key={cIdx}
                          onClick={() => setEditingCell({ rowIdx: rIdx, colIdx: cIdx })}
                          className={`px-3 py-1.5 border-r border-slate-200 font-bold text-slate-900 cursor-text ${
                            isEmpty ? 'bg-slate-50/60' : 'bg-white'
                          }`}
                          title="انقر لتعديل الخلية مباشرة"
                        >
                          {isEditing ? (
                            <input
                              type="text"
                              autoFocus
                              value={val}
                              onChange={(e) => handleCellEdit(rIdx, cIdx, e.target.value)}
                              onBlur={() => setEditingCell(null)}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter' || e.key === 'Escape') {
                                  setEditingCell(null);
                                }
                              }}
                              className="w-full px-1.5 py-0.5 bg-white border-2 border-emerald-500 rounded font-bold outline-hidden"
                            />
                          ) : isEmpty ? (
                            <span className="text-slate-300 select-none text-xs">—</span>
                          ) : (
                            val
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="px-4 py-2.5 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs font-bold text-slate-600">
            <span>
              💡 انقر على أي خلية أو عنوان عمود لتعديله مباشرة قبل التصدير. جميع الخلايا الفارغة محفوظة دون حذف.
            </span>
            <span className="font-mono text-slate-800">
              {activeTable.rows.length} صف × {activeTable.headers.length} عمود
            </span>
          </div>
        </div>
      )}
    </section>
  );
};
