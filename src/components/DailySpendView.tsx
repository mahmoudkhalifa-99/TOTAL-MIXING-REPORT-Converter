import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  Download,
  Copy,
  Check,
  FileSpreadsheet,
  Printer,
  EyeOff,
  Eye,
  Plus,
  X,
  Trash2,
  Pencil,
} from 'lucide-react';
import { SheetData } from '../types/converter';
import {
  calculateDailySpend,
  DAILY_SPEND_ITEMS,
  DailySpendCustomItem,
  DailySpendCodeOverrides,
  DAILY_SPEND_CUSTOM_ITEMS_KEY,
  DAILY_SPEND_CODE_OVERRIDES_KEY,
  loadDailySpendCustomItems,
  loadDailySpendCodeOverrides,
} from '../utils/dailySpendCalculator';
import { calculateSadatControlReport } from '../utils/sadatControlCalculator';
import { exportDailySpendStyledExcel } from '../utils/styledExcelExport';
import { exportToCsv, copyTableToClipboard } from '../utils/excelExport';

interface DailySpendViewProps {
  sheet: SheetData;
  fileName?: string;
  documentTitle?: string;
}

type UnitMode = 'ton' | 'kg';
type Metric = 'actual' | 'target';

/** A code cell the user can click to edit. Saving an empty value goes back to the automatic code. */
const EditableCode: React.FC<{
  value: string;
  isManual: boolean;
  onSave: (v: string) => void;
  label: string;
}> = ({ value, isManual, onSave, label }) => {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(value);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (editing) {
      setDraft(value);
      inputRef.current?.focus();
      inputRef.current?.select();
    }
  }, [editing]); // eslint-disable-line react-hooks/exhaustive-deps

  const commit = () => {
    setEditing(false);
    if (draft.trim() !== value) onSave(draft.trim());
  };

  if (editing) {
    return (
      <input
        ref={inputRef}
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') commit();
          if (e.key === 'Escape') setEditing(false);
        }}
        className="w-full min-w-[70px] text-center font-mono text-xs font-bold border border-emerald-500 rounded px-1 py-0.5 outline-none bg-white"
        aria-label={label}
        dir="ltr"
      />
    );
  }

  return (
    <button
      type="button"
      onClick={() => setEditing(true)}
      title={`انقر لتعديل ${label}`}
      className={`w-full min-h-[22px] font-mono text-xs font-bold cursor-pointer hover:bg-emerald-50 rounded px-1 ${
        isManual ? 'text-emerald-700' : value ? 'text-slate-900' : 'text-slate-400'
      }`}
      dir="ltr"
    >
      {value || '＋'}
    </button>
  );
};

export const DailySpendView: React.FC<DailySpendViewProps> = ({ sheet, fileName, documentTitle }) => {
  const [unit, setUnit] = useState<UnitMode>('ton');
  const [metric, setMetric] = useState<Metric>('actual');
  const [hideEmpty, setHideEmpty] = useState<boolean>(false);
  const [copied, setCopied] = useState<boolean>(false);

  const [customItems, setCustomItems] = useState<DailySpendCustomItem[]>(() => loadDailySpendCustomItems());
  const [codeOverrides, setCodeOverrides] = useState<DailySpendCodeOverrides>(() => loadDailySpendCodeOverrides());

  const [showAddForm, setShowAddForm] = useState<boolean>(false);
  const [form, setForm] = useState({ en: '', ar: '', code: '', driveCode: '' });
  const [formError, setFormError] = useState<string>('');

  // Lets the print stylesheet print ONLY this report (see index.css)
  useEffect(() => {
    document.body.classList.add('print-daily-spend');
    return () => document.body.classList.remove('print-daily-spend');
  }, []);

  const divisor = unit === 'ton' ? 1000 : 1;
  const decimals = unit === 'ton' ? 4 : 2;
  const unitLabel = unit === 'ton' ? 'طن' : 'كجم';

  const persist = (key: string, value: unknown) => {
    try {
      window.localStorage.setItem(key, JSON.stringify(value));
    } catch {
      // storage unavailable: changes still apply for this session
    }
  };

  const report = useMemo(
    () => calculateDailySpend(sheet, metric, divisor, documentTitle, customItems, codeOverrides),
    [sheet, metric, divisor, documentTitle, customItems, codeOverrides]
  );

  // Category totals (fattening / fish / duck / pets) for the WHOLE file, summed automatically by formula type.
  // They are shown AFTER deducting water, and WITHOUT any "Cleaning" group (e.g. Cleaning Aqua).
  const categoryInfo = useMemo(() => {
    try {
      const rep = calculateSadatControlReport(sheet, metric, divisor, documentTitle);
      const isCleaning = (label: string) => /clean|تنظيف/i.test(label || '');
      const nonWaterRows = rep.rows.filter((r) => !r.isWater);

      const sumFormulas = (formulas: { key: string; label: string }[]) =>
        formulas
          .filter((f) => !isCleaning(f.label))
          .reduce((acc, f) => acc + nonWaterRows.reduce((a, r) => a + (r.formulaValues[f.key] || 0), 0), 0);

      const allFormulas = [
        ...rep.fatteningFormulas,
        ...rep.fishFormulas,
        ...rep.duckFormulas,
        ...rep.petsFormulas,
      ];
      const waterDeducted = allFormulas
        .filter((f) => !isCleaning(f.label))
        .reduce(
          (acc, f) =>
            acc + rep.rows.filter((r) => r.isWater).reduce((a, r) => a + (r.formulaValues[f.key] || 0), 0),
          0
        );
      const cleaningExcluded = allFormulas
        .filter((f) => isCleaning(f.label))
        .reduce((acc, f) => acc + rep.rows.reduce((a, r) => a + (r.formulaValues[f.key] || 0), 0), 0);

      return {
        totals: [
          { label: 'اجمالى التسمين', value: sumFormulas(rep.fatteningFormulas) },
          { label: 'اجمالى السمك', value: sumFormulas(rep.fishFormulas) },
          { label: 'اجمالى البط', value: sumFormulas(rep.duckFormulas) },
          { label: 'اجمالى الحيوانات الاليفة', value: sumFormulas(rep.petsFormulas) },
        ],
        waterDeducted,
        cleaningExcluded,
      };
    } catch {
      return { totals: [], waterDeducted: 0, cleaningExcluded: 0 };
    }
  }, [sheet, metric, divisor, documentTitle]);
  const categoryTotals = categoryInfo.totals;

  const catDecimals = unit === 'ton' ? 3 : 2;
  const fmtCat = (n: number) =>
    n === 0
      ? '-'
      : n.toLocaleString('en-US', { minimumFractionDigits: catDecimals, maximumFractionDigits: catDecimals });

  const year = (report.detectedDate.match(/(\d{4})/) || [])[1] || new Date().getFullYear().toString();
  const title = `الصرف اليومي ${year} للكنترول`;
  const qtyHeader = `الكمية المصروفة (${unitLabel})`;
  const shownRows = hideEmpty ? report.rows.filter((r) => r.hasData || r.custom) : report.rows;
  const fmt = (n: number) =>
    n.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });

  const asSheet = (): SheetData => ({
    sheetName: 'الصرف_اليومي',
    isRtl: true,
    headers: ['Item', 'كود الصنف', 'كود درايف', qtyHeader, 'الصنف'],
    rows: shownRows.map((r) => [
      r.en,
      r.code,
      r.driveCode,
      r.hasData ? Number(r.qty.toFixed(decimals)) : '-',
      r.ar,
    ]),
    totalsRow: ['الإجمالي', '', '', Number(report.total.toFixed(decimals)), ''],
  });

  const rowsWithCategories = (): SheetData => {
    const base = asSheet();
    if (categoryTotals.length === 0) return base;
    return {
      ...base,
      rows: [
        ...base.rows,
        ['', '', '', '', ''],
        ...categoryTotals.map((c) => [c.label, '', '', Number(c.value.toFixed(catDecimals)), '']),
      ],
    };
  };

  const baseName = (fileName || 'daily_spend').replace(/\.[^/.]+$/, '');

  const handleCopy = async () => {
    const ok = await copyTableToClipboard(rowsWithCategories());
    if (ok) {
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    }
  };

  const nameExists = (en: string) => {
    const k = en.trim().toLowerCase();
    return (
      DAILY_SPEND_ITEMS.some((d) => d.en.toLowerCase() === k) ||
      customItems.some((c) => c.en.trim().toLowerCase() === k)
    );
  };

  const addItem = (item: DailySpendCustomItem): boolean => {
    const en = item.en.trim();
    if (!en) {
      setFormError('اكتب اسم الصنف أولاً');
      return false;
    }
    if (nameExists(en)) {
      setFormError('هذا الصنف موجود بالفعل في القائمة');
      return false;
    }
    const next = [
      ...customItems,
      {
        en,
        ar: (item.ar || '').trim(),
        code: item.code?.trim() || undefined,
        driveCode: item.driveCode?.trim() || undefined,
      },
    ];
    setCustomItems(next);
    persist(DAILY_SPEND_CUSTOM_ITEMS_KEY, next);
    setFormError('');
    return true;
  };

  const removeCustom = (en: string) => {
    const next = customItems.filter((c) => c.en !== en);
    setCustomItems(next);
    persist(DAILY_SPEND_CUSTOM_ITEMS_KEY, next);
  };

  const saveCode = (en: string, field: 'code' | 'driveCode', value: string, autoValue: string) => {
    const next: DailySpendCodeOverrides = { ...codeOverrides, [en]: { ...(codeOverrides[en] || {}) } };
    if (value === '') delete next[en][field];
    else next[en][field] = value;
    if (Object.keys(next[en]).length === 0) delete next[en];
    // A custom item keeps its own typed code, so an empty save there just clears the override
    void autoValue;
    setCodeOverrides(next);
    persist(DAILY_SPEND_CODE_OVERRIDES_KEY, next);
  };

  const submitForm = (e: React.FormEvent) => {
    e.preventDefault();
    if (addItem(form)) {
      setForm({ en: '', ar: '', code: '', driveCode: '' });
      setShowAddForm(false);
    }
  };

  const segBtn = (active: boolean) =>
    `px-3 py-1.5 text-xs font-bold transition-colors cursor-pointer ${
      active ? 'bg-slate-900 text-white' : 'bg-white text-slate-700 hover:bg-slate-50'
    }`;
  const actionBtn =
    'px-3 py-1.5 text-xs font-bold bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 rounded-lg flex items-center gap-1.5 transition-colors cursor-pointer';
  const inputCls =
    'w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-sm outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 bg-white';

  return (
    <div className="px-4 sm:px-6 py-4 space-y-4" dir="rtl">
      {/* Toolbar */}
      <div className="flex items-center gap-2.5 flex-wrap print:hidden">
        <div className="inline-flex rounded-lg border border-slate-200 overflow-hidden shadow-2xs">
          <button type="button" onClick={() => setUnit('ton')} className={segBtn(unit === 'ton')}>
            بالطن (÷ 1000)
          </button>
          <button type="button" onClick={() => setUnit('kg')} className={segBtn(unit === 'kg')}>
            بالكيلوجرام (كجم)
          </button>
        </div>

        <div className="inline-flex rounded-lg border border-slate-200 overflow-hidden shadow-2xs">
          <button type="button" onClick={() => setMetric('actual')} className={segBtn(metric === 'actual')}>
            الفعلي (Actual)
          </button>
          <button type="button" onClick={() => setMetric('target')} className={segBtn(metric === 'target')}>
            المستهدف (Target)
          </button>
        </div>

        <button
          type="button"
          onClick={() => {
            setShowAddForm((v) => !v);
            setFormError('');
          }}
          className="px-3.5 py-1.5 text-xs font-bold text-white bg-slate-900 hover:bg-slate-800 rounded-lg flex items-center gap-1.5 shadow-xs transition-colors cursor-pointer"
        >
          {showAddForm ? <X className="w-3.5 h-3.5" /> : <Plus className="w-3.5 h-3.5" />}
          <span>{showAddForm ? 'إلغاء' : 'إضافة صنف'}</span>
        </button>

        <button type="button" onClick={() => setHideEmpty((v) => !v)} className={actionBtn}>
          {hideEmpty ? <Eye className="w-3.5 h-3.5" /> : <EyeOff className="w-3.5 h-3.5" />}
          <span>{hideEmpty ? 'إظهار كل الأصناف' : 'إخفاء الأصناف الفارغة'}</span>
        </button>

        <div className="flex items-center gap-2 mr-auto flex-wrap">
          <button
            type="button"
            onClick={() =>
              exportDailySpendStyledExcel({
                rows: shownRows,
                total: report.total,
                decimals,
                title,
                qtyHeader,
                fileName: `${baseName}_daily_spend`,
                categoryTotals,
                categoryDecimals: catDecimals,
              })
            }
            className="px-3.5 py-1.5 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg flex items-center gap-1.5 shadow-xs transition-colors cursor-pointer"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>تصدير (.xlsx)</span>
          </button>
          <button type="button" onClick={() => exportToCsv(rowsWithCategories(), `${baseName}_daily_spend.csv`)} className={actionBtn}>
            <Download className="w-3.5 h-3.5" />
            <span>CSV</span>
          </button>
          <button type="button" onClick={handleCopy} className={actionBtn}>
            {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
            <span>{copied ? 'تم النسخ' : 'نسخ'}</span>
          </button>
          <button type="button" onClick={() => window.print()} className={actionBtn}>
            <Printer className="w-3.5 h-3.5" />
            <span>طباعة</span>
          </button>
        </div>
      </div>

      {/* Add item form */}
      {showAddForm && (
        <form
          onSubmit={submitForm}
          className="max-w-3xl mx-auto bg-slate-50 border border-slate-200 rounded-xl p-4 print:hidden"
        >
          <div className="text-sm font-extrabold text-slate-900 mb-3">إضافة صنف جديد للتقرير</div>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <label className="block">
              <span className="block text-xs font-bold text-slate-700 mb-1">
                اسم الصنف بالإنجليزي (زي ما هو في الملف) <span className="text-rose-600">*</span>
              </span>
              <input
                className={inputCls}
                dir="ltr"
                value={form.en}
                onChange={(e) => setForm({ ...form, en: e.target.value })}
                placeholder="مثال: Salt"
                autoFocus
              />
            </label>
            <label className="block">
              <span className="block text-xs font-bold text-slate-700 mb-1">الاسم بالعربي</span>
              <input
                className={inputCls}
                value={form.ar}
                onChange={(e) => setForm({ ...form, ar: e.target.value })}
                placeholder="مثال: ملح"
              />
            </label>
            <label className="block">
              <span className="block text-xs font-bold text-slate-700 mb-1">كود الصنف (JDE)</span>
              <input
                className={inputCls}
                dir="ltr"
                value={form.code}
                onChange={(e) => setForm({ ...form, code: e.target.value })}
                placeholder="اتركه فارغاً ليتحدد تلقائياً"
              />
            </label>
            <label className="block">
              <span className="block text-xs font-bold text-slate-700 mb-1">كود الدريف</span>
              <input
                className={inputCls}
                dir="ltr"
                value={form.driveCode}
                onChange={(e) => setForm({ ...form, driveCode: e.target.value })}
                placeholder="اتركه فارغاً ليتحدد تلقائياً"
              />
            </label>
          </div>
          {formError && <p className="text-xs font-bold text-rose-600 mt-2">{formError}</p>}
          <div className="flex items-center gap-2 mt-3">
            <button
              type="submit"
              className="px-4 py-1.5 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg flex items-center gap-1.5 cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>إضافة</span>
            </button>
            <span className="text-[11px] text-slate-500">
              لو الصنف موجود في الملف المرفوع، كميته هتظهر تلقائياً. والأكواد بتتجاب من كتالوج الدريف لو سبتها فاضية.
            </span>
          </div>
        </form>
      )}

      {/* Report */}
      <div
        id="daily-spend-report"
        className="max-w-4xl mx-auto bg-white border-2 border-slate-900 shadow-sm overflow-x-auto"
      >
        <table className="w-full border-collapse text-sm table-fixed min-w-[640px]">
          <colgroup>
            <col style={{ width: '30%' }} />
            <col style={{ width: '13%' }} />
            <col style={{ width: '12%' }} />
            <col style={{ width: '19%' }} />
            <col style={{ width: '22%' }} />
            <col style={{ width: '4%' }} className="print:hidden" />
          </colgroup>
          <thead>
            <tr className="bg-[#ffff00] text-slate-950">
              <th className="border border-slate-900 px-3 py-2.5 text-center font-extrabold leading-snug">{title}</th>
              <th className="border border-slate-900 px-2 py-2.5 text-center font-extrabold whitespace-nowrap">كود الصنف</th>
              <th className="border border-slate-900 px-2 py-2.5 text-center font-extrabold whitespace-nowrap">كود درايف</th>
              <th className="border border-slate-900 px-2 py-2.5 text-center font-extrabold leading-snug">{qtyHeader}</th>
              <th className="border border-slate-900 px-3 py-2.5 text-center font-extrabold">الصنف</th>
              <th className="border border-slate-900 w-8 print:hidden" aria-label="إجراءات" />
            </tr>
          </thead>
          <tbody>
            {shownRows.map((r) => {
              const manual = codeOverrides[r.en] || {};
              return (
                <tr key={`${r.index}-${r.en}`}>
                  <td className="border border-slate-900 bg-[#fdebd0] px-3 py-1.5 text-center font-bold text-slate-900 break-words" dir="ltr">
                    {r.en}
                  </td>
                  <td className="border border-slate-900 bg-white px-1 py-1 text-center">
                    <EditableCode
                      value={r.code}
                      isManual={manual.code !== undefined}
                      label="كود الصنف"
                      onSave={(v) => saveCode(r.en, 'code', v, r.code)}
                    />
                  </td>
                  <td className="border border-slate-900 bg-white px-1 py-1 text-center">
                    <EditableCode
                      value={r.driveCode}
                      isManual={manual.driveCode !== undefined}
                      label="كود الدريف"
                      onSave={(v) => saveCode(r.en, 'driveCode', v, r.driveCode)}
                    />
                  </td>
                  <td
                    className={`border border-slate-900 px-3 py-1.5 text-center font-mono tabular-nums font-extrabold ${
                      r.hasData ? 'bg-[#ffff00] text-slate-950' : 'bg-yellow-50 text-slate-400'
                    }`}
                    dir="ltr"
                  >
                    {r.hasData ? fmt(r.qty) : '-'}
                  </td>
                  <td className="border border-slate-900 bg-white px-3 py-1.5 text-center font-semibold text-slate-800 break-words">
                    {r.ar}
                  </td>
                  <td className="border border-slate-900 bg-white text-center print:hidden">
                    {r.custom ? (
                      <button
                        type="button"
                        onClick={() => removeCustom(r.en)}
                        className="p-1 text-rose-500 hover:text-rose-700 hover:bg-rose-50 rounded cursor-pointer"
                        title="حذف هذا الصنف المضاف"
                        aria-label={`حذف ${r.en}`}
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    ) : null}
                  </td>
                </tr>
              );
            })}
            <tr className="bg-emerald-100">
              <td className="border-2 border-slate-900 px-3 py-2 text-center font-extrabold text-slate-950">الإجمالي</td>
              <td className="border-2 border-slate-900" />
              <td className="border-2 border-slate-900" />
              <td className="border-2 border-slate-900 px-3 py-2 text-center font-mono tabular-nums font-extrabold text-slate-950" dir="ltr">
                {fmt(report.total)}
              </td>
              <td className="border-2 border-slate-900 px-3 py-2 text-center font-bold text-slate-700">{unitLabel}</td>
              <td className="border-2 border-slate-900 print:hidden" />
            </tr>
          </tbody>
        </table>

        {categoryTotals.length > 0 && (
          <table className="w-full border-collapse table-fixed min-w-[640px] mt-4 border-t-2 border-slate-900">
            <colgroup>
              <col style={{ width: '50%' }} />
              <col style={{ width: '50%' }} />
            </colgroup>
            <tbody>
              {categoryTotals.map((c) => (
                <tr key={c.label} className="h-11">
                  <td className="border border-slate-900 bg-[#00ffff] px-3 text-center text-sm font-extrabold text-slate-950">
                    {c.label}
                  </td>
                  <td
                    className="border border-slate-900 bg-white px-3 text-center font-mono tabular-nums text-base font-extrabold text-slate-950"
                    dir="ltr"
                  >
                    {fmtCat(c.value)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        {categoryTotals.length > 0 && (
          <p className="px-3 py-1.5 text-[11px] text-slate-600 border-t border-slate-300" dir="rtl">
            الإجماليات بعد خصم الماء ({fmtCat(categoryInfo.waterDeducted)} {unitLabel})
            {categoryInfo.cleaningExcluded > 0
              ? ` ودون مجموعة Cleaning (${fmtCat(categoryInfo.cleaningExcluded)} ${unitLabel})`
              : ''}
            .
          </p>
        )}
      </div>

      <p className="max-w-4xl mx-auto text-[11px] text-slate-500 flex items-center gap-1 print:hidden">
        <Pencil className="w-3 h-3" />
        <span>
          انقر على أي كود (الصنف أو الدريف) لتعديله. الأكواد المعدّلة يدوياً بتظهر باللون الأخضر، وتفريغ الخانة بيرجعها للكود التلقائي.
        </span>
      </p>

      {/* Reconciliation with the rest of the file */}
      <div className="max-w-4xl mx-auto text-xs text-slate-600 leading-6 print:hidden" dir="rtl">
        <p>
          عدد الأصناف في القائمة: <b>{report.rows.length}</b> · منها بكميات: <b>{report.rows.filter((r) => r.hasData).length}</b>
          {report.detectedDate ? <> · تاريخ التقرير: <b>{report.detectedDate}</b></> : null}
        </p>
        <p>
          إجمالي الملف كله: <b className="font-mono">{fmt(report.fileTotal)}</b> {unitLabel} · خارج هذه القائمة:{' '}
          <b className="font-mono">{fmt(report.otherTotal)}</b> {unitLabel} ({report.otherItems.length} صنف)
        </p>
        {report.otherItems.length > 0 && (
          <details className="mt-1">
            <summary className="cursor-pointer font-semibold text-slate-700">
              عرض الأصناف الأخرى غير المدرجة (اضغط + لإضافة أي صنف للتقرير)
            </summary>
            <ul className="mt-1 grid grid-cols-1 sm:grid-cols-2 gap-x-6" dir="ltr">
              {report.otherItems.map((o) => (
                <li key={o.name} className="flex items-center justify-between gap-3 border-b border-slate-100 py-0.5">
                  <span className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => addItem({ en: o.name, ar: '' })}
                      className="w-5 h-5 flex items-center justify-center rounded bg-emerald-600 text-white hover:bg-emerald-700 cursor-pointer"
                      title="إضافة هذا الصنف للتقرير"
                      aria-label={`إضافة ${o.name}`}
                    >
                      <Plus className="w-3 h-3" />
                    </button>
                    {o.name}
                  </span>
                  <span className="font-mono">{fmt(o.qty)}</span>
                </li>
              ))}
            </ul>
          </details>
        )}
      </div>
    </div>
  );
};
