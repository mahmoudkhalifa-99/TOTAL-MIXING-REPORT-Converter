import React, { useMemo, useState } from 'react';
import { Download, Copy, Check, FileSpreadsheet, Printer, EyeOff, Eye } from 'lucide-react';
import { SheetData } from '../types/converter';
import { calculateDailySpend, DAILY_SPEND_ITEMS } from '../utils/dailySpendCalculator';
import { exportDailySpendStyledExcel } from '../utils/styledExcelExport';
import { exportToCsv, copyTableToClipboard } from '../utils/excelExport';

interface DailySpendViewProps {
  sheet: SheetData;
  fileName?: string;
  documentTitle?: string;
}

type UnitMode = 'ton' | 'kg';
type Metric = 'actual' | 'target';

export const DailySpendView: React.FC<DailySpendViewProps> = ({ sheet, fileName, documentTitle }) => {
  const [unit, setUnit] = useState<UnitMode>('ton');
  const [metric, setMetric] = useState<Metric>('actual');
  const [hideEmpty, setHideEmpty] = useState<boolean>(false);
  const [copied, setCopied] = useState<boolean>(false);

  const divisor = unit === 'ton' ? 1000 : 1;
  const decimals = unit === 'ton' ? 4 : 2;
  const unitLabel = unit === 'ton' ? 'طن' : 'كجم';

  const report = useMemo(
    () => calculateDailySpend(sheet, metric, divisor, documentTitle),
    [sheet, metric, divisor, documentTitle]
  );

  const year = (report.detectedDate.match(/(\d{4})/) || [])[1] || new Date().getFullYear().toString();
  const title = `الصرف اليومي ${year} للكنترول`;
  const qtyHeader = `الكمية المصروفة (${unitLabel})`;
  const shownRows = hideEmpty ? report.rows.filter((r) => r.hasData) : report.rows;
  const fmt = (n: number) =>
    n.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });

  const asSheet = (): SheetData => ({
    sheetName: 'الصرف_اليومي',
    isRtl: true,
    headers: ['Item', qtyHeader, 'الصنف'],
    rows: shownRows.map((r) => [r.en, r.hasData ? Number(r.qty.toFixed(decimals)) : '-', r.ar]),
    totalsRow: ['الإجمالي', Number(report.total.toFixed(decimals)), ''],
  });

  const baseName = (fileName || 'daily_spend').replace(/\.[^/.]+$/, '');

  const handleCopy = async () => {
    const ok = await copyTableToClipboard(asSheet());
    if (ok) {
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    }
  };

  const segBtn = (active: boolean) =>
    `px-3 py-1.5 text-xs font-bold transition-colors cursor-pointer ${
      active ? 'bg-slate-900 text-white' : 'bg-white text-slate-700 hover:bg-slate-50'
    }`;
  const actionBtn =
    'px-3 py-1.5 text-xs font-bold bg-white text-slate-700 border border-slate-200 hover:bg-slate-50 rounded-lg flex items-center gap-1.5 transition-colors cursor-pointer';

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
              })
            }
            className="px-3.5 py-1.5 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg flex items-center gap-1.5 shadow-xs transition-colors cursor-pointer"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>تصدير (.xlsx)</span>
          </button>
          <button type="button" onClick={() => exportToCsv(asSheet(), `${baseName}_daily_spend.csv`)} className={actionBtn}>
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

      {/* Report */}
      <div className="max-w-2xl mx-auto bg-white border-2 border-slate-900 shadow-sm">
        <table className="w-full border-collapse text-sm">
          <thead>
            <tr className="bg-[#ffff00] text-slate-950">
              <th className="border border-slate-900 px-3 py-2.5 text-center font-extrabold">{title}</th>
              <th className="border border-slate-900 px-3 py-2.5 text-center font-extrabold">{qtyHeader}</th>
              <th className="border border-slate-900 px-3 py-2.5 text-center font-extrabold">الصنف</th>
            </tr>
          </thead>
          <tbody>
            {shownRows.map((r) => (
              <tr key={r.index}>
                <td className="border border-slate-900 bg-[#fdebd0] px-3 py-1.5 text-center font-bold text-slate-900" dir="ltr">
                  {r.en}
                </td>
                <td
                  className={`border border-slate-900 bg-[#ffff00] px-3 py-1.5 text-center font-mono font-extrabold ${
                    r.hasData ? 'text-slate-950' : 'text-slate-500'
                  }`}
                  dir="ltr"
                >
                  {r.hasData ? fmt(r.qty) : '-'}
                </td>
                <td className="border border-slate-900 bg-white px-3 py-1.5 text-center font-semibold text-slate-800">
                  {r.ar}
                </td>
              </tr>
            ))}
            <tr className="bg-emerald-100">
              <td className="border-2 border-slate-900 px-3 py-2 text-center font-extrabold text-slate-950">
                الإجمالي
              </td>
              <td className="border-2 border-slate-900 px-3 py-2 text-center font-mono font-extrabold text-slate-950" dir="ltr">
                {fmt(report.total)}
              </td>
              <td className="border-2 border-slate-900 px-3 py-2 text-center font-bold text-slate-700">{unitLabel}</td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* Reconciliation with the rest of the file */}
      <div className="max-w-2xl mx-auto text-xs text-slate-600 leading-6 print:hidden" dir="rtl">
        <p>
          عدد الأصناف في القائمة: <b>{DAILY_SPEND_ITEMS.length}</b> · منها بكميات:{' '}
          <b>{report.rows.filter((r) => r.hasData).length}</b>
          {report.detectedDate ? <> · تاريخ التقرير: <b>{report.detectedDate}</b></> : null}
        </p>
        <p>
          إجمالي الملف كله: <b className="font-mono">{fmt(report.fileTotal)}</b> {unitLabel} · خارج هذه القائمة:{' '}
          <b className="font-mono">{fmt(report.otherTotal)}</b> {unitLabel} ({report.otherItems.length} صنف)
        </p>
        {report.otherItems.length > 0 && (
          <details className="mt-1">
            <summary className="cursor-pointer font-semibold text-slate-700">عرض الأصناف الأخرى غير المدرجة</summary>
            <ul className="mt-1 grid grid-cols-1 sm:grid-cols-2 gap-x-6" dir="ltr">
              {report.otherItems.map((o) => (
                <li key={o.name} className="flex justify-between gap-3 border-b border-slate-100 py-0.5">
                  <span>{o.name}</span>
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
