import React, { useEffect, useRef, useState } from 'react';
import { Minus, Plus, ZoomIn, RotateCcw } from 'lucide-react';

const ZOOM_STORAGE_KEY = 'app_zoom_percent_v1';
const MIN_ZOOM = 50;
const MAX_ZOOM = 200;
const STEP = 10;
const PRESETS = [70, 80, 90, 100, 110, 125, 150];

function readStoredZoom(): number {
  try {
    const raw = window.localStorage.getItem(ZOOM_STORAGE_KEY);
    const n = raw ? parseInt(raw, 10) : 100;
    if (Number.isFinite(n) && n >= MIN_ZOOM && n <= MAX_ZOOM) return n;
  } catch {
    // storage may be unavailable; fall back to 100%
  }
  return 100;
}

/**
 * Floating control that zooms the WHOLE app (header, tables, modals) and remembers the choice.
 * - Sits bottom-left, always reachable, never grows/shrinks with the zoom itself.
 * - Click the percentage for quick presets and a slider.
 * - Hidden when printing.
 */
export const AppZoomControl: React.FC = () => {
  const [zoom, setZoom] = useState<number>(() => readStoredZoom());
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  // Apply zoom to the root element + expose a "real" 1vh that stays correct under CSS zoom
  useEffect(() => {
    const root = document.documentElement;
    const z = zoom / 100;
    (root.style as any).zoom = zoom === 100 ? '' : String(z);
    root.style.setProperty('--app-zoom', String(z));

    const updateVh = () => {
      root.style.setProperty('--app-vh', `${window.innerHeight / z / 100}px`);
    };
    updateVh();
    window.addEventListener('resize', updateVh);

    try {
      window.localStorage.setItem(ZOOM_STORAGE_KEY, String(zoom));
    } catch {
      // ignore
    }
    return () => window.removeEventListener('resize', updateVh);
  }, [zoom]);

  // Close the presets popover on outside click / Escape
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
    };
  }, [open]);

  const clamp = (n: number) => Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, n));
  const change = (delta: number) => setZoom((z) => clamp(Math.round((z + delta) / STEP) * STEP));

  const iconBtn =
    'w-8 h-8 flex items-center justify-center rounded-full text-slate-700 hover:bg-slate-100 active:bg-slate-200 disabled:opacity-30 disabled:hover:bg-transparent transition-colors cursor-pointer';

  return (
    <div
      ref={rootRef}
      dir="ltr"
      // Counter-scale so the control keeps the same on-screen size at every zoom level
      style={{ zoom: 1 / (zoom / 100) } as React.CSSProperties}
      className="fixed bottom-4 left-4 z-50 print:hidden select-none"
    >
      {open && (
        <div
          className="mb-2 w-60 bg-white border border-slate-200 rounded-2xl shadow-xl p-3"
          role="dialog"
          aria-label="خيارات تكبير وتصغير التطبيق"
        >
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-bold text-slate-800" dir="rtl">
              حجم التطبيق
            </span>
            <button
              type="button"
              onClick={() => setZoom(100)}
              disabled={zoom === 100}
              className="flex items-center gap-1 text-[11px] font-semibold text-emerald-700 hover:text-emerald-800 disabled:text-slate-400 disabled:cursor-default cursor-pointer"
            >
              <RotateCcw className="w-3 h-3" />
              <span dir="rtl">الحجم الطبيعي</span>
            </button>
          </div>

          <input
            type="range"
            min={MIN_ZOOM}
            max={MAX_ZOOM}
            step={STEP}
            value={zoom}
            onChange={(e) => setZoom(clamp(parseInt(e.target.value, 10)))}
            className="w-full accent-emerald-600 cursor-pointer"
            aria-label="نسبة التكبير"
          />
          <div className="flex justify-between text-[10px] font-mono text-slate-400 mb-2">
            <span>{MIN_ZOOM}%</span>
            <span>{MAX_ZOOM}%</span>
          </div>

          <div className="grid grid-cols-4 gap-1.5">
            {PRESETS.map((p) => (
              <button
                key={p}
                type="button"
                onClick={() => setZoom(p)}
                className={`py-1.5 text-xs font-mono font-bold rounded-lg border transition-colors cursor-pointer ${
                  zoom === p
                    ? 'bg-emerald-600 text-white border-emerald-600'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                }`}
              >
                {p}%
              </button>
            ))}
          </div>
        </div>
      )}

      <div className="flex items-center gap-0.5 bg-white/95 backdrop-blur border border-slate-200 rounded-full shadow-lg p-1">
        <button
          type="button"
          onClick={() => change(-STEP)}
          disabled={zoom <= MIN_ZOOM}
          className={iconBtn}
          aria-label="تصغير التطبيق"
          title="تصغير"
        >
          <Minus className="w-4 h-4" />
        </button>

        <button
          type="button"
          onClick={() => setOpen((o) => !o)}
          className="h-8 min-w-[58px] px-2 flex items-center justify-center gap-1 rounded-full text-xs font-mono font-bold text-slate-800 hover:bg-slate-100 transition-colors cursor-pointer"
          aria-expanded={open}
          title="خيارات الحجم"
        >
          <ZoomIn className="w-3.5 h-3.5 text-emerald-600" />
          <span>{zoom}%</span>
        </button>

        <button
          type="button"
          onClick={() => change(STEP)}
          disabled={zoom >= MAX_ZOOM}
          className={iconBtn}
          aria-label="تكبير التطبيق"
          title="تكبير"
        >
          <Plus className="w-4 h-4" />
        </button>
      </div>
    </div>
  );
};
