import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { ChevronDown } from 'lucide-react';
import { t } from '../i18n';

const keepFocus = (e: React.MouseEvent) => e.preventDefault();

export function Btn({
  icon,
  label,
  title,
  onClick,
  active,
  disabled,
  big,
}: {
  icon?: ReactNode;
  label?: string;
  title?: string;
  onClick?: () => void;
  active?: boolean;
  disabled?: boolean;
  big?: boolean;
}) {
  return (
    <button
      className={`rb-btn${big ? ' big' : ''}${active ? ' active' : ''}`}
      title={title ?? label}
      disabled={disabled}
      onMouseDown={keepFocus}
      onClick={onClick}
    >
      {icon}
      {label && <span className="rb-label">{label}</span>}
    </button>
  );
}

export function Group({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="rb-group">
      <div className="rb-group-body">{children}</div>
      <div className="rb-group-label">{label}</div>
    </div>
  );
}

/** A button that opens a floating panel. The panel is position:fixed so the ribbon can scroll. */
export function Popover({
  icon,
  label,
  title,
  big,
  disabled,
  className,
  children,
}: {
  icon?: ReactNode;
  label?: string;
  title?: string;
  big?: boolean;
  disabled?: boolean;
  className?: string;
  children: (close: () => void) => ReactNode;
}) {
  const [open, setOpen] = useState(false);
  const [pos, setPos] = useState({ left: 0, top: 0 });
  const btnRef = useRef<HTMLButtonElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    if (!open) return;
    const b = btnRef.current!.getBoundingClientRect();
    const p = panelRef.current!.getBoundingClientRect();
    // Opens under the button, lined up with its start edge (the right edge in the Arabic interface).
    const start = document.documentElement.dir === 'rtl' ? b.right - p.width : b.left;
    setPos({ left: Math.max(4, Math.min(start, window.innerWidth - p.width - 8)), top: b.bottom + 4 });
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const down = (e: PointerEvent) => {
      const t = e.target as Node;
      if (!panelRef.current?.contains(t) && !btnRef.current?.contains(t)) setOpen(false);
    };
    const key = (e: KeyboardEvent) => e.key === 'Escape' && setOpen(false);
    window.addEventListener('pointerdown', down, true);
    window.addEventListener('keydown', key);
    return () => {
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('keydown', key);
    };
  }, [open]);

  return (
    <>
      <button
        ref={btnRef}
        className={`rb-btn${big ? ' big' : ''}${open ? ' active' : ''} ${className ?? ''}`}
        title={title ?? label}
        disabled={disabled}
        onMouseDown={keepFocus}
        onClick={() => setOpen((o) => !o)}
      >
        {icon}
        {label && <span className="rb-label">{label}</span>}
        <ChevronDown size={12} className="caret" />
      </button>
      {open && (
        <div ref={panelRef} className="menu popover" style={pos} onMouseDown={keepFocus}>
          {children(() => setOpen(false))}
        </div>
      )}
    </>
  );
}

// ---- Colour picker ----

function mix(hex: string, target: string, t: number) {
  const p = (h: string) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const a = p(hex);
  const b = p(target);
  return '#' + a.map((v, i) => Math.round(v + (b[i] - v) * t).toString(16).padStart(2, '0')).join('');
}

const BASE = ['#000000', '#ffffff', '#1f2937', '#2563eb', '#dc2626', '#ea580c', '#ca8a04', '#16a34a', '#0891b2', '#7c3aed'];
const PALETTE: string[][] = [
  BASE,
  ...[0.8, 0.6, 0.4].map((t, r) => BASE.map((c) => (c === '#ffffff' ? mix(c, '#000000', 0.05 + r * 0.1) : mix(c, '#ffffff', t)))),
  ...[0.25, 0.5].map((t, r) => BASE.map((c) => (c === '#ffffff' ? mix(c, '#000000', 0.35 + r * 0.15) : mix(c, '#000000', t)))),
];

export function ColorPicker({
  icon,
  title,
  value,
  onChange,
  allowNone,
  noneLabel = 'No Color',
  disabled,
}: {
  icon: ReactNode;
  title: string;
  value: string | null | undefined;
  onChange: (c: string | null) => void;
  allowNone?: boolean;
  noneLabel?: string;
  disabled?: boolean;
}) {
  return (
    <Popover
      title={title}
      disabled={disabled}
      className="color-btn"
      icon={
        <span className="color-icon">
          {icon}
          <span className="color-bar" style={{ background: value ?? 'transparent' }} />
        </span>
      }
    >
      {(close) => (
        <div className="palette">
          {allowNone && (
            <button className="menu-item" onClick={() => (onChange(null), close())}>
              <span className="swatch none" /> {noneLabel}
            </button>
          )}
          <div className="palette-grid">
            {PALETTE.flatMap((row, r) =>
              row.map((c, i) => (
                <button
                  key={`${r}-${i}`}
                  className={`swatch${value?.toLowerCase() === c ? ' selected' : ''}`}
                  style={{ background: c }}
                  title={c}
                  onClick={() => (onChange(c), close())}
                />
              )),
            )}
          </div>
          <label className="menu-item custom-color">
            <input
              type="color"
              value={value && /^#[0-9a-f]{6}$/i.test(value) ? value : '#000000'}
              onChange={(e) => onChange(e.target.value)}
            />
            {t('More Colors…')}
          </label>
        </div>
      )}
    </Popover>
  );
}
