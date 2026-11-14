import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { t } from '../i18n';

export type MenuItem =
  | { label: string; shortcut?: string; disabled?: boolean; action: () => void }
  | { divider: true };

export function ContextMenu({ x, y, items, onClose }: { x: number; y: number; items: MenuItem[]; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ x, y });

  // Keep the menu on-screen.
  useLayoutEffect(() => {
    const r = ref.current!.getBoundingClientRect();
    // In the Arabic interface the menu opens to the left of the pointer, as in right-to-left Windows.
    const left = document.documentElement.dir === 'rtl' ? Math.max(4, x - r.width) : x;
    setPos({ x: Math.min(left, window.innerWidth - r.width - 4), y: Math.min(y, window.innerHeight - r.height - 4) });
  }, [x, y]);

  useEffect(() => {
    const down = (e: PointerEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose();
    };
    const key = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    window.addEventListener('pointerdown', down, true);
    window.addEventListener('keydown', key);
    window.addEventListener('blur', onClose);
    return () => {
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('keydown', key);
      window.removeEventListener('blur', onClose);
    };
  }, [onClose]);

  return (
    <div
      ref={ref}
      className="menu context-menu"
      style={{ left: pos.x, top: pos.y }}
      onPointerDown={(e) => e.stopPropagation()}
      onContextMenu={(e) => {
        e.preventDefault();
        e.stopPropagation();
      }}
    >
      {items.map((it, i) =>
        'divider' in it ? (
          <div key={i} className="menu-divider" />
        ) : (
          <button
            key={i}
            className="menu-item"
            disabled={it.disabled}
            onClick={() => {
              onClose();
              it.action();
            }}
          >
            <span>{t(it.label)}</span>
            {it.shortcut && <span className="menu-shortcut">{it.shortcut}</span>}
          </button>
        ),
      )}
    </div>
  );
}
