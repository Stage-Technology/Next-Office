import { useCallback, useEffect, useRef, useState } from 'react';
import { Plus } from 'lucide-react';
import { useStore } from '../store';
import { slideSize } from '../types';
import { SlideView } from './SlideView';
import { ContextMenu } from './ContextMenu';
import { copyCurrent, pasteFromButton } from '../io';
import { digits, t } from '../i18n';

const THUMB_W = 168;
/** Portrait slides get narrower thumbnails so the list stays scannable. */
const THUMB_MAX_H = 170;

/** Cut or copy the selected slides to the system clipboard (the slide panel's own menu). */
function copyToClipboard(cut = false) {
  useStore.getState().setFocusArea('panel');
  const text = copyCurrent(cut);
  if (text) navigator.clipboard?.writeText(text).catch(() => {});
}

export function SlidePanel() {
  const slides = useStore((s) => s.doc.slides);
  const size = useStore((s) => slideSize(s.doc));
  const currentId = useStore((s) => s.currentSlideId);
  const selected = useStore((s) => s.selectedSlideIds);
  const st = useStore.getState();
  const [drop, setDrop] = useState<number | null>(null);
  const [menu, setMenu] = useState<{ x: number; y: number; id: string } | null>(null);
  const dragFrom = useRef<number | null>(null);
  const listRef = useRef<HTMLDivElement>(null);
  const closeMenu = useCallback(() => setMenu(null), []);

  // Keep the current slide visible when navigating with the keyboard.
  useEffect(() => {
    listRef.current?.querySelector('.thumb-row.active')?.scrollIntoView({ block: 'nearest' });
  }, [currentId]);

  const onDragOver = (e: React.DragEvent, i: number) => {
    if (dragFrom.current === null) return;
    e.preventDefault();
    const r = e.currentTarget.getBoundingClientRect();
    setDrop(e.clientY < r.top + r.height / 2 ? i : i + 1);
  };

  const onDrop = () => {
    const from = dragFrom.current;
    if (from !== null && drop !== null) {
      const to = drop > from ? drop - 1 : drop;
      st.moveSlide(from, to);
    }
    dragFrom.current = null;
    setDrop(null);
  };

  return (
    <aside className="slide-panel" onPointerDown={() => st.setFocusArea('panel')}>
      <div className="thumb-list" ref={listRef} onDragLeave={(e) => !e.currentTarget.contains(e.relatedTarget as Node) && setDrop(null)}>
        {slides.map((slide, i) => (
          <div
            key={slide.id}
            className={`thumb-row${slide.id === currentId ? ' active' : ''}${selected.includes(slide.id) ? ' selected' : ''}${
              drop === i ? ' drop-before' : ''
            }${
              drop === i + 1 && i === slides.length - 1 ? ' drop-after' : ''
            }`}
            draggable
            onDragStart={(e) => {
              dragFrom.current = i;
              e.dataTransfer.effectAllowed = 'move';
            }}
            onDragOver={(e) => onDragOver(e, i)}
            onDrop={onDrop}
            onDragEnd={() => {
              dragFrom.current = null;
              setDrop(null);
            }}
            onClick={(e) => st.selectSlide(slide.id, e.shiftKey ? 'range' : e.ctrlKey || e.metaKey ? 'toggle' : undefined)}
            onDoubleClick={() => st.present(i)}
            onContextMenu={(e) => {
              e.preventDefault();
              // Right-clicking one of several selected slides keeps the selection, so the menu acts on all of them.
              if (slide.id !== currentId && !selected.includes(slide.id)) st.selectSlide(slide.id);
              setMenu({ x: e.clientX, y: e.clientY, id: slide.id });
            }}
          >
            <span className="thumb-num">{digits(i + 1)}</span>
            <div className="thumb">
              <SlideView slide={slide} width={Math.min(THUMB_W, Math.round((THUMB_MAX_H * size.w) / size.h))} size={size} />
            </div>
          </div>
        ))}
      </div>
      <button className="add-slide" onClick={() => st.addSlide()} title={t('New Slide (Ctrl+M)')}>
        <Plus size={16} /> {t('New slide')}
      </button>
      {menu && (
        <ContextMenu
          x={menu.x}
          y={menu.y}
          onClose={closeMenu}
          items={[
            { label: 'Cut', shortcut: 'Ctrl+X', action: () => copyToClipboard(true) },
            { label: 'Copy', shortcut: 'Ctrl+C', action: () => copyToClipboard() },
            { label: 'Paste', shortcut: 'Ctrl+V', action: () => pasteFromButton() },
            { label: 'Paste (Keep Source Formatting)', action: () => pasteFromButton({ keepSource: true }) },
            { divider: true },
            { label: 'New Slide', shortcut: 'Ctrl+M', action: () => st.addSlide() },
            { label: selected.length ? 'Duplicate Slides' : 'Duplicate Slide', shortcut: 'Ctrl+D', action: () => st.duplicateSlide(menu.id) },
            { label: selected.length ? 'Delete Slides' : 'Delete Slide', shortcut: 'Del', action: () => st.deleteSlide(menu.id) },
            { divider: true },
            { label: 'Present From Here', action: () => st.present(slides.findIndex((s) => s.id === menu.id)) },
          ]}
        />
      )}
    </aside>
  );
}
