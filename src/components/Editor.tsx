import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { currentSlideOf, useCurrentSlide, useStore } from '../store';
import { hasTextBody, isLineShape, isMediaElement, slideSize, type Paragraph, type SlideElement } from '../types';
import { useBackgroundCss } from './MediaView';
import { animNumbers } from '../anim';
import { SurfaceContext } from './ChartView';
import { boundsOf, intersects, NO_GUIDES, rectFromPoints, snapMove, type Box, type Guides } from '../geometry';
import { ElementView } from './ElementView';
import { ContextMenu, type MenuItem } from './ContextMenu';
import { insertMediaFile, insertTextBox, pasteFromButton, copyToSystemClipboard, openDroppedDeck, isPowerPointFile } from '../io';
import { digits, t } from '../i18n';
import { checkboxAt, toggleParaChecked } from '../richtext';

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';
const ALL_HANDLES: Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const LINE_HANDLES: Handle[] = ['w', 'e'];

const get = useStore.getState;

function drag(onMove: (e: PointerEvent) => void, onUp?: (e: PointerEvent) => void) {
  const move = (e: PointerEvent) => onMove(e);
  const up = (e: PointerEvent) => {
    window.removeEventListener('pointermove', move);
    window.removeEventListener('pointerup', up);
    window.removeEventListener('pointercancel', up);
    onUp?.(e);
  };
  window.addEventListener('pointermove', move);
  window.addEventListener('pointerup', up);
  window.addEventListener('pointercancel', up);
}

export function Editor() {
  const slide = useCurrentSlide();
  const selectedIds = useStore((s) => s.selectedIds);
  const editingId = useStore((s) => s.editingId);
  const zoom = useStore((s) => s.zoom);
  const showAnimBadges = useStore((s) => s.showAnimBadges);
  const size = useStore((s) => slideSize(s.doc));
  const sizeRef = useRef(size);
  const background = useBackgroundCss(slide.background);
  const animNums = useMemo(() => animNumbers(slide.animations ?? [], slide.elements), [slide.animations, slide.elements]);
  sizeRef.current = size;

  const wrapRef = useRef<HTMLDivElement>(null);
  const stageRef = useRef<HTMLDivElement>(null);
  const [fit, setFit] = useState(0.8);
  const [guides, setGuides] = useState<Guides>(NO_GUIDES);
  const [marquee, setMarquee] = useState<Box | null>(null);
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  const editSession = useRef<string | null>(null);
  /** When a checklist box was last clicked, so a quick second click doesn't also start editing. */
  const tickedAt = useRef(0);

  useLayoutEffect(() => {
    const el = wrapRef.current!;
    const refit = () => {
      const r = el.getBoundingClientRect();
      setFit(Math.max(0.05, Math.min((r.width - 64) / size.w, (r.height - 64) / size.h)));
    };
    refit();
    const ro = new ResizeObserver(refit);
    ro.observe(el);
    return () => ro.disconnect();
  }, [size.w, size.h]);

  useEffect(() => {
    editSession.current = null;
  }, [editingId]);

  const scale = zoom === 'fit' ? fit : zoom;
  const scaleRef = useRef(scale);
  scaleRef.current = scale;

  const toLogical = useCallback((e: { clientX: number; clientY: number }) => {
    const r = stageRef.current!.getBoundingClientRect();
    return { x: (e.clientX - r.left) / scaleRef.current, y: (e.clientY - r.top) / scaleRef.current };
  }, []);

  // ---- Element interactions (stable callbacks so memoised ElementViews don't re-render) ----

  const onElementDown = useCallback(
    (e: React.PointerEvent, el: SlideElement) => {
      const st = get();
      st.setFocusArea('canvas');
      if (st.editingId === el.id) return; // let the textarea handle it
      e.stopPropagation();
      if (e.button === 2) {
        if (!st.selectedIds.includes(el.id)) st.setSelection([el.id]);
        return;
      }
      if (e.button !== 0) return;
      e.preventDefault();
      // A click on a checklist box of the selected text box ticks it instead of starting a drag.
      const alone = st.selectedIds.length === 1 && st.selectedIds[0] === el.id;
      const hit = alone && !e.shiftKey && !e.ctrlKey && !e.metaKey ? checkboxAt(e.target, e.clientX, e.clientY) : null;
      if (hit) {
        tickedAt.current = Date.now();
        return st.updateElements([el.id], (d) => toggleParaChecked(d, hit.index));
      }
      if (st.editingId) {
        (document.activeElement as HTMLElement | null)?.blur();
        st.setEditing(null);
      }

      let sel = get().selectedIds;
      if (e.shiftKey || e.ctrlKey || e.metaKey) {
        if (sel.includes(el.id)) return st.setSelection(sel.filter((i) => i !== el.id));
        sel = [...sel, el.id];
      } else if (!sel.includes(el.id)) sel = [el.id];
      st.setSelection(sel);

      const current = currentSlideOf(get());
      const moving = current.elements.filter((x) => sel.includes(x.id));
      const others = current.elements.filter((x) => !sel.includes(x.id));
      const starts = new Map(moving.map((m) => [m.id, { x: m.x, y: m.y }]));
      const box = boundsOf(moving);
      const p0 = toLogical(e);
      let moved = false;

      drag(
        (ev) => {
          const p = toLogical(ev);
          let dx = p.x - p0.x;
          let dy = p.y - p0.y;
          if (!moved) {
            if (Math.hypot(dx, dy) * scaleRef.current < 3) return;
            moved = true;
            st.checkpoint();
          }
          if (ev.shiftKey) {
            if (Math.abs(dx) > Math.abs(dy)) dy = 0;
            else dx = 0;
          }
          let g = NO_GUIDES;
          if (!ev.altKey) {
            const snap = snapMove({ ...box, x: box.x + dx, y: box.y + dy }, others, 6 / scaleRef.current, sizeRef.current);
            dx += snap.dx;
            dy += snap.dy;
            g = snap.guides;
          }
          setGuides(g);
          st.updateElements(
            sel,
            (d) => {
              const s0 = starts.get(d.id)!;
              d.x = Math.round(s0.x + dx);
              d.y = Math.round(s0.y + dy);
            },
            { history: false },
          );
        },
        () => setGuides(NO_GUIDES),
      );
    },
    [toLogical],
  );

  const onElementDoubleClick = useCallback((el: SlideElement) => {
    if (Date.now() - tickedAt.current < 600) return;
    const st = get();
    st.setSelection([el.id]);
    if (hasTextBody(el)) st.setEditing(el.id);
    else if (el.type === 'chart') st.setChartEditor(el.id);
  }, []);

  const onTextChange = useCallback((el: SlideElement, paras: Paragraph[]) => {
    const st = get();
    if (editSession.current !== el.id) {
      st.checkpoint(); // one undo step per editing session
      editSession.current = el.id;
    }
    st.updateElements([el.id], (d) => void (d.paras = paras), { history: false });
  }, []);

  const onAutoHeight = useCallback((el: SlideElement, h: number) => {
    const clean = get().doc === get().savedDoc;
    get().updateElements([el.id], (d) => void (d.h = h), { history: false });
    // Growing a box to fit its text on screen is not an edit: a document without changes stays that way.
    if (clean) useStore.setState({ savedDoc: get().doc });
  }, []);

  const onEditEnd = useCallback(() => get().setEditing(null), []);

  // ---- Resize / rotate ----

  const onResizeDown = (e: React.PointerEvent, el: SlideElement, handle: Handle) => {
    e.stopPropagation();
    e.preventDefault();
    const st = get();
    const e0 = { ...el };
    const p0 = toLogical(e);
    const rad = (e0.rotation * Math.PI) / 180;
    const cos = Math.cos(rad);
    const sin = Math.sin(rad);
    const hx = handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0;
    const hy = handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0;
    const line = isLineShape(e0);
    let moved = false;

    drag((ev) => {
      const p = toLogical(ev);
      const dx = p.x - p0.x;
      const dy = p.y - p0.y;
      if (!moved) {
        moved = true;
        st.checkpoint();
      }
      // Pointer delta in the element's rotated frame.
      const lx = dx * cos + dy * sin;
      const ly = -dx * sin + dy * cos;
      let w = e0.w + hx * lx;
      let h = e0.h + hy * ly;
      const lockDefault = isMediaElement(e0);
      if (hx && hy && lockDefault !== ev.shiftKey) {
        const ratio = e0.w / e0.h;
        if (w / h > ratio) h = w / ratio;
        else w = h * ratio;
      }
      w = Math.max(10, w);
      h = Math.max(line ? e0.h : 10, h);
      // Keep the opposite edge anchored: shift the centre by half the growth, in world space.
      const cxl = (hx * (w - e0.w)) / 2;
      const cyl = (hy * (h - e0.h)) / 2;
      const cx = e0.x + e0.w / 2 + cxl * cos - cyl * sin;
      const cy = e0.y + e0.h / 2 + cxl * sin + cyl * cos;
      st.updateElements(
        [e0.id],
        (d) => {
          d.w = Math.round(w);
          d.h = Math.round(h);
          d.x = Math.round(cx - w / 2);
          d.y = Math.round(cy - h / 2);
        },
        { history: false },
      );
    });
  };

  const onRotateDown = (e: React.PointerEvent, el: SlideElement) => {
    e.stopPropagation();
    e.preventDefault();
    const st = get();
    const cx = el.x + el.w / 2;
    const cy = el.y + el.h / 2;
    let moved = false;
    drag((ev) => {
      if (!moved) {
        moved = true;
        st.checkpoint();
      }
      const p = toLogical(ev);
      let a = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
      if (ev.shiftKey) a = Math.round(a / 15) * 15;
      else {
        const near = Math.round(a / 90) * 90;
        if (Math.abs(a - near) < 4) a = near;
      }
      a = ((Math.round(a) % 360) + 360) % 360;
      st.updateElements([el.id], (d) => void (d.rotation = a), { history: false });
    });
  };

  // ---- Background: marquee selection ----

  const onWrapDown = (e: React.PointerEvent) => {
    if (e.button !== 0) return;
    const st = get();
    st.setFocusArea('canvas');
    if (st.editingId) st.setEditing(null);
    const additive = e.shiftKey || e.ctrlKey || e.metaKey;
    const base = additive ? st.selectedIds : [];
    if (!additive) st.setSelection([]);
    const p0 = toLogical(e);
    drag(
      (ev) => {
        const box = rectFromPoints(p0, toLogical(ev));
        setMarquee(box);
        const hit = currentSlideOf(get()).elements.filter((x) => intersects(box, x)).map((x) => x.id);
        st.setSelection([...new Set([...base, ...hit])]);
      },
      () => setMarquee(null),
    );
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    const at = toLogical(e);
    const files = Array.from(e.dataTransfer.files);
    const deck = files.find(isPowerPointFile);
    if (deck) return void openDroppedDeck(deck);
    files.forEach((f) => insertMediaFile(f, at));
  };

  const closeMenu = useCallback(() => setMenu(null), []);

  const menuItems = (): MenuItem[] => {
    const st = get();
    const has = st.selectedIds.length > 0;
    const single = currentSlideOf(st).elements.find((x) => x.id === st.selectedIds[0]);
    const items: MenuItem[] = [];
    if (has) {
      items.push(
        {
          label: 'Cut',
          shortcut: 'Ctrl+X',
          action: () => {
            const els = st.copySelected();
            if (els) copyToSystemClipboard(els);
            st.deleteSelected();
          },
        },
        {
          label: 'Copy',
          shortcut: 'Ctrl+C',
          action: () => {
            const els = st.copySelected();
            if (els) copyToSystemClipboard(els);
          },
        },
      );
    }
    items.push({ label: 'Paste', shortcut: 'Ctrl+V', action: pasteFromButton });
    if (has) {
      items.push(
        { label: 'Duplicate', shortcut: 'Ctrl+D', action: st.duplicateSelected },
        { label: 'Delete', shortcut: 'Del', action: st.deleteSelected },
        { divider: true },
        { label: 'Bring to Front', action: () => st.arrange('front') },
        { label: 'Bring Forward', action: () => st.arrange('forward') },
        { label: 'Send Backward', action: () => st.arrange('backward') },
        { label: 'Send to Back', action: () => st.arrange('back') },
      );
      if (st.selectedIds.length === 1 && single && hasTextBody(single)) {
        items.push({ divider: true }, { label: 'Edit Text', shortcut: 'Enter', action: () => st.setEditing(single.id) });
      }
      if (st.selectedIds.length === 1 && single?.type === 'chart') {
        items.push({ divider: true }, { label: 'Edit Data…', shortcut: 'Enter', action: () => st.setChartEditor(single.id) });
      }
    } else {
      items.push(
        { label: 'Select All', shortcut: 'Ctrl+A', action: st.selectAll },
        { divider: true },
        { label: 'Insert Text Box', action: () => insertTextBox() },
        { label: 'New Slide', shortcut: 'Ctrl+M', action: () => st.addSlide() },
        { label: 'Duplicate Slide', action: () => st.duplicateSlide() },
      );
    }
    return items;
  };

  const selEls = slide.elements.filter((e) => selectedIds.includes(e.id));
  const single = selEls.length === 1 ? selEls[0] : null;

  return (
    <div
      className="stage-wrap"
      ref={wrapRef}
      onPointerDown={onWrapDown}
      onContextMenu={(e) => {
        e.preventDefault();
        setMenu({ x: e.clientX, y: e.clientY });
      }}
      onDragOver={(e) => e.preventDefault()}
      onDrop={onDrop}
    >
      <div className="stage" ref={stageRef} style={{ width: size.w * scale, height: size.h * scale }}>
        <div
          className="slide-surface editing-surface"
          style={{ width: size.w, height: size.h, transform: `scale(${scale})`, background }}
        >
          <SurfaceContext.Provider value={slide.background.color}>
            {slide.elements.map((el) => (
              <ElementView
                key={el.id}
                el={el}
                mode="edit"
                editing={editingId === el.id}
                onPointerDown={onElementDown}
                onDoubleClick={onElementDoubleClick}
                onTextChange={onTextChange}
                onAutoHeight={onAutoHeight}
                onEditEnd={onEditEnd}
              />
            ))}
          </SurfaceContext.Provider>
        </div>

        <div className="overlay">
          {selEls.map((el) => (
            <div
              key={el.id}
              className={`sel-box${editingId === el.id ? ' editing' : ''}`}
              style={{
                left: el.x * scale,
                top: el.y * scale,
                width: el.w * scale,
                height: el.h * scale,
                transform: `rotate(${el.rotation}deg)`,
              }}
            >
              {single && editingId !== el.id && (
                <>
                  {(isLineShape(el) ? LINE_HANDLES : ALL_HANDLES).map((h) => (
                    <div key={h} className={`handle h-${h}`} onPointerDown={(e) => onResizeDown(e, el, h)} />
                  ))}
                  <div className="rot-stem" />
                  <div className="rot-handle" title={t('Rotate (Shift snaps 15°)')} onPointerDown={(e) => onRotateDown(e, el)} />
                </>
              )}
            </div>
          ))}
          {showAnimBadges &&
            slide.elements.map((el) => {
              const nums = [...new Set((slide.animations ?? []).filter((a) => a.elementId === el.id).map((a) => animNums.get(a.id)!))];
              if (!nums.length) return null;
              return (
                <div key={'n' + el.id} className="anim-badges" style={{ left: el.x * scale, top: el.y * scale }}>
                  {nums.map((n) => (
                    <span key={n}>{digits(n)}</span>
                  ))}
                </div>
              );
            })}
          {guides.v.map((x) => (
            <div key={'v' + x} className="guide guide-v" style={{ left: x * scale }} />
          ))}
          {guides.h.map((y) => (
            <div key={'h' + y} className="guide guide-h" style={{ top: y * scale }} />
          ))}
          {marquee && (
            <div
              className="marquee"
              style={{ left: marquee.x * scale, top: marquee.y * scale, width: marquee.w * scale, height: marquee.h * scale }}
            />
          )}
        </div>
      </div>
      {menu && <ContextMenu x={menu.x} y={menu.y} items={menuItems()} onClose={closeMenu} />}
    </div>
  );
}
