/**
 * PowerPoint animations (the slide's <p:timing> tree): written after pptxgenjs builds the file, and
 * read back on import. Effects carry real behaviours (visibility, fades, motion, scale, rotation) so
 * PowerPoint plays them, plus preset ids so they show up by name in PowerPoint's Animation Pane.
 */
import { buildSteps, type AnimDirection, type AnimKind, type AnimTrigger, type Animation } from './anim';
import { uid } from './factory';
import type { SlideElement } from './types';

const CLASS: Record<AnimKind, string> = { entrance: 'entr', emphasis: 'emph', exit: 'exit' };
const KIND_BY_CLASS: Record<string, AnimKind> = { entr: 'entrance', emph: 'emphasis', exit: 'exit' };
const PRESET: Record<string, number> = { appear: 1, fly: 2, fade: 10, wipe: 22, float: 42, zoom: 53, spin: 8, pulse: 26, teeter: 32 };
const EFFECT_BY_PRESET: Record<number, string> = Object.fromEntries(Object.entries(PRESET).map(([k, v]) => [v, k]));
const SUBTYPE: Record<AnimDirection, number> = { top: 1, right: 2, bottom: 4, left: 8 };
const DIR_BY_SUBTYPE: Record<number, AnimDirection> = { 1: 'top', 2: 'right', 4: 'bottom', 8: 'left' };
const WIPE: Record<AnimDirection, string> = { bottom: 'down', top: 'up', left: 'left', right: 'right' };

// ---------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------

/** <p:timing> for a slide, or '' when nothing animates. `spids` maps element ids to shape ids. */
export function timingXml(animations: Animation[], elements: SlideElement[], spids: Map<string, number>): string {
  const usable = animations.filter((a) => spids.has(a.elementId));
  const steps = buildSteps(usable, elements);
  if (!steps.length) return '';
  let next = 1;
  const id = () => next++;
  const rootId = id();
  const seqId = id();
  const built = new Set<number>();

  const clicks = steps.map((step) => {
    // Items that start together form one time group; groups start one after another.
    const groups: { start: number; items: typeof step.items }[] = [];
    for (const it of step.items) {
      const base = it.start - it.anim.delay;
      const g = groups[groups.length - 1];
      if (g && g.start === base) g.items.push(it);
      else groups.push({ start: base, items: [it] });
    }
    const clickId = id();
    const groupXml = groups
      .map((g, gi) => {
        const gid = id();
        const effects = g.items
          .map((it, ii) => {
            const spid = spids.get(it.target.elementId)!;
            if (it.target.para != null) built.add(spid);
            const nodeType =
              gi === 0 && ii === 0 ? (step.auto ? (it.anim.trigger === 'after' ? 'afterEffect' : 'withEffect') : 'clickEffect') : ii === 0 ? 'afterEffect' : 'withEffect';
            return effectXml(it.anim, spid, it.target.para, nodeType, id);
          })
          .join('');
        return `<p:par><p:cTn id="${gid}" fill="hold"><p:stCondLst><p:cond delay="${Math.round(g.start)}"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par>`;
      })
      .join('');
    const start = step.auto
      ? `<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="${seqId}"/></p:cond>`
      : '<p:cond delay="indefinite"/>';
    return `<p:par><p:cTn id="${clickId}" fill="hold"><p:stCondLst>${start}</p:stCondLst><p:childTnLst>${groupXml}</p:childTnLst></p:cTn></p:par>`;
  });

  const bld = [...built].map((spid) => `<p:bldP spid="${spid}" grpId="0" build="p"/>`).join('');
  return (
    '<p:timing><p:tnLst><p:par>' +
    `<p:cTn id="${rootId}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${seqId}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clicks.join('')}</p:childTnLst></p:cTn>` +
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>' +
    '</p:childTnLst></p:cTn></p:par></p:tnLst>' +
    (bld ? `<p:bldLst>${bld}</p:bldLst>` : '') +
    '</p:timing>'
  );
}

function effectXml(a: Animation, spid: number, para: number | undefined, nodeType: string, id: () => number): string {
  const tgt =
    para == null ? `<p:spTgt spid="${spid}"/>` : `<p:spTgt spid="${spid}"><p:txEl><p:pRg st="${para}" end="${para}"/></p:txEl></p:spTgt>`;
  const d = Math.max(1, Math.round(a.duration));
  const dir = a.direction ?? 'bottom';
  const cBhvr = (inner: string, extra = '') =>
    `<p:cBhvr${extra}><p:cTn id="${id()}" dur="${inner}" fill="hold"/><p:tgtEl>${tgt}</p:tgtEl>`;
  const set = (val: 'visible' | 'hidden', at: number) =>
    `<p:set><p:cBhvr><p:cTn id="${id()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${at}"/></p:stCondLst></p:cTn><p:tgtEl>${tgt}</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${val}"/></p:to></p:set>`;
  const filter = (transition: 'in' | 'out', f: string) =>
    `<p:animEffect transition="${transition}" filter="${f}">${cBhvr(String(d))}</p:cBhvr></p:animEffect>`;
  const prop = (attr: string, from: string, to: string) =>
    `<p:anim calcmode="lin" valueType="num">${cBhvr(String(d), ' additive="base"')}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
  const rot = (by: number, dur: number, at: number) =>
    `<p:animRot by="${Math.round(by * 60000)}"><p:cBhvr><p:cTn id="${id()}" dur="${Math.max(1, Math.round(dur))}" fill="hold"><p:stCondLst><p:cond delay="${Math.round(at)}"/></p:stCondLst></p:cTn><p:tgtEl>${tgt}</p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;

  const off = {
    bottom: ['#ppt_x', '1+#ppt_h/2'],
    top: ['#ppt_x', '0-#ppt_h/2'],
    left: ['0-#ppt_w/2', '#ppt_y'],
    right: ['1+#ppt_w/2', '#ppt_y'],
  }[dir];
  let body: string;
  const entering = a.kind === 'entrance';
  switch (a.kind === 'emphasis' ? a.effect : `${a.kind}:${a.effect}`) {
    case 'pulse':
      body = `<p:animScale><p:cBhvr><p:cTn id="${id()}" dur="${Math.round(d / 2)}" autoRev="1" fill="hold"/><p:tgtEl>${tgt}</p:tgtEl></p:cBhvr><p:by x="108000" y="108000"/></p:animScale>`;
      break;
    case 'spin':
      body = rot(360, d, 0);
      break;
    case 'teeter':
      body = [5, -10, 10, -10, 5].map((by, i) => rot(by, d / 5, (i * d) / 5)).join('');
      break;
    case 'entrance:appear':
      body = set('visible', 0);
      break;
    case 'exit:appear':
      body = set('hidden', 0);
      break;
    case 'entrance:fly':
    case 'exit:fly':
      body = entering
        ? set('visible', 0) + prop('ppt_x', off[0], '#ppt_x') + prop('ppt_y', off[1], '#ppt_y')
        : prop('ppt_x', '#ppt_x', off[0]) + prop('ppt_y', '#ppt_y', off[1]) + set('hidden', d - 1);
      break;
    case 'entrance:wipe':
    case 'exit:wipe':
      body = entering ? set('visible', 0) + filter('in', `wipe(${WIPE[dir]})`) : filter('out', `wipe(${WIPE[dir]})`) + set('hidden', d - 1);
      break;
    case 'entrance:zoom':
    case 'exit:zoom':
      body = entering
        ? set('visible', 0) + prop('ppt_w', '0', '#ppt_w') + prop('ppt_h', '0', '#ppt_h') + filter('in', 'fade')
        : prop('ppt_w', '#ppt_w', '0') + prop('ppt_h', '#ppt_h', '0') + filter('out', 'fade') + set('hidden', d - 1);
      break;
    case 'entrance:float':
    case 'exit:float':
      body = entering
        ? set('visible', 0) + filter('in', 'fade') + prop('ppt_x', '#ppt_x', '#ppt_x') + prop('ppt_y', '#ppt_y+.1', '#ppt_y')
        : filter('out', 'fade') + prop('ppt_x', '#ppt_x', '#ppt_x') + prop('ppt_y', '#ppt_y', '#ppt_y+.1') + set('hidden', d - 1);
      break;
    default: // fade
      body = entering ? set('visible', 0) + filter('in', 'fade') : filter('out', 'fade') + set('hidden', d - 1);
  }
  const sub = PRESET[a.effect] === 2 || PRESET[a.effect] === 22 ? SUBTYPE[dir] : a.effect === 'zoom' ? 16 : 0;
  return (
    `<p:par><p:cTn id="${id()}" presetID="${PRESET[a.effect] ?? 10}" presetClass="${CLASS[a.kind]}" presetSubtype="${sub}" fill="hold" grpId="0" nodeType="${nodeType}">` +
    `<p:stCondLst><p:cond delay="${Math.round(a.delay)}"/></p:stCondLst><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>`
  );
}

// ---------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------

const local = (el: Element | null | undefined, name: string): Element[] =>
  el ? Array.from(el.children).filter((c) => c.localName === name) : [];
const first = (el: Element | null | undefined, name: string) => local(el, name)[0] ?? null;
const cTnOf = (par: Element) => first(par, 'cTn');
const children = (cTn: Element | null) => local(first(cTn, 'childTnLst'), 'par');
const delayOf = (cTn: Element | null) => {
  const v = first(first(cTn, 'stCondLst'), 'cond')?.getAttribute('delay');
  return v && v !== 'indefinite' ? Number(v) || 0 : 0;
};

/** Read a slide's main animation sequence. Elements are matched through PowerPoint shape ids. */
export function readTiming(slide: Document, spids: Map<string, string[]>, elements: SlideElement[], warnings: Set<string>): Animation[] {
  const timing = slide.getElementsByTagNameNS('*', 'timing')[0];
  if (!timing) return [];
  const main = Array.from(timing.getElementsByTagNameNS('*', 'cTn')).find((c) => c.getAttribute('nodeType') === 'mainSeq');
  if (!main) return [];
  const known = new Set(elements.map((e) => e.id));
  const out: Animation[] = [];
  let simplified = false;

  for (const click of children(main)) {
    for (const group of children(cTnOf(click))) {
      for (const eff of children(cTnOf(group))) {
        const cTn = cTnOf(eff);
        const cls = cTn?.getAttribute('presetClass') ?? '';
        const kind = KIND_BY_CLASS[cls];
        if (!cTn || !kind) {
          if (cls && cls !== 'mediacall') simplified = true;
          continue;
        }
        const tgt = cTn.getElementsByTagNameNS('*', 'spTgt')[0];
        const ids = spids.get(tgt?.getAttribute('spid') ?? '')?.filter((x) => known.has(x));
        if (!ids?.length) continue;
        const presetId = Number(cTn.getAttribute('presetID'));
        let effect = EFFECT_BY_PRESET[presetId];
        if (!effect || (kind === 'emphasis') !== ['pulse', 'spin', 'teeter'].includes(effect)) {
          effect = kind === 'emphasis' ? 'pulse' : 'fade';
          simplified = true;
        }
        const node = cTn.getAttribute('nodeType');
        const trigger: AnimTrigger = node === 'clickEffect' ? 'click' : node === 'afterEffect' ? 'after' : 'with';
        const durations = Array.from(cTn.getElementsByTagNameNS('*', 'cTn'))
          // An auto-reversing behaviour (e.g. Pulse) runs twice its own duration.
          .map((c) => Number(c.getAttribute('dur')) * (c.getAttribute('autoRev') === '1' ? 2 : 1))
          .filter((v) => v > 1);
        const duration = durations.length ? Math.max(...durations) : effect === 'appear' ? 1 : 500;
        const sub = Number(cTn.getAttribute('presetSubtype'));
        const pRg = tgt?.getElementsByTagNameNS('*', 'pRg')[0];
        const para = pRg ? Number(pRg.getAttribute('st')) : undefined;

        // Paragraph builds arrive as one effect per paragraph: fold them into one animation.
        const prev = out[out.length - 1];
        if (para != null && prev?.byParagraph && prev.elementId === ids[0] && prev.kind === kind && prev.effect === effect) continue;

        ids.forEach((elementId, k) => {
          out.push({
            id: uid(),
            elementId,
            kind,
            effect,
            trigger: k === 0 ? trigger : 'with',
            duration,
            delay: delayOf(cTn),
            direction: effect === 'fly' || effect === 'wipe' ? (DIR_BY_SUBTYPE[sub] ?? 'bottom') : undefined,
            byParagraph: para != null ? true : undefined,
          });
        });
      }
    }
  }
  if (simplified) warnings.add('some animation effects were simplified');
  return out;
}
