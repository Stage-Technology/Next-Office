/** The paragraph styles as CSS (shared by the editor and HTML export), and HTML and plain-text conversion. */
import { DOMParser as PMDOMParser, DOMSerializer, type Node as PMNode } from 'prosemirror-model';
import { fontStack } from '../themes';
import { blankMeta, STYLES, type DocMeta, type ParaStyle } from './model';
import { checkChar, schema, type ParaAttrs } from './schema';
import { flattenLists } from './editor';
import { listMarker, numberLists } from './plugins';
import { checkMarker, levelDef, type NumFmt } from './lists';

/** CSS line-height for Word's "multiple" line spacing. */
export const cssLine = (multiple: number) => (multiple * 1.2).toFixed(3);

/** CSS for the document's text: Normal from the document's font and size, the other styles on top. */
export function styleCss(scope: string, meta: Pick<DocMeta, 'font' | 'size'>): string {
  const rules = [
    `${scope} .dw-p { margin: 0; font-family: ${fontStack(meta.font)}; font-size: ${meta.size}pt; font-weight: 400; font-style: normal; ` +
      `line-height: ${cssLine(1.08)}; padding: 0 0 8pt; color: inherit; white-space: pre-wrap; tab-size: 0.5in; }`,
  ];
  for (const [id, s] of Object.entries(STYLES) as [ParaStyle, (typeof STYLES)[ParaStyle]][]) {
    if (id === 'normal') continue;
    const css: string[] = [`padding-top: ${s.spaceBefore}pt`, `padding-bottom: ${s.spaceAfter}pt`];
    if (s.font) css.push(`font-family: ${fontStack(s.font)}`);
    if (s.size) css.push(`font-size: ${s.size}pt`);
    if (s.bold) css.push('font-weight: 700');
    if (s.italic) css.push('font-style: italic');
    if (s.color) css.push(`color: ${s.color}`);
    if (s.align) css.push(`text-align: ${s.align}`);
    if (s.lineHeight) css.push(`line-height: ${cssLine(s.lineHeight)}`);
    rules.push(`${scope} .dw-p.dw-${id} { ${css.join('; ')}; }`);
  }
  return rules.join('\n');
}

// ---- HTML ----

const CSS_LIST: Partial<Record<NumFmt, string>> = {
  decimal: 'decimal',
  lowerLetter: 'lower-alpha',
  upperLetter: 'upper-alpha',
  lowerRoman: 'lower-roman',
  upperRoman: 'upper-roman',
  hindiNumbers: 'arabic-indic',
};

/** The document as HTML body content, with real <ul>/<ol> lists so other apps keep them (checklist items start with a check box). */
export function toHtmlBody(doc: PMNode, d: Document = document): string {
  const ser = DOMSerializer.fromSchema(schema);
  const root = d.createElement('div');
  let stack: { el: HTMLElement; kind: string }[] = [];
  doc.forEach((node) => {
    const dom = ser.serializeNode(node, { document: d }) as HTMLElement;
    const a = node.attrs as ParaAttrs;
    if (node.type !== schema.nodes.paragraph || !a.list) {
      stack = [];
      root.appendChild(dom);
      return;
    }
    while (stack.length > a.level + 1 || (stack.length === a.level + 1 && stack[stack.length - 1].kind !== a.list)) stack.pop();
    while (stack.length < a.level + 1) {
      const list = d.createElement(a.list === 'ordered' ? 'ol' : 'ul');
      const lv = stack.length;
      if (lv === a.level && a.list === 'check') list.style.listStyleType = 'none';
      else if (lv === a.level && a.list !== 'check' && (a.listFmt || a.listStart != null)) {
        // The list's look (other apps have no "1.1." formats: those show as plain numbers).
        const def = levelDef(a.list, a.listFmt, lv);
        const css = def.fmt === 'bullet' ? JSON.stringify(`${def.text} `) : CSS_LIST[def.fmt];
        if (css) list.style.listStyleType = css;
        if (typeof a.listStart === 'number') list.setAttribute('start', String(a.listStart));
      }
      const top = stack[stack.length - 1]?.el;
      (top?.lastElementChild ?? top ?? root).appendChild(list);
      stack.push({ el: list, kind: a.list });
    }
    const li = d.createElement('li');
    for (const at of ['dir', 'style', 'class']) if (dom.getAttribute(at)) li.setAttribute(at, dom.getAttribute(at)!);
    if (a.list === 'check') {
      const box = d.createElement('input');
      box.type = 'checkbox';
      box.disabled = true;
      if (a.checked) box.setAttribute('checked', '');
      li.append(box, ' ');
    }
    li.append(...dom.childNodes);
    stack[stack.length - 1].el.appendChild(li);
  });
  return root.innerHTML;
}

const escHtml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export function toHtmlFile(doc: PMNode, meta: DocMeta): string {
  const w = meta.page.width - meta.page.margins.left - meta.page.margins.right;
  const rtl = (() => {
    let dir: string | null = null;
    doc.descendants((n) => {
      if (dir || !n.isTextblock) return !dir;
      dir = (n.attrs.dir as string | null) ?? (/[֐-ࣿ]/.test(n.textContent) ? 'rtl' : /\p{L}/u.test(n.textContent) ? 'ltr' : null);
      return false;
    });
    return dir === 'rtl';
  })();
  return `<!doctype html>
<html${rtl ? ' dir="rtl"' : ''}>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="Next Writer">
<title>${escHtml(meta.title)}</title>
<style>
body { max-width: ${w}in; margin: 0.75in auto; padding: 0 16px; color: #000; }
${styleCss('body', meta)}
body table { border-collapse: collapse; margin: 0 0 8pt; }
body td, body th { border: 1px solid #000; padding: 0 7px; vertical-align: top; }
body td .dw-p, body th .dw-p { padding-bottom: 0; }
body li.dw-p { display: list-item; padding-bottom: 0; }
body img { max-width: 100%; height: auto; }
body .dw-page-break { break-after: page; }
@media screen { body .dw-page-break { border-top: 1px dashed #bbb; margin: 12px 0; } }
</style>
</head>
<body>
${toHtmlBody(doc)}
</body>
</html>
`;
}

/** Read an HTML file (from a browser, Word's "Web Page" or another editor). */
export function fromHtml(html: string, title: string): { doc: PMNode; meta: DocMeta } {
  const d = new DOMParser().parseFromString(html, 'text/html');
  d.querySelectorAll('script, style, noscript, template, head').forEach((e) => e.remove());
  flattenLists(d.body);
  const doc = PMDOMParser.fromSchema(schema).parse(d.body);
  const t = new DOMParser().parseFromString(html, 'text/html').title.trim();
  return { doc, meta: blankMeta(t || title) };
}

// ---- Plain text ----

export function fromText(text: string, title: string): { doc: PMNode; meta: DocMeta } {
  const lines = text.replace(/^﻿/, '').split(/\r\n|\r|\n/);
  const paras = lines.map((l) => schema.nodes.paragraph.create(null, l ? schema.text(l) : undefined));
  return { doc: schema.nodes.doc.create(null, paras.length ? paras : [schema.nodes.paragraph.create()]), meta: blankMeta(title) };
}

/** Plain text: one line per paragraph, list numbers and bullets included, table cells separated by tabs. */
export function toText(doc: PMNode): string {
  const lines: string[] = [];
  const inline = (p: PMNode) => {
    let s = '';
    p.forEach((c) => (s += c.isText ? c.text : c.type === schema.nodes.hard_break ? '\n' : c.type === schema.nodes.check_box ? checkChar(c.attrs.checked) : ''));
    return s;
  };
  const markers = new Map<PMNode, string>();
  numberLists(doc, (n, _o, m) => markers.set(n, m));
  doc.forEach((node) => {
    if (node.type === schema.nodes.paragraph) {
      const a = node.attrs as ParaAttrs;
      const m = markers.get(node);
      lines.push((a.list ? '    '.repeat(a.level) + (m ?? (a.list === 'check' ? checkMarker(a.checked) : listMarker(a.list, a.level, 1))) + ' ' : '') + inline(node));
    } else if (node.type === schema.nodes.table) {
      node.forEach((row) => {
        const cells: string[] = [];
        row.forEach((cell) => {
          const ps: string[] = [];
          cell.forEach((p) => ps.push(inline(p)));
          cells.push(ps.join(' '));
        });
        lines.push(cells.join('\t'));
      });
    } else if (node.type === schema.nodes.horizontal_rule) lines.push('—'.repeat(20));
    else if (node.type === schema.nodes.page_break) lines.push('\f');
  });
  return lines.join('\r\n') + '\r\n';
}
