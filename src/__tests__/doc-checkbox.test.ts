// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection, type Command, type EditorState } from 'prosemirror-state';
import { undo } from 'prosemirror-history';
import { EditorView } from 'prosemirror-view';
import { DOMParser as PMDOMParser, DOMSerializer, type Node as PMNode } from 'prosemirror-model';
import { schema, type ParaAttrs } from '../doc/schema';
import { blankMeta } from '../doc/model';
import { createState } from '../doc/editor';
import { backspace, enter, insertCheckBox, shiftTab, tab, toggleCheck, toggleCheckAt, toggleList } from '../doc/commands';
import { buildDocx, listNumbers, readDocx } from '../doc/docx';
import { fromHtml, toHtmlBody, toText } from '../doc/convert';
import { AR } from '../i18n.ar';

const { nodes, marks } = schema;
const p = (text: string, attrs: Partial<ParaAttrs> = {}) => nodes.paragraph.create(attrs, text ? schema.text(text) : undefined);
const docOf = (...blocks: PMNode[]) => nodes.doc.create(null, blocks);
const box = (checked = false) => nodes.check_box.create({ checked });

function at(doc: PMNode, from: number, to = from) {
  const state = createState(doc);
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
}
function run(state: EditorState, cmd: Command) {
  let s = state;
  expect(cmd(s, (tr) => (s = s.apply(tr)))).toBe(true);
  return s;
}
const attrsOf = (s: EditorState | PMNode, i: number) => ('doc' in s ? s.doc : s).child(i).attrs as ParaAttrs;

/** A .docx holding only word/document.xml with this body (as Word writes it). */
async function wordFile(body: string) {
  const zip = new JSZip();
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>${body}</w:body></w:document>`,
  );
  return (await readDocx(await zip.generateAsync({ type: 'uint8array' }), 'W')).doc;
}
const sdt = (checked: boolean) =>
  `<w:sdt><w:sdtPr><w:id w:val="-1"/><w14:checkbox><w14:checked w14:val="${checked ? 1 : 0}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>` +
  `<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr><w:t>${checked ? '☒' : '☐'}</w:t></w:r></w:sdtContent></w:sdt>`;
const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

describe('Checklists: commands', () => {
  it('turns paragraphs into a checklist and back, and keeps the tick through the toggle', () => {
    const doc = docOf(p('milk'), p('eggs'));
    const s = run(at(doc, 1, 8), toggleList('check'));
    expect([0, 1].map((i) => attrsOf(s, i).list)).toEqual(['check', 'check']);
    expect(listNumbers(s.doc)).toEqual(['☐', '☐']);
    const off = run(s, toggleList('check'));
    expect(attrsOf(off, 0)).toMatchObject({ list: null, checked: null });
    // From bullets: the items become checklist items, unticked.
    const b = run(at(docOf(p('a', { list: 'bullet', level: 1, listFmt: 'arrow' })), 1), toggleList('check'));
    expect(attrsOf(b, 0)).toMatchObject({ list: 'check', level: 1, listFmt: null, checked: null });
  });

  it('ticks an item without moving the selection, as one undo step', () => {
    const doc = docOf(p('milk', { list: 'check' }), p('eggs', { list: 'check' }));
    const s0 = at(doc, 8, 10);
    const s = run(s0, toggleCheckAt(8));
    expect(attrsOf(s, 1).checked).toBe(true);
    expect(attrsOf(s, 0).checked).toBeNull();
    expect([s.selection.from, s.selection.to]).toEqual([8, 10]);
    const again = run(s, toggleCheckAt(9));
    expect(attrsOf(again, 1).checked).toBeNull();
    const back = run(again, undo);
    expect(attrsOf(back, 1).checked).toBe(true);
    expect(run(back, undo).doc.eq(doc)).toBe(true);
    // Not a checklist item: nothing to tick.
    expect(toggleCheckAt(1)(at(docOf(p('x')), 1))).toBe(false);
  });

  it('Ctrl+Alt+Enter ticks the selected items (all of them when any is clear)', () => {
    const doc = docOf(p('a', { list: 'check', checked: true }), p('b', { list: 'check' }), p('plain'));
    const s = run(at(doc, 1, 4), toggleCheck);
    expect([0, 1].map((i) => attrsOf(s, i).checked)).toEqual([true, true]);
    expect(attrsOf(run(s, toggleCheck), 0).checked).toBeNull();
    expect(toggleCheck(at(doc, 8))).toBe(false);
  });

  it('Enter makes a new unticked item; Enter on an empty item ends the list', () => {
    const doc = docOf(p('milk', { list: 'check', checked: true }));
    const s = run(at(doc, 5), enter);
    expect(s.doc.childCount).toBe(2);
    expect(attrsOf(s, 0)).toMatchObject({ list: 'check', checked: true });
    expect(attrsOf(s, 1)).toMatchObject({ list: 'check', checked: null });
    const ended = run(s, enter);
    expect(ended.doc.childCount).toBe(2);
    expect(attrsOf(ended, 1)).toMatchObject({ list: null, checked: null });
    // At the start of a ticked item: the new empty item above is unticked, the text keeps its tick.
    const top = run(at(doc, 1), enter);
    expect([attrsOf(top, 0).checked, attrsOf(top, 1).checked]).toEqual([null, true]);
    expect(top.doc.child(1).textContent).toBe('milk');
    expect(run(top, undo).doc.eq(doc)).toBe(true);
  });

  it('Tab and Shift+Tab change the level; Backspace at the start removes the box', () => {
    const doc = docOf(p('a', { list: 'check' }), p('b', { list: 'check', checked: true }));
    const s = run(at(doc, 4), tab);
    expect(attrsOf(s, 1)).toMatchObject({ list: 'check', level: 1, checked: true, listFmt: null });
    expect(attrsOf(run(s, shiftTab), 1).level).toBe(0);
    const bs = run(at(doc, 4), backspace);
    expect(attrsOf(bs, 1)).toMatchObject({ list: null, checked: null });
    expect(bs.doc.child(1).textContent).toBe('b');
  });
});

describe('Inline check boxes', () => {
  it('inserts a check box in text and in table cells; clicking it toggles (undoable)', () => {
    const cell = (t: string) => nodes.table_cell.create(null, p(t));
    const doc = docOf(p('I agree'), nodes.table.create(null, [nodes.table_row.create(null, [cell('yes'), cell('no')])]));
    let s = run(at(doc, 3), insertCheckBox);
    expect(s.doc.child(0).child(1).type).toBe(nodes.check_box);
    expect(s.selection.from).toBe(4);
    const inCell = s.doc.child(0).nodeSize + 4; // inside the first cell's paragraph
    s = run(s.apply(s.tr.setSelection(TextSelection.create(s.doc, inCell))), insertCheckBox);
    let boxes: { pos: number; checked: boolean }[] = [];
    const find = (d: PMNode) => {
      boxes = [];
      d.descendants((n, pos) => void (n.type === nodes.check_box && boxes.push({ pos, checked: n.attrs.checked })));
      return boxes;
    };
    expect(find(s.doc)).toHaveLength(2);
    const sel = s.selection.from;
    const t = run(s, toggleCheckAt(boxes[1].pos));
    expect(find(t.doc).map((b) => b.checked)).toEqual([false, true]);
    expect(t.selection.from).toBe(sel);
    expect(find(run(t, undo).doc).map((b) => b.checked)).toEqual([false, false]);
    // Ctrl+Alt+Enter next to a box toggles it too.
    expect(find(run(at(t.doc, 4), toggleCheck).doc)[0].checked).toBe(true);
  });

  it('goes with copied text (HTML and plain text) and with saved files', async () => {
    const doc = docOf(
      nodes.paragraph.create(null, [schema.text('a '), box(true), schema.text(' b', [marks.bold.create({ value: true })])]),
      p('item', { list: 'check', checked: true, level: 2 }),
    );
    const div = document.createElement('div');
    div.appendChild(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
    expect(PMDOMParser.fromSchema(schema).parse(div).toJSON()).toEqual(doc.toJSON());
    expect(doc.textBetween(0, doc.content.size, '\n')).toBe('a ☒ b\nitem');
    expect(toText(doc)).toBe('a ☒ b\r\n        ☑ item\r\n');
    const { fromFile, toFile } = await import('../doc/store');
    expect(fromFile(JSON.parse(JSON.stringify(toFile(blankMeta('n'), doc))), 'n').doc.eq(doc)).toBe(true);
  });
});

describe('Next Writer files and HTML', () => {
  it('opens documents saved before checklists existed', async () => {
    const { fromFile } = await import('../doc/store');
    const old = {
      format: 'nxdoc',
      version: 1,
      ...blankMeta('Old'),
      content: { type: 'doc', content: [{ type: 'paragraph', attrs: { style: 'normal', list: 'bullet', level: 0 }, content: [{ type: 'text', text: 'x' }] }] },
    };
    expect(fromFile(old, 'x').doc.child(0).attrs).toMatchObject({ list: 'bullet', checked: null });
  });

  it('exports checklists as lists of check boxes and reads them (and task lists) back', () => {
    const doc = docOf(p('done', { list: 'check', checked: true }), p('todo', { list: 'check' }));
    const html = toHtmlBody(doc);
    expect(html).toMatch(/<ul style="list-style-type: none;"><li[^>]*><input type="checkbox" disabled="" checked="">/);
    const back = fromHtml(`<html><body>${html}</body></html>`, 'x').doc;
    expect([0, 1].map((i) => [attrsOf(back, i).list, attrsOf(back, i).checked, back.child(i).textContent.trim()])).toEqual([
      ['check', true, 'done'],
      ['check', null, 'todo'],
    ]);
    const gh = fromHtml('<ul><li class="task-list-item"><input type="checkbox" checked disabled> ship it</li></ul>', 'x').doc;
    expect(attrsOf(gh, 0)).toMatchObject({ list: 'check', checked: true });
  });
});

describe('Editor: drawing and clicking', () => {
  function view(doc: PMNode, from: number) {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const v = new EditorView(host, { state: at(doc, from), dispatchTransaction: (tr) => v.updateState(v.state.apply(tr)) });
    return { v, done: () => (v.destroy(), host.remove()) };
  }
  const down = (el: Element) => el.dispatchEvent(new MouseEvent('mousedown', { button: 0, bubbles: true, cancelable: true }));

  it('draws a box where the bullet goes and ticks it on click, leaving the caret alone', () => {
    const { v, done } = view(docOf(p('first', { list: 'check' }), p('second', { list: 'check', checked: true }), p('plain')), 3);
    const ps = [...v.dom.querySelectorAll('.dw-p')] as HTMLElement[];
    expect(ps.map((e) => e.getAttribute('data-list'))).toEqual(['check', 'check', null]);
    expect(ps.map((e) => e.getAttribute('data-marker'))).toEqual([null, null, null]);
    const boxes = () => [...v.dom.querySelectorAll('.dw-check')];
    expect(boxes().map((b) => b.classList.contains('on'))).toEqual([false, true]);
    expect(down(boxes()[0])).toBe(false); // default prevented
    expect(attrsOf(v.state, 0).checked).toBe(true);
    expect(v.state.selection.from).toBe(3);
    expect(boxes().map((b) => b.classList.contains('on'))).toEqual([true, true]);
    down(boxes()[1]);
    expect(attrsOf(v.state, 1).checked).toBeNull();
    done();
  });

  it('toggles an inline check box on click and runs the checklist shortcuts', () => {
    const { v, done } = view(docOf(nodes.paragraph.create(null, [schema.text('ok '), box()]), p('next')), 2);
    const cb = v.dom.querySelector('.dw-cb')!;
    expect(cb.getAttribute('contenteditable')).toBe('false');
    down(cb);
    expect(v.state.doc.child(0).child(1).attrs.checked).toBe(true);
    expect(v.dom.querySelector('.dw-cb')!.classList.contains('on')).toBe(true);
    expect(v.state.selection.from).toBe(2);
    const key = (init: KeyboardEventInit) => v.dom.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }));
    v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, 7)));
    key({ key: '(', keyCode: 57, ctrlKey: true, shiftKey: true });
    expect(attrsOf(v.state, 1).list).toBe('check');
    key({ key: 'Enter', keyCode: 13, ctrlKey: true, altKey: true });
    expect(attrsOf(v.state, 1).checked).toBe(true);
    done();
  });
});

describe('Word files', () => {
  it('writes check boxes as Word Check Box content controls, and checklist items without numbering', async () => {
    const doc = docOf(
      p('Ticked', { list: 'check', checked: true }),
      p('Sub', { list: 'check', level: 1 }),
      p('Bullet', { list: 'bullet' }),
      nodes.paragraph.create(null, [box(true), schema.text(' starts with a box')]),
    );
    const zip = await JSZip.loadAsync((await buildDocx(doc, blankMeta('C'))).data);
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
    expect(xml).toContain('mc:Ignorable="w14"');
    const paras = xml.match(/<w:p>.*?<\/w:p>/g)!;
    expect(paras[0]).toMatch(
      /^<w:p><w:pPr><w:pStyle w:val="ListParagraph"\/><w:ind w:left="720" w:hanging="360"\/><\/w:pPr><w:sdt><w:sdtPr><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"\/><\/w:rPr><w:id w:val="\d+"\/>/,
    );
    expect(paras[0]).toContain(
      '<w14:checkbox><w14:checked w14:val="1"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr><w:t>☒</w:t></w:r></w:sdtContent></w:sdt><w:r><w:tab/></w:r>',
    );
    expect(paras[0]).not.toContain('numPr');
    expect(paras[1]).toContain('<w:ind w:left="1440" w:hanging="360"/>');
    expect(paras[1]).toContain('<w14:checked w14:val="0"/>');
    expect(paras[1]).toContain('<w:t>☐</w:t>');
    expect(paras[2]).toContain('<w:numPr>');
    // An inline box at the start of a paragraph is marked as such.
    expect(paras[3]).toContain('<w:tag w:val="nx-inline-check"/>');
    const numbering = await zip.file('word/numbering.xml')!.async('string');
    expect(numbering.match(/<w:num /g)).toHaveLength(1);
  });

  it('keeps checklists and check boxes through export and import', async () => {
    const cell = nodes.table_cell.create(null, nodes.paragraph.create(null, [box(true), schema.text(' Yes')]));
    const doc = docOf(
      p('Done', { list: 'check', checked: true }),
      p('Nested', { list: 'check', level: 1 }),
      nodes.paragraph.create({ list: 'check' }, [schema.text('\tTabbed '), box(true)]),
      p('عربي', { list: 'check', dir: 'rtl', checked: true }),
      nodes.paragraph.create(null, [box(), schema.text('inline first')]),
      nodes.paragraph.create(null, [schema.text('size '), nodes.check_box.create({ checked: true }, null, [marks.fontSize.create({ value: 16 })])]),
      nodes.table.create(null, [nodes.table_row.create(null, [cell])]),
    );
    const back = (await readDocx((await buildDocx(doc, blankMeta('R'))).data, 'R')).doc;
    expect(back.childCount).toBe(doc.childCount);
    expect([0, 1, 2, 3].map((i) => [attrsOf(back, i).list, attrsOf(back, i).level, attrsOf(back, i).checked, attrsOf(back, i).indentLeft, attrsOf(back, i).firstLine])).toEqual([
      ['check', 0, true, null, null],
      ['check', 1, null, null, null],
      ['check', 0, null, null, null],
      ['check', 0, true, null, null],
    ]);
    expect(back.child(0).textContent).toBe('Done');
    expect(back.child(2).toJSON()).toEqual(doc.child(2).toJSON());
    expect(attrsOf(back, 3).dir).toBe('rtl');
    expect(back.child(4).toJSON()).toEqual(doc.child(4).toJSON());
    expect(back.child(5).toJSON()).toEqual(doc.child(5).toJSON());
    expect(back.child(6).firstChild!.firstChild!.firstChild!.toJSON()).toEqual(doc.child(6).firstChild!.firstChild!.firstChild!.toJSON());
  });

  it("reads Word's check boxes: content controls, legacy form fields and box characters", async () => {
    const doc = await wordFile(
      // A box at the start of a plain paragraph, then a tab: a checklist item at the margin.
      `<w:p>${sdt(true)}<w:r><w:tab/></w:r>${r('Call the bank')}</w:p>` +
        // Indented as Word's hanging list layout, with a space after the box: level 1.
        `<w:p><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:bookmarkStart w:id="0" w:name="x"/>${sdt(false)}${r(' Pay rent')}</w:p>` +
        // A box in the middle of the text.
        `<w:p>${r('Agree ')}${sdt(true)}${r(' or not')}</w:p>` +
        // A legacy form check box.
        `<w:p>${r('Legacy ')}<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check1"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked/></w:checkBox></w:ffData></w:fldChar></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>${r(' end')}</w:p>` +
        `<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox><w:default w:val="1"/></w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>${r('by default')}</w:p>` +
        // Box characters typed in a symbol font.
        `<w:p>${r('☐', '<w:rFonts w:ascii="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol"/>')}${r(' Typed box')}</w:p>` +
        `<w:p><w:r><w:sym w:font="Wingdings" w:char="F0FE"/></w:r><w:r><w:tab/><w:t>Wingdings tick</w:t></w:r></w:p>` +
        `<w:p>${r('ý', '<w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/>')}${r(' Wingdings cross')}</w:p>` +
        // A table cell.
        `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p>${r('Yes ')}${sdt(false)}</w:p></w:tc></w:tr></w:tbl>`,
    );
    const a = (i: number) => attrsOf(doc, i);
    expect([a(0).list, a(0).checked, a(0).level, a(0).indentLeft, a(0).firstLine, doc.child(0).textContent]).toEqual(['check', true, 0, 0, 0, 'Call the bank']);
    expect([a(1).list, a(1).checked, a(1).level, a(1).indentLeft, doc.child(1).textContent]).toEqual(['check', null, 1, null, 'Pay rent']);
    expect(a(2).list).toBeNull();
    expect(doc.child(2).toJSON().content.map((n: { type: string; attrs?: { checked: boolean } }) => n.attrs?.checked ?? n.type)).toEqual(['text', true, 'text']);
    expect(doc.child(3).child(1).type).toBe(nodes.check_box);
    expect(doc.child(3).child(1).attrs.checked).toBe(true);
    expect(doc.child(3).textContent).toBe('Legacy ☒ end');
    expect(doc.child(4).child(0).attrs.checked).toBe(true);
    expect(a(4).list).toBeNull();
    expect([a(5).list, a(5).checked, doc.child(5).textContent]).toEqual(['check', null, 'Typed box']);
    expect([a(6).list, a(6).checked, doc.child(6).textContent]).toEqual(['check', true, 'Wingdings tick']);
    expect([a(7).list, a(7).checked, doc.child(7).textContent]).toEqual(['check', true, 'Wingdings cross']);
    const cellPara = doc.child(8).firstChild!.firstChild!.firstChild!;
    expect(cellPara.attrs.list).toBeNull();
    expect(cellPara.lastChild!.type).toBe(nodes.check_box);
  });
});

describe('Arabic interface', () => {
  it('translates the checklist and check box commands', () => {
    const names = ['Checklist (Ctrl+Shift+9)', 'Check Box', 'Controls', 'Check Box: click it to tick or clear it (Ctrl+Alt+Enter)'];
    expect(names.filter((n) => !(n in AR))).toEqual([]);
  });
});
