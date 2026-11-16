// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { schema, type ParaAttrs } from '../doc/schema';
import { blankMeta, normalizeMeta, type DocMeta } from '../doc/model';
import { createState, flattenLists } from '../doc/editor';
import {
  backspace,
  docStats,
  enter,
  insertPageBreak,
  insertTable,
  setDir,
  setLink,
  toggleFlag,
  toggleList,
  effectiveFlag,
} from '../doc/commands';
import { findKey, findMatches, setFind } from '../doc/plugins';
import { buildDocx, listNumbers, readDocx } from '../doc/docx';
import { fromHtml, fromText, toHtmlBody, toText } from '../doc/convert';
import { measure } from '../doc/pagination';

const { nodes, marks } = schema;
const p = (text: string, attrs: Partial<ParaAttrs> = {}, m: ReturnType<typeof marks.bold.create>[] = []) =>
  nodes.paragraph.create(attrs, text ? schema.text(text, m) : undefined);
const docOf = (...blocks: PMNode[]) => nodes.doc.create(null, blocks);

/** Run a command on a state with the caret at `pos` (or a selection from..to). */
function run(doc: PMNode, cmd: Command, from: number, to = from) {
  let state = createState(doc);
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
  cmd(state, (tr) => (state = state.apply(tr)));
  return state;
}

const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

describe('Next Writer editing', () => {
  it('numbers list paragraphs by level and restarts after other text', () => {
    const doc = docOf(
      p('one', { list: 'ordered' }),
      p('two', { list: 'ordered' }),
      p('two a', { list: 'ordered', level: 1 }),
      p('two b', { list: 'ordered', level: 1 }),
      p('three', { list: 'ordered' }),
      p('break'),
      p('bullet', { list: 'bullet' }),
      p('again', { list: 'ordered' }),
    );
    expect(listNumbers(doc)).toEqual(['1.', '2.', 'a.', 'b.', '3.', '•', '1.']);
  });

  it('Enter after a heading starts a Normal paragraph; Enter in an empty list item ends the list', () => {
    const doc = docOf(p('Heading', { style: 'h1' }));
    const s = run(doc, enter, 8);
    expect(s.doc.childCount).toBe(2);
    expect(s.doc.child(1).attrs.style).toBe('normal');

    const list = docOf(p('item', { list: 'bullet' }), p('', { list: 'bullet' }));
    const s2 = run(list, enter, 7);
    expect(s2.doc.child(1).attrs.list).toBeNull();
    // Enter in the middle of a list item keeps the list.
    const s3 = run(docOf(p('item', { list: 'ordered' })), enter, 3);
    expect([s3.doc.child(0).attrs.list, s3.doc.child(1).attrs.list]).toEqual(['ordered', 'ordered']);
  });

  it('Backspace at the start of a list item removes the bullet', () => {
    const s = run(docOf(p('item', { list: 'bullet' })), backspace, 1);
    expect(s.doc.child(0).attrs.list).toBeNull();
    expect(s.doc.child(0).textContent).toBe('item');
  });

  it('toggles lists on and off', () => {
    const doc = docOf(p('a'), p('b'));
    const on = run(doc, toggleList('ordered'), 1, 5);
    expect([on.doc.child(0).attrs.list, on.doc.child(1).attrs.list]).toEqual(['ordered', 'ordered']);
    let s = on;
    toggleList('ordered')(s, (tr) => (s = s.apply(tr)));
    expect(s.doc.child(0).attrs.list).toBeNull();
  });

  it('bold toggles what the reader sees, even in a bold-free heading style', () => {
    const doc = docOf(p('Heading text', { style: 'h1' }));
    const s = run(doc, toggleFlag('bold'), 1, 8);
    const first = s.doc.child(0).child(0);
    expect(marks.bold.isInSet(first.marks)?.attrs.value).toBe(true);
    let s2 = s.apply(s.tr.setSelection(TextSelection.create(s.doc, 2)));
    expect(effectiveFlag(s2, 'bold')).toBe(true);
    toggleFlag('bold')(s2.apply(s2.tr.setSelection(TextSelection.create(s2.doc, 1, 8))), (tr) => (s2 = s2.apply(tr)));
    expect(marks.bold.isInSet(s2.doc.child(0).child(0).marks)).toBeUndefined();
  });

  it('right-to-left mirrors an explicit left or right alignment', () => {
    const s = run(docOf(p('مرحبا', { align: 'left' })), setDir('rtl'), 1);
    expect(s.doc.child(0).attrs).toMatchObject({ dir: 'rtl', align: 'right' });
  });

  it('a page break splits the paragraph, and a table puts the caret in its first cell', () => {
    const s = run(docOf(p('HelloWorld')), insertPageBreak, 6);
    expect(s.doc.content.content.map((n) => n.type.name)).toEqual(['paragraph', 'page_break', 'paragraph']);
    expect([s.doc.child(0).textContent, s.doc.child(2).textContent]).toEqual(['Hello', 'World']);
    expect(s.selection.$from.parent.textContent).toBe('World');

    const t = run(docOf(p('Text')), insertTable(2, 3, 600), 5);
    expect(t.doc.child(1).type.name).toBe('table');
    expect(t.doc.child(1).childCount).toBe(2);
    expect(t.doc.child(1).child(0).childCount).toBe(3);
    expect(t.selection.$from.node(-1).type.name).toBe('table_cell');
    expect(t.doc.lastChild!.type.name).toBe('paragraph');
  });

  it('adds links over the selection or types the address', () => {
    const s = run(docOf(p('see site')), setLink('https://example.com'), 5, 9);
    expect(marks.link.isInSet(s.doc.child(0).child(1).marks)?.attrs.href).toBe('https://example.com');
    const s2 = run(docOf(p('')), setLink('https://a.org'), 1);
    expect(s2.doc.textContent).toBe('https://a.org');
  });

  it('finds matches inside paragraphs, ignoring case unless asked', () => {
    const doc = docOf(p('Cat cat CAT'), p('concatenate'));
    expect(findMatches(doc, 'cat', false)).toHaveLength(4);
    expect(findMatches(doc, 'cat', true)).toHaveLength(2);
    let state = createState(doc);
    state = state.apply(setFind(state.tr, { query: 'cat' }));
    expect(findKey.getState(state)!.matches).toHaveLength(4);
    state = state.apply(state.tr.insertText('cat ', 1));
    expect(findKey.getState(state)!.matches).toHaveLength(5);
  });

  it('counts words in English and Arabic', () => {
    const s = docStats(docOf(p('Hello, world — 2026!'), p('مرحبا بالعالم'), p('')));
    expect(s.words).toBe(5);
    expect(s.paragraphs).toBe(2);
  });
});

describe('Word files', () => {
  const meta: DocMeta = {
    ...blankMeta('Report'),
    font: 'Arial',
    size: 12,
    page: { width: 11.69, height: 8.27, margins: { top: 0.5, right: 0.75, bottom: 0.5, left: 0.75 } },
    header: { text: 'Quarterly report', align: 'center' },
    footer: { text: '', align: 'left' },
    pageNumbers: { pos: 'footer', align: 'right' },
  };

  function sample() {
    const cell = (text: string, attrs = {}) => nodes.table_cell.create(attrs, p(text));
    return docOf(
      p('The Title', { style: 'title' }),
      p('Intro', { style: 'h1' }),
      p('', {}),
      nodes.paragraph.create({ align: 'justify', lineHeight: 1.5, spaceAfter: 12 }, [
        schema.text('Plain '),
        schema.text('bold', [marks.bold.create({ value: true })]),
        schema.text(' red', [marks.color.create({ value: '#ff0000' }), marks.fontSize.create({ value: 14 })]),
        schema.text(' link', [marks.link.create({ href: 'https://example.com/' })]),
        schema.text('\tx'),
        schema.text('2', [marks.sup.create()]),
        nodes.hard_break.create(),
        nodes.image.create({ src: PNG, width: 40, height: 20, alt: 'dot' }),
      ]),
      p('مرحبا بالعالم', { align: 'left' }),
      p('first', { list: 'ordered' }),
      p('nested', { list: 'ordered', level: 1 }),
      p('second', { list: 'ordered' }),
      p('point', { list: 'bullet' }),
      nodes.page_break.create(),
      nodes.table.create(null, [
        nodes.table_row.create(null, [cell('A', { colspan: 2, colwidth: [100, 120] }), cell('C', { rowspan: 2, background: '#ffff00' })]),
        nodes.table_row.create(null, [cell('D'), cell('E')]),
      ]),
      nodes.horizontal_rule.create(),
      p('End'),
    );
  }

  it('round-trips text, styles, lists, tables, pictures, links, breaks and page setup', async () => {
    const { data, skippedImages } = await buildDocx(sample(), meta);
    expect(skippedImages).toBe(0);
    const { doc, meta: m, warnings } = await readDocx(data, 'Report');
    expect(warnings).toEqual([]);

    const names = doc.content.content.map((n) => n.type.name);
    expect(names.filter((n) => n === 'page_break')).toHaveLength(1);
    expect(names).toContain('table');
    expect(names).toContain('paragraph');

    const byText = (s: string) => {
      let found: PMNode | undefined;
      doc.descendants((n) => {
        if (!found && n.type === nodes.paragraph && n.textContent.startsWith(s)) found = n;
        return !found;
      });
      return found!;
    };
    expect(byText('The Title').attrs.style).toBe('title');
    expect(byText('Intro').attrs.style).toBe('h1');

    const body = byText('Plain');
    expect(body.attrs).toMatchObject({ align: 'justify', lineHeight: 1.5, spaceAfter: 12 });
    const runs = body.content.content;
    expect(runs.find((r) => r.text === 'bold')!.marks.map((mk) => mk.type.name)).toContain('bold');
    const red = runs.find((r) => r.text === ' red')!;
    expect(marks.color.isInSet(red.marks)?.attrs.value).toBe('#ff0000');
    expect(marks.fontSize.isInSet(red.marks)?.attrs.value).toBe(14);
    expect(marks.link.isInSet(runs.find((r) => r.text === ' link')!.marks)?.attrs.href).toBe('https://example.com/');
    expect(body.textContent).toContain('\tx2');
    expect(marks.sup.isInSet(runs.find((r) => r.text === '2')!.marks)).toBeTruthy();
    const img = runs.find((r) => r.type === nodes.image)!;
    expect(img.attrs).toMatchObject({ width: 40, height: 20, alt: 'dot' });
    expect(img.attrs.src).toMatch(/^data:image\/png;base64,/);
    expect(runs.some((r) => r.type === nodes.hard_break)).toBe(true);

    // Arabic: right-to-left, and an explicit "left" alignment survives Word's mirrored values.
    expect(byText('مرحبا').attrs).toMatchObject({ dir: 'rtl', align: 'left' });

    expect([byText('first'), byText('nested'), byText('second'), byText('point')].map((n) => [n.attrs.list, n.attrs.level])).toEqual([
      ['ordered', 0],
      ['ordered', 1],
      ['ordered', 0],
      ['bullet', 0],
    ]);
    expect(listNumbers(doc)).toEqual(['1.', 'a.', '2.', '•']);

    const table = doc.content.content.find((n) => n.type === nodes.table)!;
    const r0 = table.child(0);
    expect(r0.child(0).attrs).toMatchObject({ colspan: 2, colwidth: [100, 120] });
    expect(r0.child(1).attrs).toMatchObject({ rowspan: 2, background: '#ffff00' });
    expect(table.child(1).childCount).toBe(2);

    expect(m.font).toBe('Arial');
    expect(m.size).toBe(12);
    expect(m.page.width).toBeCloseTo(11.69, 2);
    expect(m.page.margins.left).toBeCloseTo(0.75, 2);
    expect(m.header).toEqual({ text: 'Quarterly report', align: 'center' });
    expect(m.pageNumbers).toEqual({ pos: 'footer', align: 'right' });
  });

  it('writes a package Word can open: content types, relationships and numbering', async () => {
    const { data } = await buildDocx(sample(), meta);
    const zip = await JSZip.loadAsync(data);
    const types = await zip.file('[Content_Types].xml')!.async('string');
    for (const part of ['/word/document.xml', '/word/styles.xml', '/word/numbering.xml', '/word/footer1.xml', '/word/header1.xml'])
      expect(types).toContain(part);
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels).toContain('TargetMode="External"');
    expect(rels).toContain('media/image1.png');
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagName('parsererror')).toHaveLength(0);
    expect(xml).toContain('w:orient="landscape"');
    expect(xml).toContain('<w:br w:type="page"/>');
    expect(await zip.file('word/footer1.xml')!.async('string')).toContain('PAGE');
  });

  it('reads Word documents with theme fonts, style-based lists, merged cells and fields', async () => {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
    zip.file(
      '_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    );
    zip.file(
      'word/theme/theme1.xml',
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
    );
    zip.file(
      'word/styles.xml',
      `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>` +
        `<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
        `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
        `<w:style w:type="paragraph" w:styleId="Titre1"><w:name w:val="heading 1"/></w:style>` +
        `<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style>` +
        `<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style></w:styles>`,
    );
    zip.file(
      'word/numbering.xml',
      `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num></w:numbering>`,
    );
    zip.file(
      'word/document.xml',
      `<w:document xmlns:w="${W}"><w:body>` +
        `<w:p><w:pPr><w:pStyle w:val="Titre1"/></w:pPr><w:r><w:t>Chapitre</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>Point</w:t></w:r></w:p>` +
        `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>5/10/2026</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` +
        `<w:ins><w:r><w:t xml:space="preserve"> added</w:t></w:r></w:ins><w:del><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>` +
        `<w:tbl><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>` +
        `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>tall</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>` +
        `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>d</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080"/></w:sectPr></w:body></w:document>`,
    );
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const { doc, meta: m, warnings } = await readDocx(bytes, 'Imported');
    expect(m.font).toBe('Aptos');
    expect(m.size).toBe(12);
    expect(m.page.width).toBeCloseTo(8.5);
    expect(m.page.margins.left).toBeCloseTo(0.75);
    expect(doc.child(0).attrs.style).toBe('h1');
    const point = doc.child(1);
    expect(point.attrs.list).toBe('bullet');
    expect(marks.bold.isInSet(point.child(0).marks)?.attrs.value).toBe(true);
    // The document's "no space after, single spacing" defaults are kept.
    expect(point.attrs).toMatchObject({ spaceAfter: 0, lineHeight: 1 });
    expect(doc.child(2).textContent).toBe('5/10/2026 added');
    expect(warnings).toContain('tracked changes (accepted)');
    const table = doc.child(3);
    expect(table.child(0).child(0).attrs.rowspan).toBe(2);
    expect(table.child(1).childCount).toBe(1);
  });
});

describe('HTML and text', () => {
  it('turns nested and Word-style lists into list paragraphs', () => {
    const div = document.createElement('div');
    div.innerHTML =
      '<ul><li>a<ul><li>a1</li></ul></li><li><p>b</p></li></ul><ol><li>one</li></ol>' +
      '<p class="MsoListParagraph" style="mso-list:l0 level2 lfo1"><span style="mso-list:Ignore">2.<span> </span></span>word</p>';
    flattenLists(div);
    const ps = [...div.querySelectorAll('p')].map((e) => [e.textContent, e.getAttribute('data-list'), e.getAttribute('data-level')]);
    expect(ps).toEqual([
      ['a', 'bullet', '0'],
      ['a1', 'bullet', '1'],
      ['b', 'bullet', '0'],
      ['one', 'ordered', '0'],
      ['word', 'ordered', '1'],
    ]);
  });

  it('opens HTML with headings, formatting and lists, and writes real lists back', () => {
    const { doc, meta } = fromHtml(
      '<html><head><title>Notes</title></head><body><h2>Plan</h2><p><b>Bold</b> and <i>italic</i></p><ol><li>x</li><li>y<ol><li>z</li></ol></li></ol></body></html>',
      'file',
    );
    expect(meta.title).toBe('Notes');
    expect(doc.child(0).attrs.style).toBe('h2');
    expect(marks.bold.isInSet(doc.child(1).child(0).marks)).toBeTruthy();
    expect(listNumbers(doc)).toEqual(['1.', '2.', 'a.']);
    const html = toHtmlBody(doc);
    expect(html).toMatch(/<ol><li[^>]*>x<\/li><li[^>]*>y<ol><li[^>]*>z<\/li><\/ol><\/li><\/ol>/);
  });

  it('reads and writes plain text', () => {
    const { doc } = fromText('﻿line 1\r\nline 2\n\nlast', 't');
    expect(doc.childCount).toBe(4);
    const out = toText(docOf(p('Intro'), p('a', { list: 'ordered' }), p('b', { list: 'ordered' })));
    expect(out).toBe('Intro\r\n1. a\r\n2. b\r\n');
  });
});

describe('Next Writer files and printing', () => {
  it('saves and restores .nxdoc documents', async () => {
    const { toFile, fromFile } = await import('../doc/store');
    const meta = { ...blankMeta('Letter'), pageNumbers: { pos: 'footer' as const, align: 'center' as const } };
    const doc = docOf(p('Dear reader,', {}, [marks.italic.create({ value: true })]));
    const file = JSON.parse(JSON.stringify(toFile(meta, doc)));
    const back = fromFile(file, 'x');
    expect(back.doc.eq(doc)).toBe(true);
    expect(back.meta).toEqual(meta);
    expect(() => fromFile({ format: 'nxsheet' }, 'x')).toThrow();
  });

  it('repairs page setup read from a damaged file', () => {
    const m = normalizeMeta({ page: { width: -1 } as never, font: '', size: 0, header: { text: 'H', align: 'middle' as never } }, 'T');
    expect(m.page.width).toBeCloseTo(8.27);
    expect(m.font).toBe('Calibri');
    expect(m.header).toEqual({ text: 'H', align: 'left' });
  });

  it('prints with the page size, margins, header and page numbers', async () => {
    const { pageCss } = await import('../doc/io');
    const css = pageCss({ ...blankMeta('x'), header: { text: 'Say "hi"', align: 'right' }, pageNumbers: { pos: 'footer', align: 'center' } });
    expect(css).toContain('size: 8.27in 11.69in');
    expect(css).toContain('margin: 1in 1in 1in 1in');
    expect(css).toContain('@top-right { content: "Say \\"hi\\""');
    expect(css).toContain('@bottom-center { content: counter(page)');
  });

  it("names copies and PDFs after the document's file", async () => {
    vi.useFakeTimers();
    const { documentName, suggestedPath, printDocument } = await import('../doc/io');
    const { useDoc } = await import('../doc/store');
    const print = vi.spyOn(window, 'print').mockImplementation(() => {});
    try {
      useDoc.setState({ filePath: 'C:\\Docs\\Hussein CV.docx', meta: { ...blankMeta('Résumé') } });
      expect(documentName()).toBe('Hussein CV');
      expect(suggestedPath('pdf')).toBe('C:\\Docs\\Hussein CV.pdf');
      expect(suggestedPath('txt')).toBe('C:\\Docs\\Hussein CV.txt');

      // Save as PDF goes through the print dialog, which names the file after the page title.
      document.title = 'Next Office';
      printDocument();
      vi.runAllTimers();
      expect(print).toHaveBeenCalled();
      expect(document.title).toBe('Hussein CV');
      window.dispatchEvent(new Event('afterprint'));
      expect(document.title).toBe('Next Office');

      // A document never saved: its title, made safe for a file name.
      useDoc.setState({ filePath: null, meta: { ...blankMeta('Q3: plan?') } });
      expect(suggestedPath('docx')).toBe('Q3 plan.docx');
    } finally {
      print.mockRestore();
      vi.useRealTimers();
    }
  });

  it("prints without the browser's own header and footer (date, title, URL, page count)", async () => {
    const { pageCss } = await import('../doc/io');
    // Chromium leaves out its header and footer when the page defines margin boxes; the corners are always empty.
    const css = pageCss(blankMeta('x'));
    for (const corner of ['top-left', 'top-right', 'bottom-left', 'bottom-right']) expect(css).toContain(`@${corner}-corner { content: ""; }`);
  });
});

describe('pagination', () => {
  /** jsdom has no layout: give each block a position and height. */
  function layout(view: EditorView, heights: number[]) {
    let y = 0;
    view.state.doc.forEach((_n, offset, i) => {
      const dom = view.nodeDOM(offset) as HTMLElement;
      const top = y;
      Object.defineProperty(dom, 'offsetTop', { configurable: true, get: () => top });
      Object.defineProperty(dom, 'offsetHeight', { configurable: true, get: () => heights[i] });
      y += heights[i];
    });
    Object.defineProperty(view.dom, 'offsetHeight', { configurable: true, get: () => y });
    view.dom.getBoundingClientRect = () => ({ top: 0, left: 0, right: 600, bottom: y, width: 600, height: y, x: 0, y: 0, toJSON() {} });
  }

  it('moves blocks that cross the bottom margin to the next page, and honours page breaks', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const doc = docOf(
      nodes.horizontal_rule.create(),
      nodes.horizontal_rule.create(),
      nodes.horizontal_rule.create(),
      nodes.page_break.create(),
      nodes.horizontal_rule.create(),
    );
    const view = new EditorView(host, { state: EditorState.create({ doc }) });
    layout(view, [400, 400, 400, 2, 100]);
    const r = measure(view, { contentH: 1000, stride: 1200 })!;
    // The third block (800–1200) doesn't fit on page 1 (0–1000): it starts page 2 (1200).
    expect(r.breaks[0]).toEqual({ pos: 2, height: 400, inline: false });
    // After the page break on page 2, the last block goes to page 3.
    expect(r.breaks[1].pos).toBe(doc.content.size - 1);
    expect(r.pages).toBe(3);
    view.destroy();
    host.remove();
  });
});
