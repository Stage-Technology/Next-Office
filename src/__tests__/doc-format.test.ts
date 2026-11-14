// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection, type Command, type EditorState } from 'prosemirror-state';
import { undo } from 'prosemirror-history';
import { EditorView } from 'prosemirror-view';
import { DOMParser as PMDOMParser, DOMSerializer, type Node as PMNode } from 'prosemirror-model';
import { paraStyleCss, schema, type ParaAttrs } from '../doc/schema';
import { blankMeta } from '../doc/model';
import { createState } from '../doc/editor';
import {
  changeCase,
  cycleCase,
  hangingIndent,
  indent,
  setBorders,
  setListFormat,
  setListStart,
  setMultilevel,
  setSpace,
  setUnderline,
  sortParas,
  toggleSpaceBefore,
} from '../doc/commands';
import { joinStyles } from '../doc/plugins';
import { buildDocx, listNumbers, readDocx } from '../doc/docx';
import { toHtmlBody } from '../doc/convert';
import { BULLET_LIBRARY, MULTILEVEL_LIBRARY, listMarker } from '../doc/lists';
import { AR } from '../i18n.ar';
import { UNDERLINE_LABEL } from '../doc/components/FormatMenus';

const { nodes, marks } = schema;
const p = (text: string, attrs: Partial<ParaAttrs> = {}, m: ReturnType<typeof marks.bold.create>[] = []) =>
  nodes.paragraph.create(attrs, text ? schema.text(text, m) : undefined);
const docOf = (...blocks: PMNode[]) => nodes.doc.create(null, blocks);

function at(doc: PMNode, from: number, to = from) {
  const state = createState(doc);
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
}
function run(state: EditorState, cmd: Command) {
  let s = state;
  expect(cmd(s, (tr) => (s = s.apply(tr)))).toBe(true);
  return s;
}
const attrsOf = (s: EditorState, i: number) => s.doc.child(i).attrs as ParaAttrs;

describe('Lists: libraries, multilevel and numbering value', () => {
  it('draws library and multilevel formats, start values and continued numbering', () => {
    const doc = docOf(
      p('a', { list: 'ordered', listFmt: 'upper-roman' }),
      p('b', { list: 'ordered', listFmt: 'upper-roman' }),
      p('text'),
      p('c', { list: 'ordered', listFmt: 'upper-roman', listStart: 'continue' }),
      p('text'),
      p('x', { list: 'ordered', listFmt: 'ml-legal' }),
      p('x.1', { list: 'ordered', listFmt: 'ml-legal', level: 1 }),
      p('x.1.1', { list: 'ordered', listFmt: 'ml-legal', level: 2 }),
      p('y', { list: 'ordered', listFmt: 'ml-legal' }),
      p('text'),
      p('five', { list: 'ordered', listStart: 5, listFmt: 'lower-letter-paren' }),
      p('six', { list: 'ordered', listFmt: 'lower-letter-paren' }),
      p('check', { list: 'bullet', listFmt: 'check' }),
      p('ml', { list: 'bullet', listFmt: 'ml-bullets', level: 2 }),
      p('ar', { list: 'ordered', listFmt: 'hindi', level: 0 }),
    );
    expect(listNumbers(doc)).toEqual(['I.', 'II.', 'III.', '1.', '1.1.', '1.1.1.', '2.', 'e)', 'f)', '✓', '■', '١.']);
    expect(listMarker('ordered', 0, 2, 'arabic-alpha')).toBe('ب.');
  });

  it('applies a library format to the items at the same level of the list, with undo', () => {
    const doc = docOf(p('one', { list: 'bullet' }), p('sub', { list: 'bullet', level: 1 }), p('two', { list: 'bullet' }), p('after'));
    const s = run(at(doc, 2), setListFormat('bullet', 'arrow'));
    expect([0, 1, 2].map((i) => attrsOf(s, i).listFmt)).toEqual(['arrow', null, 'arrow']);
    expect(attrsOf(s, 3).list).toBeNull();
    const back = run(s, undo);
    expect(back.doc.eq(doc)).toBe(true);
    // "None" takes the paragraph out of the list.
    expect(attrsOf(run(at(doc, 2), setListFormat('bullet', null)), 0).list).toBeNull();
  });

  it('gives an indented item the format of its new level; multilevel keeps its format', () => {
    const doc = docOf(p('a', { list: 'ordered', listFmt: 'upper-roman' }), p('b', { list: 'ordered', listFmt: 'upper-letter', level: 1 }), p('c', { list: 'ordered', listFmt: 'upper-roman' }));
    const s = run(at(doc, 7), indent(1));
    expect(attrsOf(s, 2)).toMatchObject({ level: 1, listFmt: 'upper-letter' });
    const ml = run(at(doc, 2), setMultilevel('ml-legal'));
    expect([0, 1, 2].map((i) => attrsOf(ml, i).listFmt)).toEqual(['ml-legal', 'ml-legal', 'ml-legal']);
    expect(listNumbers(ml.doc)).toEqual(['1.', '1.1.', '2.']);
    const deeper = run(at(ml.doc, 2), indent(1));
    expect(attrsOf(deeper, 0).listFmt).toBe('ml-legal');
  });

  it('sets the numbering value; a new item after Enter counts on', () => {
    const doc = docOf(p('a', { list: 'ordered' }), p('b', { list: 'ordered' }));
    const s = run(at(doc, 5), setListStart(7));
    expect(listNumbers(s.doc)).toEqual(['1.', '7.']);
    expect(run(at(doc, 5), setListStart('continue')).doc.child(1).attrs.listStart).toBe('continue');
  });
});

describe('Paragraph formatting', () => {
  it('draws indents, hanging indents, exact line spacing, borders and shading', () => {
    const base = nodes.paragraph.create().attrs as ParaAttrs;
    expect(paraStyleCss({ ...base, indentLeft: 36, indentRight: 18, firstLine: -18 })).toBe(
      'padding-inline-start: 36pt; padding-inline-end: 18pt; text-indent: -18pt',
    );
    expect(paraStyleCss({ ...base, lineHeight: 14, lineRule: 'exact' })).toBe('line-height: 14pt');
    expect(paraStyleCss({ ...base, list: 'bullet', indentLeft: 54 })).toContain('--dw-marker: 36pt');
    const box = paraStyleCss({ ...base, borders: 'tb', shading: '#ffff00' });
    expect(box).toContain('border-top: 0.5pt solid #000');
    expect(box).toContain('border-bottom: 0.5pt solid #000');
    expect(box).not.toContain('border-left');
    expect(box).toContain('background-image: linear-gradient(#ffff00, #ffff00)');
    // The box moves out by the border gap, so the text stays where it was.
    expect(box).toContain('margin-inline-start: -4pt');
    expect(box).toContain('margin-bottom: 8pt');
  });

  it("drops the space between list items unless 'Don't add space…' is off, and joins bordered paragraphs", () => {
    const styles = (doc: PMNode) => {
      const out: string[] = [];
      joinStyles(doc, (_n, offset, j) => (out[doc.resolve(offset).index()] = `${j.cls}${j.style ? ` | ${j.style}` : ''}`));
      return Array.from({ length: doc.childCount }, (_, i) => out[i] ?? '');
    };
    expect(styles(docOf(p('a', { list: 'bullet' }), p('b', { list: 'bullet' }), p('after')))).toEqual(['dw-no-after', 'dw-no-before', '']);
    expect(styles(docOf(p('a', { list: 'bullet', contextual: false }), p('b', { list: 'bullet', contextual: false })))).toEqual(['', '']);
    expect(styles(docOf(p('a', { contextual: true }), p('b', { contextual: true })))).toEqual(['dw-no-after', 'dw-no-before']);
    // Other styles keep their space; a list item next to Normal text too.
    expect(styles(docOf(p('a', { contextual: true }), p('b', { style: 'h1', contextual: true }), p('c', { list: 'bullet' })))).toEqual(['', '', '']);
    expect(styles(docOf(p('a', { borders: 'tlbr' }), p('b', { borders: 'tlbr' }), p('c', { borders: 'b' })))).toEqual([
      'dw-join-down | --dw-join-bottom: 8pt',
      'dw-join-up | --dw-join-top: 0pt',
      '',
    ]);
  });

  it('adding space after list items turns off the same-style rule; Ctrl+0 toggles 12 pt before', () => {
    const doc = docOf(p('a', { list: 'bullet' }), p('b', { list: 'bullet' }));
    const s = run(at(doc, 1, 6), setSpace('spaceAfter', true));
    expect(attrsOf(s, 0)).toMatchObject({ spaceAfter: 8, contextual: false });
    const t = run(at(docOf(p('x')), 1), toggleSpaceBefore);
    expect(attrsOf(t, 0).spaceBefore).toBe(12);
    expect(attrsOf(run(t, toggleSpaceBefore), 0).spaceBefore).toBe(0);
  });

  it('Ctrl+T makes a hanging indent that keeps the first line in place', () => {
    const s = run(at(docOf(p('x')), 1), hangingIndent(1));
    expect(attrsOf(s, 0)).toMatchObject({ indentLeft: 36, firstLine: -36 });
    expect(attrsOf(run(s, hangingIndent(-1)), 0)).toMatchObject({ indentLeft: 0, firstLine: 0 });
  });

  it('turns borders on and off by side', () => {
    let s = run(at(docOf(p('x')), 1), setBorders('b'));
    expect(attrsOf(s, 0).borders).toBe('b');
    s = run(s, setBorders('t'));
    expect(attrsOf(s, 0).borders).toBe('tb');
    s = run(s, setBorders('box'));
    expect(attrsOf(s, 0).borders).toBe('tlbr');
    expect(attrsOf(run(s, setBorders('none')), 0).borders).toBeNull();
  });

  it('sorts paragraphs ascending and descending', () => {
    const doc = docOf(p('pear'), p('Apple'), p('item 10'), p('item 9'));
    const s = run(at(doc, 1, doc.content.size - 1), sortParas(false));
    expect(s.doc.children.map((n) => n.textContent)).toEqual(['Apple', 'item 9', 'item 10', 'pear']);
    const d = run(at(doc, 1, doc.content.size - 1), sortParas(true));
    expect(d.doc.child(0).textContent).toBe('pear');
  });
});

describe('Font formatting', () => {
  it('changes case keeping the formatting of each run, with undo', () => {
    const doc = docOf(p('hello WORLD. again'));
    const range = (cmd: Command) => run(at(doc, 1, 19), cmd).doc.textContent;
    expect(range(changeCase('upper'))).toBe('HELLO WORLD. AGAIN');
    expect(range(changeCase('lower'))).toBe('hello world. again');
    expect(range(changeCase('sentence'))).toBe('Hello world. Again');
    expect(range(changeCase('capitalize'))).toBe('Hello World. Again');
    expect(range(changeCase('toggle'))).toBe('HELLO world. AGAIN');
    const bold = docOf(p('ab', {}, [marks.bold.create({ value: true })]));
    const s = run(at(bold, 1, 3), changeCase('upper'));
    expect(marks.bold.isInSet(s.doc.child(0).child(0).marks)).toBeTruthy();
    expect(run(s, undo).doc.textContent).toBe('ab');
    // Shift+F3 on the word at the caret: lowercase → Capitalize → UPPERCASE → lowercase.
    let c = at(docOf(p('word here')), 2);
    c = run(c, cycleCase);
    expect(c.doc.textContent).toBe('Word here');
    c = run(c, cycleCase);
    expect(c.doc.textContent).toBe('WORD here');
    c = run(c, cycleCase);
    expect(c.doc.textContent).toBe('word here');
  });

  it('underlines with a style and colour', () => {
    const s = run(at(docOf(p('text')), 1, 5), setUnderline('wavy', '#ff0000'));
    expect(marks.underline.isInSet(s.doc.child(0).child(0).marks)?.attrs).toEqual({ style: 'wavy', color: '#ff0000' });
    const again = run(s, setUnderline('double'));
    expect(marks.underline.isInSet(again.doc.child(0).child(0).marks)?.attrs).toEqual({ style: 'double', color: '#ff0000' });
  });
});

describe('Word files: formatting round trip', () => {
  const run = (text: string, ms: ReturnType<typeof marks.bold.create>[]) => schema.text(text, ms);
  function sample() {
    return docOf(
      p('Roman one', { list: 'ordered', listFmt: 'upper-roman' }),
      p('Roman two', { list: 'ordered', listFmt: 'upper-roman' }),
      p('Between'),
      p('Roman three', { list: 'ordered', listFmt: 'upper-roman', listStart: 'continue' }),
      p('Legal', { list: 'ordered', listFmt: 'ml-legal' }),
      p('Legal sub', { list: 'ordered', listFmt: 'ml-legal', level: 1 }),
      p('Again'),
      p('Five', { list: 'ordered', listStart: 5, listFmt: 'decimal-paren' }),
      p('Checked', { list: 'bullet', listFmt: 'check', contextual: false, spaceAfter: 6 }),
      p('ML bullet', { list: 'bullet', listFmt: 'ml-bullets', level: 1 }),
      p('Indented', { indentLeft: 50, indentRight: 20, firstLine: -25 }),
      p('First line', { firstLine: 36 }),
      p('Exact', { lineRule: 'exact', lineHeight: 15 }),
      p('At least', { lineRule: 'atLeast', lineHeight: 18 }),
      p('Same style', { contextual: true }),
      p('Boxed', { borders: 'tlbr', shading: '#ddeeff' }),
      p('Bottom only', { borders: 'b' }),
      p('عربي', { dir: 'rtl', indentLeft: 72, firstLine: -18 }),
      nodes.paragraph.create(null, [
        run('caps ', [marks.caps.create()]),
        run('small ', [marks.smallCaps.create()]),
        run('dstrike ', [marks.dstrike.create()]),
        run('under ', [marks.underline.create({ style: 'double', color: '#ff0000' })]),
        run('dotted ', [marks.underline.create({ style: 'dotted' })]),
        run('wide', [marks.spacing.create({ value: 2 })]),
        run('narrow', [marks.spacing.create({ value: -1.5 })]),
      ]),
    );
  }

  it('keeps list formats, numbering values, indents, spacing, borders and character effects', async () => {
    const doc = sample();
    const { data } = await buildDocx(doc, blankMeta('F'));
    const back = (await readDocx(data, 'F')).doc;
    expect(listNumbers(back)).toEqual(listNumbers(doc));
    expect(listNumbers(back)).toEqual(['I.', 'II.', 'III.', '4.', '4.1.', '5)', '✓', '○']);
    const byText = (s: string) => {
      let found: PMNode | undefined;
      back.descendants((n) => {
        if (!found && n.type === nodes.paragraph && n.textContent === s) found = n;
        return !found;
      });
      return found!.attrs as ParaAttrs;
    };
    expect(byText('Roman one').listFmt).toBe('upper-roman');
    expect(byText('Roman three').listStart).toBe('continue');
    expect(byText('Legal sub')).toMatchObject({ listFmt: 'ml-legal', level: 1 });
    expect(byText('Five')).toMatchObject({ listStart: 5, listFmt: 'decimal-paren' });
    expect(byText('Checked')).toMatchObject({ listFmt: 'check', contextual: false, spaceAfter: 6 });
    expect(byText('ML bullet').listFmt).toBe('ml-bullets');
    expect(byText('Indented')).toMatchObject({ indentLeft: 50, indentRight: 20, firstLine: -25 });
    expect(byText('First line')).toMatchObject({ firstLine: 36, indentLeft: null });
    expect(byText('Exact')).toMatchObject({ lineRule: 'exact', lineHeight: 15 });
    expect(byText('At least')).toMatchObject({ lineRule: 'atLeast', lineHeight: 18 });
    expect(byText('Same style').contextual).toBe(true);
    expect(byText('Boxed')).toMatchObject({ borders: 'tlbr', shading: '#ddeeff' });
    expect(byText('Bottom only').borders).toBe('b');
    expect(byText('عربي')).toMatchObject({ dir: 'rtl', indentLeft: 72, firstLine: -18 });

    const runs = back.lastChild!.content.content;
    const m = (text: string) => runs.find((r) => r.text === text)!.marks;
    expect(marks.caps.isInSet(m('caps '))).toBeTruthy();
    expect(marks.smallCaps.isInSet(m('small '))).toBeTruthy();
    expect(marks.dstrike.isInSet(m('dstrike '))).toBeTruthy();
    expect(marks.underline.isInSet(m('under '))?.attrs).toEqual({ style: 'double', color: '#ff0000' });
    expect(marks.underline.isInSet(m('dotted '))?.attrs.style).toBe('dotted');
    expect(marks.spacing.isInSet(m('wide'))?.attrs.value).toBe(2);
    expect(marks.spacing.isInSet(m('narrow'))?.attrs.value).toBe(-1.5);
  });

  it('writes Word numbering, indents, spacing and run properties', async () => {
    const zip = await JSZip.loadAsync((await buildDocx(sample(), blankMeta('F'))).data);
    const num = await zip.file('word/numbering.xml')!.async('string');
    expect(new DOMParser().parseFromString(num, 'application/xml').getElementsByTagName('parsererror')).toHaveLength(0);
    expect(num).toContain('<w:numFmt w:val="upperRoman"/>');
    expect(num).toContain('<w:lvlText w:val="%1.%2."/>');
    expect(num).toContain('<w:lvlText w:val="%1)"/>');
    expect(num).toContain('<w:lvlText w:val="✓"/>');
    expect(num).toContain('<w:startOverride w:val="5"/>');
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagName('parsererror')).toHaveLength(0);
    expect(xml).toContain('<w:ind w:left="1000" w:right="400" w:hanging="500"/>');
    expect(xml).toContain('<w:ind w:firstLine="720"/>');
    expect(xml).toContain('<w:spacing w:line="300" w:lineRule="exact"/>');
    expect(xml).toContain('<w:spacing w:line="360" w:lineRule="atLeast"/>');
    expect(xml).toContain('<w:contextualSpacing w:val="0"/>');
    expect(xml).toContain('<w:contextualSpacing/>');
    expect(xml).toMatch(/<w:pBdr><w:top [^>]*\/><w:left [^>]*\/><w:bottom [^>]*\/><w:right [^>]*\/><\/w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"\/>/);
    expect(xml).toContain('<w:caps/>');
    expect(xml).toContain('<w:smallCaps/>');
    expect(xml).toContain('<w:dstrike/>');
    expect(xml).toContain('<w:u w:val="double" w:color="FF0000"/>');
    expect(xml).toContain('<w:spacing w:val="40"/>');
    expect(xml).toContain('<w:spacing w:val="-30"/>');
  });

  it("reads Word's own numbering: symbol-font bullets, start values, continued lists, indents and underline styles", async () => {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const zip = new JSZip();
    zip.file(
      '_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    );
    const lvl = (i: number, fmt: string, text: string, extra = '') =>
      `<w:lvl w:ilvl="${i}"><w:start w:val="${i === 0 && fmt === 'upperRoman' ? 3 : 1}"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/>` +
      `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr>${extra}</w:lvl>`;
    zip.file(
      'word/numbering.xml',
      `<w:numbering xmlns:w="${W}">` +
        `<w:abstractNum w:abstractNumId="0">${lvl(0, 'bullet', '', '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr>')}${lvl(1, 'bullet', 'o')}</w:abstractNum>` +
        `<w:abstractNum w:abstractNumId="1">${lvl(0, 'upperRoman', '%1.')}${lvl(1, 'lowerLetter', '%2)')}</w:abstractNum>` +
        `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
        `<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>` +
        `<w:num w:numId="3"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>` +
        `</w:numbering>`,
    );
    const item = (num: number, ilvl: number, text: string, pPr = '') =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${num}"/></w:numPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    zip.file(
      'word/document.xml',
      `<w:document xmlns:w="${W}"><w:body>` +
        item(1, 0, 'arrow') +
        item(1, 1, 'circle') +
        item(2, 0, 'three') +
        item(2, 1, 'sub', '<w:ind w:left="1800" w:hanging="720"/>') +
        `<w:p><w:r><w:t>text between</w:t></w:r></w:p>` +
        item(2, 0, 'four') +
        item(3, 0, 'restart') +
        `<w:p><w:pPr><w:spacing w:line="280" w:lineRule="atLeast"/><w:ind w:firstLine="567"/><w:pBdr><w:bottom w:val="single" w:sz="4"/><w:left w:val="nil"/></w:pBdr></w:pPr>` +
        `<w:r><w:rPr><w:u w:val="wave" w:color="00FF00"/></w:rPr><w:t>wavy</w:t></w:r><w:r><w:rPr><w:u w:val="dashLong"/></w:rPr><w:t>dash</w:t></w:r></w:p>` +
        `</w:body></w:document>`,
    );
    const { doc } = await readDocx(await zip.generateAsync({ type: 'uint8array' }), 'W');
    expect(listNumbers(doc)).toEqual(['➢', '◦', 'III.', 'a)', 'IV.', 'I.']);
    const a = (i: number) => doc.child(i).attrs as ParaAttrs;
    expect(a(0).listFmt).toBe('arrow');
    expect(a(1).listFmt).toBeNull(); // Word's "o" is the default second-level bullet
    expect(a(2)).toMatchObject({ listStart: 3, listFmt: 'upper-roman' });
    expect(a(3)).toMatchObject({ listFmt: 'lower-letter-paren', indentLeft: 90, firstLine: -36 });
    expect(a(5).listStart).toBe('continue');
    expect(a(6).listStart).toBe(1);
    expect(a(7)).toMatchObject({ lineRule: 'atLeast', lineHeight: 14, firstLine: 28.35, borders: 'b' });
    const runs = doc.child(7).content.content;
    expect(marks.underline.isInSet(runs[0].marks)?.attrs).toEqual({ style: 'wavy', color: '#00ff00' });
    expect(marks.underline.isInSet(runs[1].marks)?.attrs.style).toBe('dashed');
  });
});

describe('Next Writer files, copy and paste', () => {
  it('opens documents saved before these attributes existed', async () => {
    const { fromFile, toFile } = await import('../doc/store');
    const old = {
      format: 'nxdoc',
      version: 1,
      ...blankMeta('Old'),
      content: {
        type: 'doc',
        content: [
          { type: 'paragraph', attrs: { style: 'normal', align: null, dir: null, list: 'bullet', level: 0, lineHeight: null, spaceBefore: null, spaceAfter: null }, content: [{ type: 'text', text: 'item', marks: [{ type: 'underline' }] }] },
        ],
      },
    };
    const { doc } = fromFile(old, 'x');
    expect(doc.child(0).attrs).toMatchObject({ listFmt: null, listStart: null, indentLeft: null, contextual: null, borders: null });
    expect(marks.underline.isInSet(doc.child(0).child(0).marks)?.attrs).toEqual({ style: 'single', color: null });
    // New attributes survive a save.
    const rich = docOf(p('x', { list: 'ordered', listFmt: 'ml-legal', listStart: 3, borders: 'tb', shading: '#eeeeee', lineRule: 'exact', lineHeight: 12 }));
    const again = fromFile(JSON.parse(JSON.stringify(toFile(blankMeta('n'), rich))), 'n');
    expect(again.doc.eq(rich)).toBe(true);
  });

  it('keeps paragraph and character formatting through the clipboard (HTML)', () => {
    const doc = docOf(
      nodes.paragraph.create({ indentLeft: 40, firstLine: -20, spaceAfter: 3, contextual: true, borders: 'b', shading: '#ff0000', lineRule: 'exact', lineHeight: 13 }, [
        schema.text('a', [marks.underline.create({ style: 'wavy', color: '#00ff00' }), marks.smallCaps.create(), marks.spacing.create({ value: 1.5 })]),
        schema.text('b', [marks.dstrike.create()]),
      ]),
      p('n', { list: 'ordered', listFmt: 'upper-roman', listStart: 4 }),
    );
    const div = document.createElement('div');
    div.appendChild(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
    const back = PMDOMParser.fromSchema(schema).parse(div);
    expect(back.toJSON()).toEqual(doc.toJSON());
    // HTML export writes the list's look and start.
    expect(toHtmlBody(docOf(p('n', { list: 'ordered', listFmt: 'upper-roman', listStart: 4 })))).toMatch(/<ol style="list-style-type: upper-roman;" start="4">/);
  });
});

describe('Keyboard shortcuts and drawing in the editor', () => {
  function view(doc: PMNode, from: number, to = from) {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const v = new EditorView(host, { state: at(doc, from, to) });
    const key = (key: string, mods: { ctrl?: boolean; shift?: boolean } = {}) =>
      v.dom.dispatchEvent(
        new KeyboardEvent('keydown', { key, ctrlKey: !!mods.ctrl, shiftKey: !!mods.shift, keyCode: /^[a-z]$/i.test(key) ? key.toUpperCase().charCodeAt(0) : 0, bubbles: true, cancelable: true }),
      );
    return { v, key, done: () => (v.destroy(), host.remove()) };
  }

  it("runs Word's paragraph and character shortcuts", () => {
    const { v, key, done } = view(docOf(p('some words')), 1, 5);
    const a = () => v.state.doc.child(0).attrs as ParaAttrs;
    key('5', { ctrl: true });
    expect(a().lineHeight).toBe(1.5);
    key('2', { ctrl: true });
    expect(a().lineHeight).toBe(2);
    key('1', { ctrl: true });
    expect(a().lineHeight).toBe(1);
    key('0', { ctrl: true });
    expect(a().spaceBefore).toBe(12);
    key('t', { ctrl: true });
    expect(a()).toMatchObject({ indentLeft: 36, firstLine: -36 });
    key('m', { ctrl: true });
    expect(a().indentLeft).toBe(72);
    key('A', { ctrl: true, shift: true });
    expect(marks.caps.isInSet(v.state.doc.child(0).child(0).marks)).toBeTruthy();
    key('K', { ctrl: true, shift: true });
    expect(marks.smallCaps.isInSet(v.state.doc.child(0).child(0).marks)).toBeTruthy();
    key('D', { ctrl: true, shift: true });
    expect(marks.underline.isInSet(v.state.doc.child(0).child(0).marks)?.attrs.style).toBe('double');
    key('F3', { shift: true });
    expect(v.state.doc.textContent).toBe('Some words');
    key('L', { ctrl: true, shift: true });
    expect(a().list).toBe('bullet');
    done();
  });

  it('draws list markers and the spacing between list items', () => {
    const { v, done } = view(docOf(p('a', { list: 'ordered', listFmt: 'upper-roman' }), p('b', { list: 'ordered', listFmt: 'upper-roman' }), p('c')), 1);
    const ps = [...v.dom.querySelectorAll('.dw-p')] as HTMLElement[];
    expect(ps.map((e) => e.getAttribute('data-marker'))).toEqual(['I.', 'II.', null]);
    expect(ps[0].classList.contains('dw-no-after')).toBe(true);
    expect(ps[1].classList.contains('dw-no-before')).toBe(true);
    expect(ps[1].classList.contains('dw-no-after')).toBe(false);
    done();
  });
});

describe('Arabic interface for the new lists and menus', () => {
  it('translates library and menu names', () => {
    const names = [
      ...BULLET_LIBRARY.map((b) => b.name),
      ...Object.values(UNDERLINE_LABEL),
      ...['Restart at 1', 'Continue Numbering', 'Set Numbering Value…', 'Automatic', 'Sort Ascending (A to Z)', 'Sort Descending (Z to A)'],
      ...['Single', '1.5 lines', 'Double', 'At least', 'Exactly', 'Multiple', 'Centered', 'Justified', 'Before text:', 'After text:', 'Left:', 'Right:'],
      ...['Strikethrough', 'Double strikethrough', 'Superscript', 'Subscript', 'Small caps', 'All caps', 'Regular', 'Italic', 'Bold', 'Bold Italic'],
      ...['cm', 'pt', 'lines', 'Indent', 'Before text', 'After text'],
    ];
    expect(names.filter((n) => !(n in AR))).toEqual([]);
    expect(MULTILEVEL_LIBRARY.length).toBeGreaterThanOrEqual(3);
  });
});
