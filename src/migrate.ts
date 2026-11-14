import { DOC_VERSION, type Presentation, type SlideElement } from './types';
import { parasFromText } from './richtext';

/**
 * Upgrade documents and elements saved by older versions.
 * v1 → v2: `text: string` became rich-text `paras`.
 */
export function migrateElement(el: SlideElement): SlideElement {
  const legacy = el as SlideElement & { text?: unknown };
  if (!Array.isArray(legacy.paras)) {
    legacy.paras = parasFromText(typeof legacy.text === 'string' ? legacy.text : '');
  }
  delete legacy.text;
  return el;
}

export function migrateDoc(doc: Presentation): Presentation {
  if ((doc.version ?? 1) >= DOC_VERSION) return doc;
  for (const slide of doc.slides) slide.elements.forEach(migrateElement);
  doc.version = DOC_VERSION;
  return doc;
}
