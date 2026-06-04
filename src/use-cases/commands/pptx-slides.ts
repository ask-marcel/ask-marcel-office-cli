import { posix } from 'node:path';
import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { attrOf, collectText, findAll, parseXml } from './ooxml-xml-walker.ts';

/**
 * Per-slide extraction: the hidden flag (`<p:sld show="0">`), the title
 * placeholder text, and the speaker notes resolved through the slide's
 * relationship part. Speaker notes are the headline pptx side-channel —
 * presenter-authored text that never appears on the rendered slide.
 */

type Slide = { readonly name: string; readonly hidden: boolean; readonly title: string; readonly notes: string; readonly text: string };

const SLIDE_RE = /^ppt\/slides\/slide(\d+)\.xml$/;

const slideNumber = (path: string): number => {
  const m = SLIDE_RE.exec(path);
  return m === null ? 0 : Number(m[1]);
};

const slideTitle = (slide: unknown): string => {
  for (const sp of findAll(slide, 'p:sp')) {
    const ph = findAll(sp, 'p:ph')[0];
    if (ph === undefined) continue;
    const type = attrOf(ph, 'type');
    if (type === 'title' || type === 'ctrTitle') return collectText(sp, 'a:t');
  }
  return '';
};

const relsPathFor = (partPath: string): string => `${posix.dirname(partPath)}/_rels/${posix.basename(partPath)}.rels`;

const notesPathFor = (zip: OoxmlZip, slidePath: string): string | undefined => {
  const rels = parseXml(zip.read(relsPathFor(slidePath)));
  for (const rel of findAll(rels, 'Relationship')) {
    if (attrOf(rel, 'Type').endsWith('notesSlide')) return posix.normalize(posix.join(posix.dirname(slidePath), attrOf(rel, 'Target')));
  }
  return undefined;
};

const notesText = (zip: OoxmlZip, slidePath: string): string => {
  const notesPath = notesPathFor(zip, slidePath);
  if (notesPath === undefined) return '';
  const root = parseXml(zip.read(notesPath));
  return findAll(root, 'a:p')
    .map((p) => collectText(p, 'a:t'))
    .filter((t) => t.trim() !== '')
    .join(' ');
};

// Every visible text paragraph on the slide (title placeholder, body bullets,
// text boxes, table cells), one line per `<a:p>`, in document order. Document
// order is not guaranteed visual reading order — the lossy-flatten caveat.
const slideBodyText = (slideRoot: unknown): string =>
  findAll(slideRoot, 'a:p')
    .map((p) => collectText(p, 'a:t'))
    .filter((t) => t.trim() !== '')
    .join('\n');

const toSlide = (zip: OoxmlZip, slidePath: string): Slide => {
  const root = parseXml(zip.read(slidePath));
  const sld = findAll(root, 'p:sld')[0];
  return {
    name: posix.basename(slidePath),
    hidden: sld !== undefined && attrOf(sld, 'show') === '0',
    title: slideTitle(root),
    notes: notesText(zip, slidePath),
    text: slideBodyText(root),
  };
};

const extractSlides = (zip: OoxmlZip): ReadonlyArray<Slide> =>
  zip
    .list()
    .filter((p) => SLIDE_RE.test(p))
    .sort((a, b) => slideNumber(a) - slideNumber(b))
    .map((p) => toSlide(zip, p));

export { extractSlides };
export type { Slide };
