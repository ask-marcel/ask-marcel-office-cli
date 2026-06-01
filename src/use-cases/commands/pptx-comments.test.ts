import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { extractCommentAuthors, extractComments } from './pptx-comments.ts';

const P_NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const RELS_NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const COMMENTS_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const LAYOUT_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';

const cm = (text: string): string => `<p:cm authorId="0" dt="2026-01-01T00:00:00Z"><p:text>${text}</p:text></p:cm>`;
const cmLst = (text: string): string => `<?xml version="1.0"?><p:cmLst ${P_NS}>${cm(text)}</p:cmLst>`;
const slide = (): string => `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree/></p:cSld></p:sld>`;
const rel = (type: string, target: string): string => `<?xml version="1.0"?><Relationships ${RELS_NS}><Relationship Id="rId1" Type="${type}" Target="${target}"/></Relationships>`;

// A deck exercising every branch of the slide↔comment correlation:
// - slide1 references comment1 via a `comments` rel (→ anchored) and also embeds a
//   stray <p:cm> that must be ignored (only ppt/comments/*.xml are scanned);
// - slide2 references comment2 via a NON-comments (slideLayout) rel (→ unanchored);
// - notesSlide1 references comment2 via a `comments` rel, but notes parts are NOT
//   scanned for slides, so comment2 still stays unanchored;
// - slide10 (multi-digit) references comment10 via a comments rel (→ anchored).
const buildDeck = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('ppt/commentAuthors.xml', `<?xml version="1.0"?><p:cmAuthorLst ${P_NS}><p:cmAuthor id="0" name="Author Zero" initials="AZ"/></p:cmAuthorLst>`);
  zip.file('ppt/comments/comment1.xml', cmLst('anchored'));
  zip.file('ppt/comments/comment2.xml', cmLst('decoy'));
  zip.file('ppt/comments/comment10.xml', cmLst('tenth'));
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0"?><p:sld ${P_NS}><p:cSld><p:spTree/></p:cSld>${cm('STRAY')}</p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', rel(COMMENTS_TYPE, '../comments/comment1.xml'));
  zip.file('ppt/slides/slide2.xml', slide());
  zip.file('ppt/slides/_rels/slide2.xml.rels', rel(LAYOUT_TYPE, '../comments/comment2.xml'));
  zip.file('ppt/notesSlides/notesSlide1.xml', `<?xml version="1.0"?><p:notes ${P_NS}><p:cSld><p:spTree/></p:cSld></p:notes>`);
  zip.file('ppt/notesSlides/_rels/notesSlide1.xml.rels', rel(COMMENTS_TYPE, '../comments/comment2.xml'));
  zip.file('ppt/slides/slide10.xml', slide());
  zip.file('ppt/slides/_rels/slide10.xml.rels', rel(COMMENTS_TYPE, '../comments/comment10.xml'));
  return zip.generateAsync({ type: 'uint8array' });
};

describe('extractComments — slide anchoring', () => {
  it('anchors via the slide’s comments rel, ignores non-comments rels + non-slide parts + stray comment-like elements, and resolves multi-digit slide numbers', async () => {
    const zipR = await openOoxmlZip(await buildDeck());
    expect(zipR.ok).toBe(true);
    if (!zipR.ok) return;
    const comments = extractComments(zipR.value, extractCommentAuthors(zipR.value));
    const bytext = (t: string): string | undefined => comments.find((c) => c.text === t)?.slide;
    // only the three real ppt/comments/*.xml parts are picked up — the stray <p:cm> inside slide1 is NOT
    expect(comments.map((c) => c.text).toSorted()).toEqual(['anchored', 'decoy', 'tenth']);
    expect(bytext('anchored')).toBe('slide1.xml'); // comments-typed rel anchors
    expect(bytext('decoy')).toBeUndefined(); // a slideLayout rel and a notesSlide's comments rel both must NOT anchor
    expect(bytext('tenth')).toBe('slide10.xml'); // slide\d+ matches a two-digit slide number
  });
});
