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
    expect(comments.map((c) => c.text).toSorted((a, b) => a.localeCompare(b))).toEqual(['anchored', 'decoy', 'tenth']);
    expect(bytext('anchored')).toBe('slide1.xml'); // comments-typed rel anchors
    expect(bytext('decoy')).toBeUndefined(); // a slideLayout rel and a notesSlide's comments rel both must NOT anchor
    expect(bytext('tenth')).toBe('slide10.xml'); // slide\d+ matches a two-digit slide number
  });
});

const Q_NS = 'xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main"';
const X188_NS = 'xmlns:x188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"';
const DM_NS = 'xmlns:dm="http://schemas.openxmlformats.org/drawingml/2006/main"';

// A deck written by a tool that binds the OOXML namespaces to its own prefixes: `q:`
// for presentationml, `x188:` for the 2018/8 modern-comment schema, `dm:` for DrawingML.
// A prefix is a per-document alias for a namespace URI, not part of the schema, so the
// extractor must key on local names, never on the qualified name PowerPoint happens to write.
const buildForeignPrefixDeck = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('ppt/commentAuthors.xml', `<?xml version="1.0"?><q:cmAuthorLst ${Q_NS}><q:cmAuthor id="0" name="Author Zero" initials="AZ"/></q:cmAuthorLst>`);
  zip.file('ppt/authors.xml', `<?xml version="1.0"?><x188:authorLst ${X188_NS}><x188:author id="{G1}" name="Robin Chen" initials="RC"/></x188:authorLst>`);
  zip.file('ppt/comments/comment1.xml', `<?xml version="1.0"?><q:cmLst ${Q_NS}><q:cm authorId="0" dt="2026-01-01T00:00:00Z"><q:text>legacy body</q:text></q:cm></q:cmLst>`);
  zip.file(
    'ppt/comments/modernComment1.xml',
    `<?xml version="1.0"?><x188:cmLst ${X188_NS} ${DM_NS}><x188:cm authorId="{G1}" created="2026-02-02T00:00:00Z"><x188:txBody><dm:bodyPr/><dm:p><dm:r><dm:t>modern body</dm:t></dm:r></dm:p></x188:txBody></x188:cm></x188:cmLst>`
  );
  return zip.generateAsync({ type: 'uint8array' });
};

describe('extractComments — a namespace prefix is a per-document alias, not part of the schema', () => {
  it('reads both comment generations from a deck whose writer bound its own prefixes', async () => {
    const zipR = await openOoxmlZip(await buildForeignPrefixDeck());
    expect(zipR.ok).toBe(true);
    if (!zipR.ok) return;
    const authors = extractCommentAuthors(zipR.value);
    const comments = extractComments(zipR.value, authors);
    expect(authors.map((a) => a.name).toSorted((a, b) => a.localeCompare(b))).toEqual(['Author Zero', 'Robin Chen']);
    // author resolved through both id schemes (integer + GUID), body read through both
    // element names (q:text + dm:t), date read through both attributes (dt + created)
    expect(comments.map((c) => `${c.author}|${c.date}|${c.text}`).toSorted((a, b) => a.localeCompare(b))).toEqual([
      'Author Zero|2026-01-01T00:00:00Z|legacy body',
      'Robin Chen|2026-02-02T00:00:00Z|modern body',
    ]);
  });
});

const P188_NS = 'xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"';
const P15_NS = 'xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"';

// PowerPoint stamps the commenter's directory identity as `S::<UPN>::<Entra object id>`:
// on the modern <author> element directly, but on the legacy <cmAuthor> only through a
// <presenceInfo> extension nested two levels down. A third-party writer may store the
// address bare, and a non-AD provider stamps a display name that is no address at all.
const buildIdentityDeck = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    'ppt/commentAuthors.xml',
    `<?xml version="1.0"?><p:cmAuthorLst ${P_NS}><p:cmAuthor id="0" name="Robin Chen" initials="RC"><p:extLst><p:ext uri="{19B8F6BF-5375-455C-9EA6-DF929625EA0E}"><p15:presenceInfo ${P15_NS} userId="S::robin.chen@contoso.com::cccccccc-3333-4333-8333-cccccccccccc" providerId="AD"/></p:ext></p:extLst></p:cmAuthor></p:cmAuthorLst>`
  );
  zip.file(
    'ppt/authors.xml',
    `<?xml version="1.0"?><p188:authorLst ${P188_NS}>` +
      '<p188:author id="{A1}" name="Alex Kim" initials="AK" userId="S::alex.kim@contoso.com::dddddddd-4444-4444-8444-dddddddddddd" providerId="AD"/>' +
      '<p188:author id="{A2}" name="Jordan Avery" initials="JA" userId="Jordan Avery" providerId="None"/>' +
      '<p188:author id="{A3}" name="Sam Rivera" initials="SR" userId="sam.rivera@contoso.com" providerId="AD"/>' +
      '</p188:authorLst>'
  );
  return zip.generateAsync({ type: 'uint8array' });
};

describe('extractCommentAuthors — the commenter’s directory identity', () => {
  it('reads the address wherever the writer put it, and reports none when the provider stamped no address', async () => {
    const zipR = await openOoxmlZip(await buildIdentityDeck());
    expect(zipR.ok).toBe(true);
    if (!zipR.ok) return;
    const identities = extractCommentAuthors(zipR.value).map((a) => `${a.name}|${a.email}`);
    expect(identities.toSorted((a, b) => a.localeCompare(b))).toEqual([
      'Alex Kim|alex.kim@contoso.com', // modern: S:: triple on the element
      'Jordan Avery|', // non-AD provider: a display name, not an address
      'Robin Chen|robin.chen@contoso.com', // legacy: S:: triple inside the presenceInfo extension
      'Sam Rivera|sam.rivera@contoso.com', // a writer that stored the address bare
    ]);
  });
});
