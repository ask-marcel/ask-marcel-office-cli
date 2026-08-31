import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import {
  buildDocxWithHeaderFooterTextbox,
  buildMacroDocm,
  buildMalformedDocx,
  buildMoveAndFormatDocx,
  buildRichDocx,
  buildSampleDocx,
  buildSideChannelDocx,
  buildTrackedChangesDocx,
} from '../../test-helpers/office-fixtures.ts';
import { formatDocxMetadata } from './docx-metadata-to-markdown.ts';
import { extractDocxMetadata } from './docx-metadata.ts';

/**
 * Hand-rolled DOCX zip carrying the XML shapes the `docx` package can't
 * synthesise from its public API: a `word/people.xml` (w15:person registry),
 * a `w:fldSimple` field (the SimpleField cousin, distinct from the
 * `w:instrText` form the docx package emits), and a run with multiple
 * sibling `<w:t>` children (the array shape collectText must flatten).
 * Keeping it in the test file because it's the only caller — promoting it
 * to test-helpers would be premature abstraction.
 */
const buildCraftedDocx = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:fldSimple w:instr="DOCVARIABLE Region"><w:r><w:t>cached</w:t></w:r></w:fldSimple>
    </w:p>
    <w:p>
      <w:r>
        <w:instrText> HYPERLINK "https://example.org/legacy" </w:instrText>
        <w:instrText>MERGEFIELD AccountManager</w:instrText>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:instrText>DOCVARIABLE OneOff</w:instrText></w:r>
    </w:p>
    <w:p>
      <w:ins w:id="200" w:author="Bob" w:date="2026-05-13T09:00:00Z">
        <w:r>
          <w:t>multi-one </w:t>
          <w:t>multi-two</w:t>
        </w:r>
      </w:ins>
    </w:p>
  </w:body>
</w:document>`
  );
  zip.file(
    'word/people.xml',
    `<?xml version="1.0"?>
<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w15:person w15:author="Alice Smith">
    <w15:presenceInfo w15:providerId="AD" w15:userId="alice@contoso.com"/>
  </w15:person>
</w15:people>`
  );
  const buffer = await zip.generateAsync({ type: 'uint8array' });
  return buffer;
};

describe('extractDocxMetadata', () => {
  it('returns every section populated for a rich docx — core/custom props, comment, tracked ins/del, hidden text, external rel, MERGEFIELD, bookmark', async () => {
    const bytes = await buildRichDocx();
    const result = await extractDocxMetadata(bytes);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const m = result.value;
    expect(m.core.creator).toBe('Jordan Avery');
    expect(m.core.title).toBe('Q4 Report');
    expect(m.custom).toContainEqual({ name: 'ClientID', value: 'ACME-42' });
    expect(m.custom).toContainEqual({ name: 'ReviewStatus', value: 'pending' });
    expect(m.comments).toHaveLength(1);
    expect(m.comments[0]?.author).toBe('Jordan Avery');
    expect(m.comments[0]?.text).toContain('Please double-check');
    // the comment is anchored to its commentRange span in document.xml
    expect(m.comments[0]?.anchor).toBe('the Q4 revenue figure');
    // The fixture's insertion and deletion are adjacent siblings by the same
    // author, which is a replacement: they land paired rather than as two loose
    // halves, and the loose lists are empty because that is the only revision.
    expect(m.replacements).toHaveLength(1);
    expect(m.replacements[0]?.before).toBe('deleted-phrase');
    expect(m.replacements[0]?.after).toBe('inserted-phrase');
    expect(m.replacements[0]?.author).toBe('Jordan Avery');
    expect(m.insertions).toEqual([]);
    expect(m.deletions).toEqual([]);
    expect(m.hiddenText.some((h) => h.includes('This is hidden.'))).toBe(true);
    expect(m.externalRels.some((r) => r.target === 'https://example.com/secret-portal')).toBe(true);
    expect(m.fields.some((f) => f.instruction.includes('CustomerName'))).toBe(true);
    expect(m.bookmarks.some((b) => b.name === 'BM_intro')).toBe(true);
  });

  it('captures header/footer body prose and text-box (w:txbxContent) text that mammoth drops', async () => {
    const result = await extractDocxMetadata(await buildDocxWithHeaderFooterTextbox());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const m = result.value;
    expect(m.textBoxes).toEqual(['Callout box text']);
    expect(m.headersFooters).toEqual([
      { part: 'word/header1.xml', text: 'Confidential draft' },
      { part: 'word/footer1.xml', text: 'Page footer note' },
    ]);
    // the regular body paragraph is mammoth's job — it must NOT leak into the text-box list
    expect(m.textBoxes).not.toContain('Body paragraph.');
  });

  it('pins every side-channel field and its empty/whitespace filtering: comment attrs, tracked id/author/date, hidden-text, bookmark name filter, trimmed/empty fields, text boxes, two-digit + decoy header parts', async () => {
    const result = await extractDocxMetadata(await buildSideChannelDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const m = result.value;

    // Comments — id + initials are otherwise unasserted anywhere
    expect(m.comments).toEqual([{ id: '5', author: 'Commenter', initials: 'CC', date: '2026-03-03T00:00:00Z', text: 'comment-body' }]);

    // Tracked changes — id/author/date pinned; the empty-text insertion is filtered out
    expect(m.insertions).toEqual([{ id: '20', author: 'InsAuthor', date: '2026-01-01T00:00:00Z', text: 'kept-ins' }]);
    expect(m.deletions).toEqual([{ id: '30', author: 'DelAuthor', date: '2026-02-02T00:00:00Z', text: 'kept-del' }]);

    // Hidden text — the empty-vanish run is filtered; the rPr-less run is never treated as hidden
    expect(m.hiddenText).toEqual(['secret-hidden']);

    // Bookmarks — the empty-name bookmark is filtered out
    expect(m.bookmarks).toEqual([{ id: '10', name: 'BM_named' }]);

    // Fields — instrText trimmed, whitespace-only instrText + empty w:fldSimple filtered, header field discovered
    expect(m.fields.map((f) => f.instruction).toSorted((a, b) => a.localeCompare(b))).toEqual(['DOCVARIABLE FS', 'MERGEFIELD Spaced', 'PAGE']);
    expect(m.fields.every((f) => f.instruction !== '')).toBe(true);
    expect(m.fields).toContainEqual({ source: 'word/document.xml', instruction: 'MERGEFIELD Spaced' });
    expect(m.fields).toContainEqual({ source: 'word/header1.xml', instruction: 'PAGE' });

    // Text boxes — trimmed; the whitespace-only box is filtered
    expect(m.textBoxes).toEqual(['box-text']);

    // Headers/footers — trimmed; whitespace-only header2 filtered; two-digit header10 found
    // (the `\d+` quantifier); ^/$-anchored regex rejects notword/* and *.xmlbak decoys
    expect(m.headersFooters).toHaveLength(3);
    expect(m.headersFooters).toContainEqual({ part: 'word/header1.xml', text: 'HeaderOneProse' });
    expect(m.headersFooters).toContainEqual({ part: 'word/header10.xml', text: 'HeaderTenProse' });
    expect(m.headersFooters).toContainEqual({ part: 'word/footer1.xml', text: 'FooterOneProse' });
    expect(m.headersFooters.some((h) => h.text.includes('DECOY'))).toBe(false);
  });

  it('returns empty arrays for every list section on a barebones docx with no side-channel content (no people, no comments, no tracked changes, no hidden text, no bookmarks)', async () => {
    const bytes = await buildSampleDocx();
    const result = await extractDocxMetadata(bytes);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.people).toEqual([]);
    expect(result.value.comments).toEqual([]);
    expect(result.value.insertions).toEqual([]);
    expect(result.value.deletions).toEqual([]);
    expect(result.value.hiddenText).toEqual([]);
    expect(result.value.externalRels).toEqual([]);
    expect(result.value.bookmarks).toEqual([]);
    expect(result.value.custom).toEqual([]);
  });

  it('returns an api_error Result when the docx zip is malformed (the openOoxmlZip try/catch in the infra adapter translates the JSZip throw to a Result.err)', async () => {
    const result = await extractDocxMetadata(buildMalformedDocx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('ooxml zip parse failed');
    }
  });

  it('extracts the people registry, w:fldSimple field instructions, and flattens multi-w:t runs from a docx carrying XML shapes the `docx` package cannot synthesise', async () => {
    const bytes = await buildCraftedDocx();
    const result = await extractDocxMetadata(bytes);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.people).toEqual([{ author: 'Alice Smith', providerId: 'AD', userId: 'alice@contoso.com' }]);
    expect(result.value.fields.some((f) => f.instruction === 'DOCVARIABLE Region')).toBe(true);
    expect(result.value.fields.some((f) => f.instruction.includes('HYPERLINK') && f.instruction.includes('example.org/legacy'))).toBe(true);
    // Multiple `<w:instrText>` children of the same `<w:r>` collapse to an
    // array in fast-xml-parser — exercises the array branch of findAllTexts.
    expect(result.value.fields.some((f) => f.instruction === 'MERGEFIELD AccountManager')).toBe(true);
    // Single `<w:instrText>` (no sibling) — exercises the non-array branch of findAllTexts.
    expect(result.value.fields.some((f) => f.instruction === 'DOCVARIABLE OneOff')).toBe(true);
    // Tracked-insertion flattening of a multi-`<w:t>` run — exercises the array
    // branch of collectText that single-`<w:t>` runs (the docx package's default
    // shape) leave untested.
    expect(result.value.insertions.some((i) => i.text === 'multi-one multi-two')).toBe(true);

    // Render the metadata block end-to-end so the renderer's people-row
    // formatting (which the rich-docx fixture's empty people array would skip)
    // is exercised against a populated registry.
    const rendered = formatDocxMetadata(result.value);
    expect(rendered).toContain('### People registry');
    expect(rendered).toContain('Alice Smith');
    expect(rendered).toContain('alice@contoso.com');
  });
});

describe('VBA macro detection (shared across docx / xlsx / pptx)', () => {
  it('flags the vbaProject.bin part of a macro-enabled document and renders a `### Macros (VBA)` warning', async () => {
    const result = await extractDocxMetadata(await buildMacroDocm());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.macros).toEqual(['word/vbaProject.bin']);
    const rendered = formatDocxMetadata(result.value);
    expect(rendered).toContain('### Macros (VBA)');
    expect(rendered).toContain('word/vbaProject.bin');
    expect(rendered).toContain('can execute code when opened');
  });

  it('reports no macros (and renders `_(none)_`) for a document with no vbaProject.bin', async () => {
    const result = await extractDocxMetadata(await buildSampleDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.macros).toEqual([]);
    expect(formatDocxMetadata(result.value)).toContain('### Macros (VBA)\n\n_(none)_');
  });
});

describe('extractDocxMetadata — replacement pairing', () => {
  // OOXML has no "replace" revision: Word records one as a deletion next to an
  // insertion. Reported as two loose halves, a single edit reads as an unrelated
  // cut plus an unrelated addition, and nothing says the two are the same edit.
  it('pairs a deletion immediately followed by an insertion into one replacement, and takes both halves out of the loose lists', async () => {
    const result = await extractDocxMetadata(await buildTrackedChangesDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.replacements).toContainEqual({
      deletionId: '1',
      insertionId: '2',
      author: 'Robin Chen',
      date: '2026-04-01T00:00:00Z',
      before: 'Q3',
      after: 'Q4',
    });
    expect(result.value.deletions.some((d) => d.id === '1')).toBe(false);
    expect(result.value.insertions.some((i) => i.id === '2')).toBe(false);
  });

  it('pairs the insertion-then-deletion order too, which is what Word writes when the cursor sat before the selection', async () => {
    const result = await extractDocxMetadata(await buildTrackedChangesDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.replacements).toContainEqual({
      deletionId: '4',
      insertionId: '3',
      author: 'Robin Chen',
      date: '2026-04-01T00:00:00Z',
      before: 'old-first',
      after: 'new-first',
    });
  });

  it('refuses to pair across untouched prose: a cut here and an addition later in the paragraph are two edits, not one replacement', async () => {
    const result = await extractDocxMetadata(await buildTrackedChangesDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.replacements.some((r) => r.before === 'far-del' || r.after === 'far-ins')).toBe(false);
    expect(result.value.deletions.some((d) => d.text === 'far-del')).toBe(true);
    expect(result.value.insertions.some((i) => i.text === 'far-ins')).toBe(true);
  });

  it('refuses to pair two authors, since one person deleting and another inserting is not one person replacing a span', async () => {
    const result = await extractDocxMetadata(await buildTrackedChangesDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.replacements.some((r) => r.before === 'hers-del')).toBe(false);
    expect(result.value.deletions.some((d) => d.text === 'hers-del')).toBe(true);
    expect(result.value.insertions.some((i) => i.text === 'theirs-ins')).toBe(true);
  });

  it('still pairs across a bookmark marker, which carries no visible text and so does not separate the two halves', async () => {
    const result = await extractDocxMetadata(await buildTrackedChangesDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.replacements).toContainEqual({
      deletionId: '9',
      insertionId: '10',
      author: 'Robin Chen',
      date: '2026-04-01T00:00:00Z',
      before: 'marked-del',
      after: 'marked-ins',
    });
  });

  it('leaves an insertion with nothing beside it in the insertions list', async () => {
    const result = await extractDocxMetadata(await buildTrackedChangesDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.insertions.some((i) => i.text === 'lonely-ins')).toBe(true);
    expect(result.value.replacements.some((r) => r.after === 'lonely-ins')).toBe(false);
  });
});
describe('extractDocxMetadata — moves and format changes', () => {
  it('joins a move by its range name rather than its text, so two moves of the same sentence stay two moves', async () => {
    const result = await extractDocxMetadata(await buildMoveAndFormatDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const joined = result.value.moves.filter((m) => m.halves === 'both');
    expect(joined).toHaveLength(2);
    expect(joined.map((m) => m.name).toSorted((a, b) => a.localeCompare(b))).toEqual(['move-alpha', 'move-beta']);
    expect(joined.every((m) => m.text === 'the same sentence')).toBe(true);
    expect(joined.every((m) => m.author === 'Robin Chen')).toBe(true);
  });

  it('still reports a move whose destination never arrived, marked as the half it is', async () => {
    const result = await extractDocxMetadata(await buildMoveAndFormatDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.moves).toContainEqual({
      name: 'move-orphan',
      author: 'Robin Chen',
      date: '2026-04-02T00:00:00Z',
      text: 'orphaned sentence',
      halves: 'from-only',
    });
  });

  it('names the properties that changed when one formatting tag replaces another', async () => {
    const result = await extractDocxMetadata(await buildMoveAndFormatDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const styled = result.value.formatChanges.find((f) => f.text === 'styled words');
    expect(styled?.scope).toBe('run');
    expect(styled?.author).toBe('Alex Kim');
    expect(styled?.properties.toSorted((a, b) => a.localeCompare(b))).toEqual(['w:b', 'w:i']);
  });

  // The case a tag-name comparison misses: w:color is on both sides and only
  // the value moved, so the paragraph reads as unchanged unless attributes count.
  it('catches a property whose tag is unchanged and whose value moved', async () => {
    const result = await extractDocxMetadata(await buildMoveAndFormatDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const recoloured = result.value.formatChanges.find((f) => f.text === 'recoloured words');
    expect(recoloured?.properties).toEqual(['w:color']);
  });

  // Two attributes per side is what makes the comparison order them at all. It
  // is also the shape a real font change takes, where ascii and hAnsi move together.
  it('compares a property carrying several attributes as one value, so a font swap is one changed property', async () => {
    const result = await extractDocxMetadata(await buildMoveAndFormatDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const refonted = result.value.formatChanges.find((f) => f.text === 'refonted words');
    expect(refonted?.properties).toEqual(['w:rFonts']);
  });

  it('reports a paragraph-property change against the paragraph text, scoped apart from run formatting', async () => {
    const result = await extractDocxMetadata(await buildMoveAndFormatDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const realigned = result.value.formatChanges.find((f) => f.text === 'realigned paragraph');
    expect(realigned?.scope).toBe('paragraph');
    expect(realigned?.properties).toEqual(['w:jc']);
  });

  it('invents no revision for a run that was simply formatted and never edited', async () => {
    const result = await extractDocxMetadata(await buildMoveAndFormatDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.formatChanges.some((f) => f.text === 'plain bold text')).toBe(false);
  });

  // Guard on the pairing rule next door: a moveTo carries w:t like an insertion
  // does, so a del/ins pairing that keyed on "opposite kind" would swallow it.
  it('keeps moves out of the insertion, deletion and replacement lists entirely', async () => {
    const result = await extractDocxMetadata(await buildMoveAndFormatDocx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.insertions).toEqual([]);
    expect(result.value.deletions).toEqual([]);
    expect(result.value.replacements).toEqual([]);
  });
});
