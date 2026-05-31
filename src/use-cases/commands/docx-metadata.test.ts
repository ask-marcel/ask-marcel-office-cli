import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { buildDocxWithHeaderFooterTextbox, buildMacroDocm, buildMalformedDocx, buildRichDocx, buildSampleDocx, buildSideChannelDocx } from '../../test-helpers/office-fixtures.ts';
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
    expect(m.core.creator).toBe('Vincent Delacourt');
    expect(m.core.title).toBe('Q4 Report');
    expect(m.custom).toContainEqual({ name: 'ClientID', value: 'ACME-42' });
    expect(m.custom).toContainEqual({ name: 'ReviewStatus', value: 'pending' });
    expect(m.comments).toHaveLength(1);
    expect(m.comments[0]?.author).toBe('Vincent Delacourt');
    expect(m.comments[0]?.text).toContain('Please double-check');
    expect(m.insertions.some((i) => i.text.includes('inserted-phrase'))).toBe(true);
    expect(m.deletions.some((d) => d.text.includes('deleted-phrase'))).toBe(true);
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
    expect(m.fields.map((f) => f.instruction).toSorted()).toEqual(['DOCVARIABLE FS', 'MERGEFIELD Spaced', 'PAGE']);
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
