import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { buildMalformedDocx, buildMinimalOdt, buildTrackedChangesOdt, buildRichOdt } from '../../test-helpers/office-fixtures.ts';
import { formatOdfMetadata } from './odf-metadata-to-markdown.ts';
import { extractOdfMetadata } from './odf-metadata.ts';
import { odfToMarkdown } from './odf-to-markdown.ts';

// A meta.xml exercising the extractor's skip/filter branches: office:meta
// carries an attribute (the @_ skip) and inter-element whitespace (the #text
// skip); a blank keyword and a name-less user-defined property must be filtered
// out; the keyword + user-defined tags must stay out of the flat property record.
const buildEdgeCaseOdt = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    'meta.xml',
    '<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">' +
      '<office:meta office:flag="x">\n  <dc:title>T</dc:title>\n  <meta:generator>plain</meta:generator>\n  <meta:keyword>real</meta:keyword><meta:keyword>   </meta:keyword>\n  <meta:user-defined meta:name="Has">v</meta:user-defined><meta:user-defined>orphan</meta:user-defined>\n  </office:meta>' +
      '</office:document-meta>'
  );
  return zip.generateAsync({ type: 'uint8array' });
};

// A meta.xml that pins the remaining filter branches: an empty-text property
// (`dc:subject`) that must NOT enter the flat record, a single `meta:user-defined`
// whose `meta:name` is empty (an object the walker keeps, so the name-filter must
// drop it) — which also proves user-defined tags never leak into the property map.
const buildEmptyFieldsOdt = async (): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    'meta.xml',
    '<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">' +
      '<office:meta><dc:title>T</dc:title><dc:subject></dc:subject><meta:keyword>kw</meta:keyword><meta:user-defined meta:name="">noname</meta:user-defined></office:meta>' +
      '</office:document-meta>'
  );
  return zip.generateAsync({ type: 'uint8array' });
};

describe('extractOdfMetadata', () => {
  it('excludes empty-text properties from the flat record, never folds user-defined into it, and drops an empty-name user-defined field', async () => {
    const result = await extractOdfMetadata(await buildEmptyFieldsOdt());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.properties).toEqual({ title: 'T' });
    expect(result.value.keywords).toEqual(['kw']);
    expect(result.value.userDefined).toEqual([]);
  });

  it('extracts Dublin Core + ODF meta properties, the keyword list, and user-defined custom fields', async () => {
    const result = await extractOdfMetadata(await buildRichOdt());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.properties).toEqual({
      generator: 'LibreOffice/7.4.2',
      title: 'Q4 Plan',
      creator: 'Jordan',
      description: 'Internal draft',
      'initial-creator': 'Alice',
      'creation-date': '2026-05-01T10:00:00',
      'editing-cycles': '7',
    });
    expect(result.value.keywords).toEqual(['budget', 'confidential']);
    expect(result.value.userDefined).toEqual([
      { name: 'ClientID', value: 'ACME-42' },
      { name: 'Reviewer', value: 'Bob' },
    ]);
  });

  it('returns empty sections for a package with no meta.xml', async () => {
    const result = await extractOdfMetadata(await buildMinimalOdt());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.properties).toEqual({});
    expect(result.value.keywords).toEqual([]);
    expect(result.value.userDefined).toEqual([]);
  });

  it('returns an api_error Result when the package is not a valid zip', async () => {
    const result = await extractOdfMetadata(buildMalformedDocx());
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type === 'api_error') expect(result.error.message).toContain('ooxml zip parse failed');
  });

  it('skips office:meta attributes and whitespace, drops blank keywords and name-less user-defined fields, and keeps keyword/user-defined tags out of the property record', async () => {
    const result = await extractOdfMetadata(await buildEdgeCaseOdt());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.properties).toEqual({ title: 'T', generator: 'plain' });
    expect(result.value.keywords).toEqual(['real']);
    expect(result.value.userDefined).toEqual([{ name: 'Has', value: 'v' }]);
  });
});

describe('formatOdfMetadata', () => {
  it('renders the full `## OpenDocument metadata` block with every section, exact rows, and ordered keywords', async () => {
    const extracted = await extractOdfMetadata(await buildRichOdt());
    expect(extracted.ok).toBe(true);
    if (!extracted.ok) return;
    expect(formatOdfMetadata(extracted.value)).toBe(
      '## OpenDocument metadata\n\n' +
        '### Document properties\n\n' +
        '- **generator**: LibreOffice/7.4.2\n- **title**: Q4 Plan\n- **creator**: Jordan\n- **description**: Internal draft\n- **initial-creator**: Alice\n- **creation-date**: 2026-05-01T10:00:00\n- **editing-cycles**: 7\n\n' +
        '### Keywords\n\n- budget\n- confidential\n\n' +
        '### User-defined properties\n\n| name | value |\n| --- | --- |\n| ClientID | ACME-42 |\n| Reviewer | Bob |\n\n' +
        '### Tracked changes — replacements\n\n_(none)_\n\n### Tracked changes — insertions\n\n_(none)_\n\n### Tracked changes — deletions\n\n_(none)_\n\n### Tracked changes — formatting\n\n_(none)_\n'
    );
  });

  it('renders the four tracked-change sections with exact rows, in the docx order and vocabulary', async () => {
    const extracted = await extractOdfMetadata(await buildTrackedChangesOdt());
    expect(extracted.ok).toBe(true);
    if (!extracted.ok) return;
    const out = formatOdfMetadata(extracted.value);
    expect(out).toContain(
      '### Tracked changes — replacements\n\n| deletionId | insertionId | author | date | before | after |\n| --- | --- | --- | --- | --- | --- |\n| ct3 | ct4 | Robin Chen | 2026-09-01T10:00:00 | Q3 | Q4 |\n\n'
    );
    expect(out).toContain(
      '### Tracked changes — insertions\n\n| id | author | date | text |\n| --- | --- | --- | --- |\n| ct1 | Robin Chen | 2026-09-01T10:00:00 | newly added |\n'
    );
    expect(out).toContain(
      '### Tracked changes — deletions\n\n| id | author | date | text |\n| --- | --- | --- | --- |\n| ct2 | Robin Chen | 2026-09-01T10:00:00 | obsolete  sentence here now |\n'
    );
    expect(out).toContain('### Tracked changes — formatting\n\n| author | date | text |\n| --- | --- | --- |\n| Alex Kim | 2026-09-01T11:30:00 | reformatted words |\n');
  });

  it('emits `_(none)_` for every section on a barebones package', async () => {
    const extracted = await extractOdfMetadata(await buildMinimalOdt());
    expect(extracted.ok).toBe(true);
    if (!extracted.ok) return;
    expect(formatOdfMetadata(extracted.value)).toBe(
      '## OpenDocument metadata\n\n### Document properties\n\n_(none)_\n\n### Keywords\n\n_(none)_\n\n### User-defined properties\n\n_(none)_\n\n### Tracked changes — replacements\n\n_(none)_\n\n### Tracked changes — insertions\n\n_(none)_\n\n### Tracked changes — deletions\n\n_(none)_\n\n### Tracked changes — formatting\n\n_(none)_\n'
    );
  });
});

describe('odfToMarkdown', () => {
  it('converts the body to a text/markdown envelope and omits the metadata block by default', async () => {
    const result = await odfToMarkdown(await buildRichOdt());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.contentType).toBe('text/markdown');
    expect(result.value.size).toBe(new TextEncoder().encode(result.value.text).byteLength);
    expect(result.value.text).toContain('# Heading One');
    expect(result.value.text).toContain('Final paragraph.');
    expect(result.value.text).not.toContain('## OpenDocument metadata');
  });

  it('appends the `## OpenDocument metadata` block after the body when includeMetadata is true', async () => {
    const result = await odfToMarkdown(await buildRichOdt(), { includeMetadata: true });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.text).toContain('# Heading One');
    expect(result.value.text).toContain('## OpenDocument metadata');
    expect(result.value.text).toContain('Q4 Plan');
    // body precedes the metadata block
    expect(result.value.text.indexOf('# Heading One')).toBeLessThan(result.value.text.indexOf('## OpenDocument metadata'));
  });

  it('returns the metadata block alone when the body is empty (no content.xml) and includeMetadata is true', async () => {
    const result = await odfToMarkdown(await buildMinimalOdt(), { includeMetadata: true });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.text.startsWith('## OpenDocument metadata')).toBe(true);
  });

  it('propagates the zip-parse error for a malformed package', async () => {
    const result = await odfToMarkdown(buildMalformedDocx());
    expect(result.ok).toBe(false);
  });
});

describe('extractOdfMetadata — tracked changes', () => {
  // ODF declares each change once under text:tracked-changes and anchors it in
  // the body with marks; an insertion's text is what sits between its marks.
  it('reads an insertion from the text between its change marks, with the author and date of its region', async () => {
    const r = await extractOdfMetadata(await buildTrackedChangesOdt());
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.value.insertions).toContainEqual({ id: 'ct1', author: 'Robin Chen', date: '2026-09-01T10:00:00', text: 'newly added' });
    expect(r.value.insertions).toContainEqual({ id: 'ct9', author: 'Robin Chen', date: '2026-09-01T10:00:00', text: 'far-ins' });
  });

  it('reads a deletion from the paragraphs kept inside its region, since the body only holds a mark', async () => {
    const r = await extractOdfMetadata(await buildTrackedChangesOdt());
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    // Two counted spaces, a bare text:s, a tab read as one space, and the
    // reviewer's annotation left out: the deleted text as a reader saw it.
    expect(r.value.deletions).toContainEqual({ id: 'ct2', author: 'Robin Chen', date: '2026-09-01T10:00:00', text: 'obsolete  sentence here now' });
    expect(r.value.deletions.map((d) => d.id)).toEqual(['ct2', 'ct6', 'ct8']);
  });

  it('pairs a deletion mark immediately followed by an insertion by the same author into one replacement', async () => {
    const r = await extractOdfMetadata(await buildTrackedChangesOdt());
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.value.replacements).toEqual([{ deletionId: 'ct3', insertionId: 'ct4', author: 'Robin Chen', date: '2026-09-01T10:00:00', before: 'Q3', after: 'Q4' }]);
    expect(r.value.deletions.some((d) => d.id === 'ct3')).toBe(false);
    expect(r.value.insertions.some((i) => i.id === 'ct4')).toBe(false);
  });

  it('refuses to pair across authors or across untouched prose, the same rule as docx', async () => {
    const r = await extractOdfMetadata(await buildTrackedChangesOdt());
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.value.insertions.map((i) => i.id)).toEqual(['ct1', 'ct7', 'ct9']);
  });

  it('reports a format change with who, when and the text it covers; ODF records no property names', async () => {
    const r = await extractOdfMetadata(await buildTrackedChangesOdt());
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.value.formatChanges).toEqual([{ author: 'Alex Kim', date: '2026-09-01T11:30:00', text: 'reformatted words' }]);
  });

  it('leaves every tracked-change list empty for a document that has none', async () => {
    const r = await extractOdfMetadata(await buildRichOdt());
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect([r.value.insertions, r.value.deletions, r.value.replacements, r.value.formatChanges]).toEqual([[], [], [], []]);
  });
});
