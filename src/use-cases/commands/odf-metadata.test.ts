import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { buildMalformedDocx, buildMinimalOdt, buildRichOdt } from '../../test-helpers/office-fixtures.ts';
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
      creator: 'Vincent',
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
        '- **generator**: LibreOffice/7.4.2\n- **title**: Q4 Plan\n- **creator**: Vincent\n- **description**: Internal draft\n- **initial-creator**: Alice\n- **creation-date**: 2026-05-01T10:00:00\n- **editing-cycles**: 7\n\n' +
        '### Keywords\n\n- budget\n- confidential\n\n' +
        '### User-defined properties\n\n| name | value |\n| --- | --- |\n| ClientID | ACME-42 |\n| Reviewer | Bob |\n'
    );
  });

  it('emits `_(none)_` for every section on a barebones package', async () => {
    const extracted = await extractOdfMetadata(await buildMinimalOdt());
    expect(extracted.ok).toBe(true);
    if (!extracted.ok) return;
    expect(formatOdfMetadata(extracted.value)).toBe(
      '## OpenDocument metadata\n\n### Document properties\n\n_(none)_\n\n### Keywords\n\n_(none)_\n\n### User-defined properties\n\n_(none)_\n'
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
