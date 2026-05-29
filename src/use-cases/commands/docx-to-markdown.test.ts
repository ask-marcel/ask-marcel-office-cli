import { describe, expect, it } from 'bun:test';
import { buildMalformedDocx, buildRichDocx, buildSampleDocx } from '../../test-helpers/office-fixtures.ts';
import { docxToMarkdown, promoteFirstRowToThead } from './docx-to-markdown.ts';

describe('docxToMarkdown', () => {
  it('converts a docx into a markdown envelope with heading, bold/italic, table, and inline image as data URI', async () => {
    const bytes = await buildSampleDocx();
    const result = await docxToMarkdown(bytes);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.contentType).toBe('text/markdown');
      expect(result.value.size).toBeGreaterThan(0);
      expect(result.value.text).toContain('# Sample Heading');
      expect(result.value.text).toContain('**world**');
      expect(result.value.text).toContain('_italic_');
      expect(result.value.text).toMatch(/\|\s*A\s*\|\s*B\s*\|/);
      expect(result.value.text).toContain('data:image/png;base64,');
    }
  });

  it('propagates the api_error from the mammoth adapter when the input is not a valid docx', async () => {
    const result = await docxToMarkdown(buildMalformedDocx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('docx conversion failed');
    }
  });

  it('defaults to NOT appending the metadata block when the includeMetadata flag is omitted (backward-compat — existing callers must see the same envelope)', async () => {
    const bytes = await buildRichDocx();
    const result = await docxToMarkdown(bytes);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.text).not.toContain('## DOCX metadata');
      expect(result.value.text).not.toContain('Core properties');
    }
  });

  it('defaults to NOT appending the metadata block when the includeMetadata flag is explicitly false', async () => {
    const bytes = await buildRichDocx();
    const result = await docxToMarkdown(bytes, { includeMetadata: false });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.text).not.toContain('## DOCX metadata');
  });
});

describe('docxToMarkdown — with --include-metadata true', () => {
  it('appends a populated metadata block surfacing every side-channel surface mammoth drops (core/app/custom props, comment, tracked ins/del, hidden text, external link, MERGEFIELD, bookmark)', async () => {
    const bytes = await buildRichDocx();
    const result = await docxToMarkdown(bytes, { includeMetadata: true });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const text = result.value.text;

    // Body still converts via mammoth
    expect(text).toContain('# Sample Heading');
    // Metadata block appears as a separate appended section
    expect(text).toContain('## DOCX metadata');

    // Core properties — creator, title, subject visible
    expect(text).toContain('### Core properties');
    expect(text).toContain('**creator**: Vincent Delacourt');
    expect(text).toContain('**title**: Q4 Report');

    // Custom properties — embedded as a markdown table
    expect(text).toContain('### Custom properties');
    expect(text).toMatch(/\|\s*ClientID\s*\|\s*ACME-42\s*\|/);

    // Comments — author + body visible
    expect(text).toContain('### Comments');
    expect(text).toContain('Vincent Delacourt');
    expect(text).toContain('Please double-check this figure.');

    // Tracked changes — insertion + deletion text both appear under their own sections
    expect(text).toContain('### Tracked changes — insertions');
    expect(text).toContain('inserted-phrase');
    expect(text).toContain('### Tracked changes — deletions');
    expect(text).toContain('deleted-phrase');

    // Hidden text (w:vanish) — surfaced verbatim, even though mammoth strips it from the rendered body
    expect(text).toContain('### Hidden-formatted text (w:vanish)');
    expect(text).toContain('This is hidden.');

    // External hyperlink — appears as both an "External relationships" row AND a HYPERLINK field instruction
    expect(text).toContain('### External relationships');
    expect(text).toContain('https://example.com/secret-portal');

    // Fields — MERGEFIELD CustomerName + HYPERLINK
    expect(text).toContain('### Fields (MERGEFIELD / HYPERLINK / DOCVARIABLE)');
    expect(text).toContain('CustomerName');

    // Bookmarks
    expect(text).toContain('### Bookmarks');
    expect(text).toContain('BM_intro');
  });

  it('emits `_(none)_` placeholders for every empty section so the output is grep-stable even on a barebones docx', async () => {
    const bytes = await buildSampleDocx();
    const result = await docxToMarkdown(bytes, { includeMetadata: true });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const text = result.value.text;
    expect(text).toContain('## DOCX metadata');
    // No comments, tracked changes, hidden text, bookmarks, or external rels in the simple fixture
    expect(text).toContain('### Comments\n\n_(none)_');
    expect(text).toContain('### Tracked changes — insertions\n\n_(none)_');
    expect(text).toContain('### Tracked changes — deletions\n\n_(none)_');
    expect(text).toContain('### Hidden-formatted text (w:vanish)\n\n_(none)_');
    expect(text).toContain('### External relationships\n\n_(none)_');
  });

  it('propagates the zip-parse error when --include-metadata true is requested on a malformed docx (the same Result.err path the body conversion takes)', async () => {
    const result = await docxToMarkdown(buildMalformedDocx(), { includeMetadata: true });
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      // mammoth fails first on a corrupted zip — that error wins over the metadata-extractor zip error
      expect(result.error.message).toContain('docx conversion failed');
    }
  });
});

describe('promoteFirstRowToThead — table walker edge cases', () => {
  it('returns HTML unchanged when there is no `<table>` at all', () => {
    const input = '<h1>title</h1><p>just text, no tables</p>';
    expect(promoteFirstRowToThead(input)).toBe(input);
  });

  it('returns HTML unchanged when a `<table>` opens but never closes (malformed)', () => {
    const input = '<p>before</p><table><tr><td>A</td></tr><p>truncated';
    expect(promoteFirstRowToThead(input)).toBe(input);
  });

  it('promotes the first row when the table is wrapped in `<tbody>` (defensive — mammoth does not currently emit this, but the walker handles it)', () => {
    const input = '<table><tbody><tr><td>H</td></tr><tr><td>1</td></tr></tbody></table>';
    expect(promoteFirstRowToThead(input)).toBe('<table><thead><tr><td>H</td></tr></thead><tbody><tr><td>1</td></tr></tbody></table>');
  });

  it('returns the original table when there is no `</tr>` inside (degenerate empty table)', () => {
    const input = '<table></table>';
    expect(promoteFirstRowToThead(input)).toBe('<table></table>');
  });

  it('walks past content between two tables and promotes both', () => {
    const input = '<table><tr><td>A</td></tr></table>middle<table><tr><td>B</td></tr></table>';
    const out = promoteFirstRowToThead(input);
    expect(out).toContain('<thead><tr><td>A</td></tr></thead>');
    expect(out).toContain('middle');
    expect(out).toContain('<thead><tr><td>B</td></tr></thead>');
  });
});
