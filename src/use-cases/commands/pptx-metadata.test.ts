import { describe, expect, it } from 'bun:test';
import JSZip from 'jszip';
import { buildMalformedPptx, buildMinimalPptx, buildRichPptx } from '../../test-helpers/office-fixtures.ts';
import { formatPptxMetadata } from './pptx-metadata-to-markdown.ts';
import { extractPptxMetadata } from './pptx-metadata.ts';

describe('extractPptxMetadata', () => {
  it('collects p:tag entries ONLY from ppt/tags/tag{N}.xml parts — the ^/$ anchored path filter excludes look-alike paths and p:tags elsewhere in the package', async () => {
    const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const tagPart = (name: string, val: string): string => `<?xml version="1.0"?><p:tags xmlns:p="${P}"><p:tag name="${name}" val="${val}"/></p:tags>`;
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
    zip.file('ppt/tags/tag1.xml', tagPart('KEEP', 'kept')); // a real single-digit tag part
    zip.file('ppt/tags/tag10.xml', tagPart('KEEP10', 'kept10')); // a real two-digit tag part (the \d+ quantifier must match it)
    zip.file('notppt/tags/tag1.xml', tagPart('NOCARET', 'x')); // matches only if the leading ^ is dropped
    zip.file('ppt/tags/tag1.xmlbak', tagPart('NODOLLAR', 'x')); // matches only if the trailing $ is dropped
    zip.file('ppt/foo/decoy.xml', tagPart('FILTERED', 'x')); // a p:tag outside the tags folder — only the no-filter mutant reads it
    const result = await extractPptxMetadata(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.slideTags).toEqual([
      { source: 'ppt/tags/tag1.xml', name: 'KEEP', value: 'kept' },
      { source: 'ppt/tags/tag10.xml', name: 'KEEP10', value: 'kept10' },
    ]);
  });

  it('surfaces the authored-but-invisible content a slide PDF never shows — custom props, slide tag, legacy + modern comments, hidden slide, speaker notes, external hyperlink', async () => {
    const result = await extractPptxMetadata(await buildRichPptx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const m = result.value;

    expect(m.core.creator).toBe('Vincent Delacourt');
    expect(m.custom).toContainEqual({ name: 'ClientID', value: 'ACME-42' });

    // Slide tag — the pptx equivalent of a docVar
    expect(m.slideTags.some((t) => t.name === 'REVIEW_STATE' && t.value === 'confidential-draft')).toBe(true);

    // Comment authors — legacy (commentAuthors.xml) + modern (authors.xml); initials pinned too
    expect(m.commentAuthors.find((a) => a.name === 'Alice Smith')?.initials).toBe('AS');
    expect(m.commentAuthors.some((a) => a.name === 'Bob Jones')).toBe(true);

    // Comments — legacy p:cm (author resolved by index, date from `dt`) + modern p188:cm
    // (author resolved by GUID, date from `created`). The legacy comment's part is referenced
    // from slide1's rels → anchored to that slide; the modern comment's part isn't → unanchored.
    expect(m.comments.some((c) => c.author === 'Alice Smith' && c.text.includes('revenue figure') && c.date === '2026-05-15T09:00:00Z' && c.slide === 'slide1.xml')).toBe(true);
    expect(m.comments.some((c) => c.author === 'Bob Jones' && c.text.includes('add a source') && c.date === '2026-05-16T11:00:00Z' && c.slide === undefined)).toBe(true);

    // Slides — visible slide with title + resolved speaker notes; hidden slide flagged
    const visible = m.slides.find((s) => s.name === 'slide1.xml');
    expect(visible?.hidden).toBe(false);
    expect(visible?.title).toBe('Quarterly Review');
    expect(visible?.notes).toContain('Q3 shortfall');
    const hidden = m.slides.find((s) => s.name === 'slide2.xml');
    expect(hidden?.hidden).toBe(true);
    expect(hidden?.title).toBe('Internal Only — Do Not Present');
    expect(hidden?.notes).toBe('');

    // External hyperlink surfaces via the shared all-*.rels scan
    expect(m.externalRels.some((r) => r.target.includes('board-portal'))).toBe(true);
  });

  it('returns empty arrays for every list section on a barebones deck (one untitled slide, no tags/comments/notes/custom)', async () => {
    const result = await extractPptxMetadata(await buildMinimalPptx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.custom).toEqual([]);
    expect(result.value.slideTags).toEqual([]);
    expect(result.value.commentAuthors).toEqual([]);
    expect(result.value.comments).toEqual([]);
    expect(result.value.externalRels).toEqual([]);
    expect(result.value.slides).toHaveLength(1);
    expect(result.value.slides[0]?.title).toBe('');
    expect(result.value.slides[0]?.notes).toBe('');
    expect(result.value.slides[0]?.hidden).toBe(false);
  });

  it('returns an api_error Result when the pptx zip is malformed', async () => {
    const result = await extractPptxMetadata(buildMalformedPptx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.message).toContain('ooxml zip parse failed');
    }
  });
});

describe('formatPptxMetadata', () => {
  it('renders a populated `## PPTX metadata` block with every section heading and the authored content', async () => {
    const extracted = await extractPptxMetadata(await buildRichPptx());
    expect(extracted.ok).toBe(true);
    if (!extracted.ok) return;
    const md = formatPptxMetadata(extracted.value);
    expect(md).toContain('## PPTX metadata');
    expect(md).toContain('### Core properties');
    expect(md).toContain('### Custom properties');
    expect(md).toContain('### Slide tags');
    expect(md).toContain('REVIEW_STATE');
    expect(md).toContain('### Comment authors');
    expect(md).toContain('### Comments');
    expect(md).toContain('revenue figure');
    expect(md).toContain('### Slides');
    expect(md).toContain('Quarterly Review');
    expect(md).toContain('Q3 shortfall');
  });

  it('emits `_(none)_` placeholders for empty sections on a barebones deck', async () => {
    const extracted = await extractPptxMetadata(await buildMinimalPptx());
    expect(extracted.ok).toBe(true);
    if (!extracted.ok) return;
    const md = formatPptxMetadata(extracted.value);
    expect(md).toContain('### Slide tags\n\n_(none)_');
    expect(md).toContain('### Comments\n\n_(none)_');
  });
});
