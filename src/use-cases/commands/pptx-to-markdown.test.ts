import { describe, expect, it } from 'bun:test';
import { buildEmptyPptx, buildMalformedPptx, buildRichPptx } from '../../test-helpers/office-fixtures.ts';
import { pptxToMarkdown } from './pptx-to-markdown.ts';

describe('pptxToMarkdown', () => {
  it('renders each slide as a `## Slide N` section — body text then speaker notes inline, hidden slides flagged, sections blank-line separated', async () => {
    const result = await pptxToMarkdown(await buildRichPptx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.contentType).toBe('text/markdown');
    // Exact output pins the section header, the ` (hidden)` marker, the `**Speaker notes:** ` prefix,
    // and the blank-line joins — so a string/conditional/join mutant in slideSection can't survive.
    expect(result.value.text).toBe(
      [
        '## Slide 1',
        '',
        'Quarterly Review\nsee the portal',
        '',
        '**Speaker notes:** Remember to mention the Q3 shortfall and the ACME contract renewal.',
        '',
        '## Slide 2 (hidden)',
        '',
        'Internal Only — Do Not Present',
      ].join('\n')
    );
    expect(result.value.text).not.toContain('## PPTX metadata'); // no side-channel block without the flag
  });

  it('appends the `## PPTX metadata` side-channel block when includeMetadata is true (body still present)', async () => {
    const result = await pptxToMarkdown(await buildRichPptx(), { includeMetadata: true });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.text).toContain('## Slide 1'); // slide body still rendered
    expect(result.value.text).toContain('## PPTX metadata'); // + metadata block
  });

  it('renders a slide with no text or notes as a bare `## Slide N` header (no empty body / notes lines)', async () => {
    const result = await pptxToMarkdown(await buildEmptyPptx());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.text).toBe('## Slide 1'); // text/notes branches both skipped — no trailing blank lines
  });

  it('propagates the zip-parse error for a malformed deck', async () => {
    const result = await pptxToMarkdown(buildMalformedPptx());
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
  });
});
