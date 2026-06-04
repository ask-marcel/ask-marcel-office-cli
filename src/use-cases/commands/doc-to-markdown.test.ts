import { describe, expect, it } from 'bun:test';
import { buildSampleDoc } from '../../test-helpers/office-fixtures.ts';
import { docToMarkdown } from './doc-to-markdown.ts';

describe('docToMarkdown', () => {
  it('returns the legacy .doc body as a text/plain envelope', async () => {
    const result = await docToMarkdown(await buildSampleDoc());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.contentType).toBe('text/plain');
    expect(result.value.text).toContain('Hello from the legacy doc');
    expect(result.value.size).toBeGreaterThan(0);
  });

  it('propagates the extraction error for bytes that are not a parseable .doc', async () => {
    const result = await docToMarkdown(Uint8Array.from([1, 2, 3]));
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
  });
});
