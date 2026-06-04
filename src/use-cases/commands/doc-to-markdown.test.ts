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
    // size = byte length of the extracted text, NOT the source .doc (audit B3)
    expect(result.value.size).toBe(new TextEncoder().encode(result.value.text).byteLength);
  });

  it('propagates the extraction error for bytes that are not a parseable .doc', async () => {
    const result = await docToMarkdown(Uint8Array.from([1, 2, 3]));
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
  });
});
