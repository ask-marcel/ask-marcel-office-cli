import { describe, expect, it } from 'bun:test';
import { buildSampleDoc } from '../test-helpers/office-fixtures.ts';
import { extractDocText } from './legacy-doc-extractor.ts';

describe('extractDocText', () => {
  it('extracts the body text of a legacy .doc (the pre-2007 OLE binary Word format)', async () => {
    const result = await extractDocText(await buildSampleDoc());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toContain('Hello from the legacy doc');
    expect(result.value).toContain('Second paragraph here');
  });

  it('returns an api_error for bytes that are not a parseable .doc', async () => {
    const result = await extractDocText(Uint8Array.from([1, 2, 3, 4, 5]));
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.message : '').toContain('doc text extraction failed');
  });
});
