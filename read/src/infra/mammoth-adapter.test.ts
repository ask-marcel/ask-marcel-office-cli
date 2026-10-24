import { describe, expect, it } from 'bun:test';
import { buildMalformedDocx, buildSampleDocx } from '../test-helpers/office-fixtures.ts';
import { mammothToHtml } from './mammoth-adapter.ts';

describe('mammothToHtml', () => {
  it('converts a docx with heading + paragraph + bold + italic + table + inline image into HTML with images embedded as data: URIs', async () => {
    const bytes = await buildSampleDocx();
    const result = await mammothToHtml(bytes);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toContain('<h1>Sample Heading</h1>');
      expect(result.value).toContain('<strong>world</strong>');
      expect(result.value).toContain('<em>italic</em>');
      expect(result.value).toContain('<table>');
      expect(result.value).toContain('data:image/png;base64,');
    }
  });

  it('returns err({ type: api_error }) when the bytes are not a valid docx archive', async () => {
    const result = await mammothToHtml(buildMalformedDocx());
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'api_error') {
      expect(result.error.status).toBe(500);
      expect(result.error.message).toContain('docx conversion failed');
    }
  });
});
