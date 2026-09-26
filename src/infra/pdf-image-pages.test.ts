import { describe, expect, it } from 'bun:test';
import { buildScannedPdf } from '../test-helpers/office-fixtures.ts';
import { extractPdfImages } from './pdf-image-extractor.ts';

describe('the page images of a scanned PDF', () => {
  it('walks every page by default, one image each', async () => {
    const r = await extractPdfImages(buildScannedPdf(3));
    if (!r.ok) throw new Error(r.error.message);
    expect(r.value.map((part) => part.path.split('/')[1])).toEqual(['page1', 'page2', 'page3']);
  });

  it('walks only the pages the caller picks', async () => {
    const r = await extractPdfImages(buildScannedPdf(3), (page) => page !== 2);
    if (!r.ok) throw new Error(r.error.message);
    expect(r.value.map((part) => part.path.split('/')[1])).toEqual(['page1', 'page3']);
  });
});
