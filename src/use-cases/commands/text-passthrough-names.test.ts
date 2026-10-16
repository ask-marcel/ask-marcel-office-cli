import { describe, expect, it } from 'bun:test';
import { extensionOf, isPdfSource } from './text-passthrough.ts';

// The file-name checks the conversions use before asking Graph to convert
// (core, package split phase 3).

describe('the extension of a file name', () => {
  it('is the lower-case text after the last dot, or nothing when there is no dot', () => {
    expect(extensionOf('Report.Final.PDF')).toBe('pdf');
    expect(extensionOf('Makefile')).toBe('');
    expect(extensionOf('archive.')).toBe('');
  });
});

describe('a PDF source', () => {
  it('is a file whose extension is pdf, in any case', () => {
    expect(isPdfSource('scan.PDF')).toBe(true);
    expect(isPdfSource('scan.pdfx')).toBe(false);
    expect(isPdfSource('pdf')).toBe(false);
  });
});
