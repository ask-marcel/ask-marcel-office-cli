import { describe, expect, it } from 'bun:test';
import { renderToString } from './render-to-string.ts';

describe('a media answer in text output', () => {
  it('keeps the note the answer carried below the image line, as a text answer does', () => {
    const out = renderToString(
      {
        count: 1,
        media: [{ path: 'scan.pdf/pdf/page1/img_p1_1.png', contentType: 'image/png', sizeBytes: 2048, base64: 'AAAA' }],
        note: 'Neither images nor documents, so skipped: notes.txt',
      },
      'text'
    );
    expect(out).toBe(
      '1 image(s), 2 KB total — use --output-dir <dir> to extract them to disk (base64 omitted from text output; add --output json to inline it)\nnote: Neither images nor documents, so skipped: notes.txt\n'
    );
  });

  it('prints the image line alone when there is no note', () => {
    expect(renderToString({ count: 1, media: [{ path: 'a.png', sizeBytes: 1024, base64: 'AAAA' }] }, 'text')).toBe(
      '1 image(s), 1 KB total — use --output-dir <dir> to extract them to disk (base64 omitted from text output; add --output json to inline it)\n'
    );
  });
});
