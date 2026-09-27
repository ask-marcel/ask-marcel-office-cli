import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import { persistMediaIfRequested } from './output-path.ts';

const b64 = (text: string): string => Buffer.from(text).toString('base64');

describe('--output-dir on a media answer', () => {
  it('keeps the note and the skipped list the answer carried, beside the saved files', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/out', {
      count: 1,
      media: [{ path: 'scan.pdf/pdf/page1/img_p1_1.png', contentType: 'image/png', base64: b64('page one') }],
      skipped: ['notes.txt'],
      note: 'Neither images nor documents, so skipped: notes.txt',
    });
    if (!result.ok) throw new Error(result.error.type);
    expect(result.value).toEqual({
      count: 1,
      media: [{ path: 'scan.pdf/pdf/page1/img_p1_1.png', contentType: 'image/png', savedTo: '/out/scan.pdf_pdf_page1_img_p1_1.png' }],
      skipped: ['notes.txt'],
      note: 'Neither images nor documents, so skipped: notes.txt',
    });
  });

  it('never lets two images with one name overwrite each other: the second is saved beside the first', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/out', {
      count: 3,
      media: [
        { path: 'image.png', base64: b64('first') },
        { path: 'image.png', base64: b64('second') },
        { path: 'image.png', base64: b64('third') },
      ],
    });
    if (!result.ok) throw new Error(result.error.type);
    const saved = (result.value as { media: ReadonlyArray<{ savedTo: string }> }).media.map((m) => m.savedTo);
    expect(saved).toEqual(['/out/image.png', '/out/image-2.png', '/out/image-3.png']);
    expect(new TextDecoder().decode(fs.snapshotBytes('/out/image-2.png'))).toBe('second');
    expect(new TextDecoder().decode(fs.snapshotBytes('/out/image.png'))).toBe('first');
  });

  it('numbers a repeated name that starts with a dot after the whole name', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/out', {
      count: 2,
      media: [
        { path: '.thumb', base64: b64('a') },
        { path: '.thumb', base64: b64('b') },
      ],
    });
    if (!result.ok) throw new Error(result.error.type);
    expect((result.value as { media: ReadonlyArray<{ savedTo: string }> }).media.map((m) => m.savedTo)).toEqual(['/out/.thumb', '/out/.thumb-2']);
  });

  it('numbers a repeated name that has no extension too', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/out', {
      count: 2,
      media: [
        { path: 'Im0', base64: b64('a') },
        { path: 'Im0', base64: b64('b') },
      ],
    });
    if (!result.ok) throw new Error(result.error.type);
    expect((result.value as { media: ReadonlyArray<{ savedTo: string }> }).media.map((m) => m.savedTo)).toEqual(['/out/Im0', '/out/Im0-2']);
  });
});
