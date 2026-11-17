import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import type { FileSystem } from '../ports/filesystem.ts';
import { persistIfRequested, persistMediaIfRequested } from './output-path.ts';

/*
 * An archive result is what a zip conversion answers: one entry for each file
 * of the archive, `{ count, files: [{ path, text } | { path, note }] }`, and
 * with `--include-images true` each entry can carry `images: [{ path, base64 }]`.
 * The QA of 2026-11-16 (P1-1) found that no output flag knew this shape.
 */

// base64 of the bytes [0x89,0x50] and [0xff,0xd8]
const PNG_B64 = btoa(String.fromCharCode(0x89, 0x50));
const JPG_B64 = btoa(String.fromCharCode(0xff, 0xd8));

const image = (base64: string): { path: string; contentType: string; sizeBytes: number; base64: string } => ({
  path: 'word/media/image1.png',
  contentType: 'image/png',
  sizeBytes: 2,
  base64,
});

const archive = {
  count: 3,
  files: [
    { path: 'docs/report.docx', contentType: 'text/markdown', size: 8, text: '# Report', images: [image(PNG_B64)] },
    { path: 'plan.docx', contentType: 'text/markdown', size: 6, text: '# Plan', images: [image(JPG_B64)] },
    { path: 'photo.png', note: 'png is an image' },
  ],
};

const bigArchive = { count: 1, files: [{ path: 'report.docx', text: '# Report', images: [image('A'.repeat(600_000)), image('B'.repeat(600_000))] }] };

describe('an archive result with no output flag', () => {
  it('is refused when the images of its entries total more than 1 MB of base64, so they do not flood stdout', async () => {
    const result = await persistIfRequested(createFileSystemFake(), undefined, bigArchive);
    expect(result).toEqual(err({ type: 'inline_too_large', base64Length: 1_200_000, inImages: true }));
  });

  it('is printed unchanged when the images of its entries stay below the limit', async () => {
    const result = await persistIfRequested(createFileSystemFake(), undefined, archive);
    expect(result).toEqual(ok(archive));
  });

  it('is printed when the images of its entries total exactly 1 MB of base64: the limit is more than 1 MB', async () => {
    const atLimit = { count: 1, files: [{ path: 'report.docx', text: '# Report', images: [image('A'.repeat(400_000)), image('B'.repeat(600_000))] }] };
    const result = await persistIfRequested(createFileSystemFake(), undefined, atLimit);
    expect(result).toEqual(ok(atLimit));
  });
});

describe('a result that is not an archive result', () => {
  it.each([
    ['an entry of its files is not an object', { count: 1, files: [null] }],
    ['an entry has no path', { count: 1, files: [{ text: '# Report' }] }],
    ['an entry has no text and no note', { count: 1, files: [{ path: 'report.docx' }] }],
    ['it has no count', { files: [{ path: 'report.docx', text: '# Report' }] }],
  ])('gets the refusal of a plain result from the two output flags when %s', async (_why, data) => {
    const fs = createFileSystemFake();
    expect(await persistIfRequested(fs, '/work/out.md', data)).toEqual(err({ type: 'no_inlined_bytes' }));
    expect(await persistMediaIfRequested(fs, '/work/media', data)).toEqual(err({ type: 'no_media' }));
    expect(fs.has('/work/out.md')).toBe(false);
  });
});

describe('an archive result with --output-dir', () => {
  it('writes each entry image under its entry path and its image path, and keeps the text of each entry', async () => {
    const fs = createFileSystemFake();
    const result = await persistMediaIfRequested(fs, '/work/media', archive);
    const saved = (name: string): { path: string; contentType: string; sizeBytes: number; savedTo: string } => ({
      path: 'word/media/image1.png',
      contentType: 'image/png',
      sizeBytes: 2,
      savedTo: `/work/media/${name}`,
    });
    expect(result).toEqual(
      ok({
        count: 3,
        files: [
          { path: 'docs/report.docx', contentType: 'text/markdown', size: 8, text: '# Report', images: [saved('docs_report.docx_word_media_image1.png')] },
          { path: 'plan.docx', contentType: 'text/markdown', size: 6, text: '# Plan', images: [saved('plan.docx_word_media_image1.png')] },
          { path: 'photo.png', note: 'png is an image' },
        ],
      })
    );
    expect(Array.from(fs.snapshotBytes('/work/media/docs_report.docx_word_media_image1.png') ?? [])).toEqual([0x89, 0x50]);
    expect(Array.from(fs.snapshotBytes('/work/media/plan.docx_word_media_image1.png') ?? [])).toEqual([0xff, 0xd8]);
  });

  it('keeps the image of an entry inside the directory when the entry path has `..` and backslashes', async () => {
    const fs = createFileSystemFake();
    const hostile = { count: 1, files: [{ path: '..\\..\\x.docx', text: '# Report', images: [image(PNG_B64)] }] };
    const result = await persistMediaIfRequested(fs, '/work/media', hostile);
    const savedTo = '/work/media/.._.._x.docx_word_media_image1.png';
    expect(result).toEqual(
      ok({ count: 1, files: [{ path: '..\\..\\x.docx', text: '# Report', images: [{ path: 'word/media/image1.png', contentType: 'image/png', sizeBytes: 2, savedTo }] }] })
    );
    expect(Array.from(fs.snapshotBytes(savedTo) ?? [])).toEqual([0x89, 0x50]);
  });

  it('leaves an entry image that is already on disk as it is, and gives the next inline image its own file', async () => {
    const fs = createFileSystemFake();
    const onDisk = { path: 'word/media/image0.png', savedTo: '/old/image0.png' };
    const partly = { count: 1, files: [{ path: 'report.docx', text: '# Report', images: [onDisk, image(PNG_B64)] }] };
    const result = await persistMediaIfRequested(fs, '/work/media', partly);
    const saved = { path: 'word/media/image1.png', contentType: 'image/png', sizeBytes: 2, savedTo: '/work/media/report.docx_word_media_image1.png' };
    expect(result).toEqual(ok({ count: 1, files: [{ path: 'report.docx', text: '# Report', images: [onDisk, saved] }] }));
    expect(fs.has('/old/image0.png')).toBe(false);
  });

  it('writes no file and says why when no entry has an image', async () => {
    const fs = createFileSystemFake();
    const noImages = { count: 1, files: [{ path: 'notes.txt', text: 'hello' }] };
    const result = await persistMediaIfRequested(fs, '/work/media', noImages);
    expect(result).toEqual(
      ok({ ...noImages, note: '--output-dir wrote no image: no entry of the archive has an image. Pass `--include-images true` to extract the images of the entries.' })
    );
  });

  it('reports a write failure of an entry image as write_failed', async () => {
    const failingFs: FileSystem = { ...createFileSystemFake(), writeBytes: async () => err({ type: 'io_failed', message: 'EACCES: permission denied' }) };
    const result = await persistMediaIfRequested(failingFs, '/root/media', archive);
    expect(result).toEqual(err({ type: 'write_failed', message: 'EACCES: permission denied' }));
  });
});

describe('an archive result with --output-path', () => {
  it('writes one markdown file with a heading for each entry, and replaces the text of the entries with savedTo', async () => {
    const fs = createFileSystemFake();
    const noImages = {
      count: 2,
      files: [
        { path: 'docs/report.docx', contentType: 'text/markdown', size: 8, text: '# Report' },
        { path: 'photo.png', note: 'png is an image' },
      ],
    };
    const result = await persistIfRequested(fs, '/work/out.md', noImages);
    expect(result).toEqual(
      ok({
        count: 2,
        files: [
          { path: 'docs/report.docx', contentType: 'text/markdown', size: 8 },
          { path: 'photo.png', note: 'png is an image' },
        ],
        savedTo: '/work/out.md',
      })
    );
    expect(fs.snapshot('/work/out.md')).toBe('# File: docs/report.docx\n\n# Report\n\n---\n\n# File: photo.png\n\nNot converted: png is an image\n');
  });

  it('is refused when more than 1 MB of entry images stays inline, before a file is written', async () => {
    const fs = createFileSystemFake();
    const result = await persistIfRequested(fs, '/work/out.md', bigArchive);
    expect(result).toEqual(err({ type: 'inline_too_large', base64Length: 1_200_000, inImages: true }));
    expect(fs.has('/work/out.md')).toBe(false);
  });

  it('is refused under a binary file name, because the file would hold markdown', async () => {
    const result = await persistIfRequested(createFileSystemFake(), '/work/out.zip', archive);
    expect(result).toEqual(err({ type: 'text_under_binary_extension', contentType: 'text/markdown', requestedExtension: '.zip' }));
  });
});
