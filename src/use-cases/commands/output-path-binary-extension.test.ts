import { describe, expect, it } from 'bun:test';
import { formatOutputPathError } from '../../composition/run-registry-command.ts';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import { persistIfRequested } from './output-path.ts';

describe('saving a text answer under a binary file name', () => {
  it('refuses a PDF text layer saved as .pdf, names the extension and the content type, and writes nothing', async () => {
    const fs = createFileSystemFake();
    const result = await persistIfRequested(fs, '/work/deck.pdf', { contentType: 'text/plain', size: 8, text: 'PDF text' });
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error).toEqual({ type: 'text_under_binary_extension', requestedExtension: '.pdf', contentType: 'text/plain' });
    expect(fs.has('/work/deck.pdf')).toBe(false);
  });

  it('names the content type as text/plain when the envelope carries none', async () => {
    const result = await persistIfRequested(createFileSystemFake(), '/work/deck.pdf', { text: 'PDF text' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toEqual({ type: 'text_under_binary_extension', requestedExtension: '.pdf', contentType: 'text/plain' });
  });

  it.each([
    ['/work/call-log.XLSX', '.xlsx'],
    ['/work/brief.docx', '.docx'],
    ['/work/deck.pptm', '.pptm'],
    ['/work/sheet.ods', '.ods'],
    ['/work/old.doc', '.doc'],
    ['/work/old.xls', '.xls'],
    ['/work/binary.xlsb', '.xlsb'],
    ['/work/old.ppt', '.ppt'],
    ['/work/bundle.zip', '.zip'],
    ['/work/mail.msg', '.msg'],
    ['/work/shot.png', '.png'],
    ['/work/photo.jpg', '.jpg'],
    ['/work/photo.jpeg', '.jpeg'],
    ['/work/loop.gif', '.gif'],
    ['/work/pic.webp', '.webp'],
    ['/work/scan.bmp', '.bmp'],
    ['/work/scan.tif', '.tif'],
    ['/work/scan.tiff', '.tiff'],
  ])('refuses converted markdown saved as %s, naming its content type', async (path, extension) => {
    const result = await persistIfRequested(createFileSystemFake(), path, { contentType: 'text/markdown', size: 3, text: '# x' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toEqual({ type: 'text_under_binary_extension', requestedExtension: extension, contentType: 'text/markdown' });
  });

  it.each(['/work/notes.md', '/work/figures.csv', '/work/raw.eml', '/work/page.html', '/work/log.txt', '/work/no-extension'])('still writes text saved as %s', async (path) => {
    const fs = createFileSystemFake();
    const result = await persistIfRequested(fs, path, { contentType: 'text/plain', size: 5, text: 'hello' });
    expect(result.ok).toBe(true);
    expect(fs.snapshot(path)).toBe('hello');
  });

  it('points at the raw-bytes commands in the CLI message', () => {
    const message = formatOutputPathError({ type: 'text_under_binary_extension', requestedExtension: '.pdf', contentType: 'text/plain' }, 'read-mail-attachment', 'cli');
    expect(message).toBe(
      '--output-path: read-mail-attachment returned text (`text/plain`), not the bytes of a `.pdf`; saving it under that name would make a file no reader can open. Save it as `.md` or `.txt`, or fetch the original bytes with `get-mail-attachment` (a mail attachment), `download-drive-item-content` (a drive file) or `download-drive-item-version --format original` (a version).'
    );
  });
});
