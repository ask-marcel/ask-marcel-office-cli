import { describe, expect, it } from 'bun:test';
import { Document, ImageRun, Packer, Paragraph, TextRun } from 'docx';
import JSZip from 'jszip';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import type { FileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { buildCli } from './cli.ts';

/*
 * QA 2026-11-16, P1-1: `convert-local-file-to-markdown --include-images true`
 * on a zip printed megabytes of base64, and no output flag could take them.
 * The archive here is synthetic: a document with one 800 KB image, and a note.
 */

const IMAGE_BYTES = 800_000; // 1 066 668 characters of base64: above the 1 MB inline limit

const bigImage = (): Uint8Array => Uint8Array.from({ length: IMAGE_BYTES }, (_, i) => (i * 31) % 251);

const buildArchive = async (): Promise<Uint8Array> => {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ children: [new TextRun('Survey of the site')] }),
          new Paragraph({ children: [new ImageRun({ type: 'png', data: bigImage(), transformation: { width: 100, height: 100 } })] }),
        ],
      },
    ],
  });
  const zip = new JSZip();
  zip.file('report.docx', new Uint8Array(await Packer.toBuffer(doc)));
  zip.file('notes.txt', new TextEncoder().encode('hello from the archive'));
  return zip.generateAsync({ type: 'uint8array' });
};

const captureStdout = async (run: () => Promise<unknown>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run();
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

const convert = async (flags: ReadonlyArray<string>): Promise<{ out: string; fs: FileSystemFake }> => {
  const fs = createFileSystemFake();
  fs.seedBytes('/work/handover.zip', await buildArchive());
  const cli = buildCli({ graph: fakeGraphClient(), logger: createLoggerFake(), fs });
  const out = await captureStdout(() =>
    cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'convert-local-file-to-markdown', '--path', '/work/handover.zip', ...flags])
  );
  return { out, fs };
};

type SavedImage = { readonly path: string; readonly savedTo?: string; readonly base64?: string };
type Converted = {
  readonly ok: boolean;
  readonly errorCode?: string;
  readonly error?: string;
  readonly data?: { readonly savedTo?: string; readonly files: ReadonlyArray<{ readonly path: string; readonly text?: string; readonly images?: ReadonlyArray<SavedImage> }> };
};

describe('convert-local-file-to-markdown on a zip that holds a document with a large image', () => {
  it('refuses to print the images with --include-images true and no output flag, and names --output-dir', async () => {
    const { out } = await convert(['--include-images', 'true']);
    const parsed = JSON.parse(out.trim()) as Converted;
    expect(parsed.ok).toBe(false);
    expect(parsed.errorCode).toBe('inline_too_large');
    expect(parsed.error).toContain('--output-dir <dir>');
    expect(out.length).toBeLessThan(2_000);
  });

  it('writes the image to --output-dir, and the result names the file and keeps the text', async () => {
    const { out, fs } = await convert(['--include-images', 'true', '--output-dir', '/work/media']);
    const parsed = JSON.parse(out.trim()) as Converted;
    const report = parsed.data?.files.find((f) => f.path === 'report.docx');
    const saved = report?.images?.[0];
    expect(saved?.savedTo).toStartWith('/work/media/report.docx_word_media_');
    expect(saved?.base64).toBeUndefined();
    expect(fs.snapshotBytes(saved?.savedTo ?? '')).toEqual(bigImage());
    expect(report?.text).toContain('Survey of the site');
  });

  it('writes the markdown of all the entries to --output-path', async () => {
    const { out, fs } = await convert(['--output-path', '/work/out.md']);
    const parsed = JSON.parse(out.trim()) as Converted;
    expect(parsed.data?.savedTo).toBe('/work/out.md');
    expect(parsed.data?.files.map((f) => f.text)).toEqual([undefined, undefined]);
    const written = fs.snapshot('/work/out.md') ?? '';
    expect(written).toContain('# File: report.docx\n\nSurvey of the site');
    expect(written).toContain('# File: notes.txt\n\nhello from the archive');
  });
});
