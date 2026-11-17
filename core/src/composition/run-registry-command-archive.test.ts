import { describe, expect, it } from 'bun:test';
import { ok } from '../domain/result.ts';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import type { Command } from '../use-cases/commands/command-types.ts';
import type { RunRegistryCommandRequest } from './run-registry-command.ts';
import { runRegistryCommand } from './run-registry-command.ts';

/*
 * A command that answers an archive result (one entry for each file of a zip,
 * with the images of the entry inline) and the two output flags: QA 2026-11-16, P1-1.
 */

const PNG_B64 = btoa(String.fromCharCode(0x89, 0x50));

const answering = (value: unknown): Command => ({ ...fakeCommand(), execute: async () => ok(value) });

const archiveOf = (base64: string): unknown => ({
  count: 1,
  files: [{ path: 'report.docx', text: '# Report', images: [{ path: 'word/media/image1.png', base64 }] }],
});

const run = async (
  value: unknown,
  over: Partial<RunRegistryCommandRequest>
): Promise<{ result: Awaited<ReturnType<typeof runRegistryCommand>>; fs: ReturnType<typeof createFileSystemFake> }> => {
  const fs = createFileSystemFake();
  const deps = { graph: fakeGraphClient(), fs, commandPackage: fakeCommandPackage() };
  const result = await runRegistryCommand(deps, { name: 'unpack-widgets', command: answering(value), params: {}, surface: 'cli', ...over });
  return { result, fs };
};

describe('a command that answers an archive result', () => {
  it('writes the entry images to --output-dir and the markdown to --output-path when the call has the two flags', async () => {
    const { result, fs } = await run(archiveOf(PNG_B64), { outputDir: '/work/media', outputPath: '/work/out.md' });
    expect(result).toEqual(
      ok({
        count: 1,
        files: [{ path: 'report.docx', images: [{ path: 'word/media/image1.png', savedTo: '/work/media/report.docx_word_media_image1.png' }] }],
        savedTo: '/work/out.md',
      })
    );
    expect(Array.from(fs.snapshotBytes('/work/media/report.docx_word_media_image1.png') ?? [])).toEqual([0x89, 0x50]);
    expect(fs.snapshot('/work/out.md')).toBe('# File: report.docx\n\n# Report\n');
  });

  it('refuses more than 1 MB of inline entry images with a message that names --output-dir', async () => {
    const { result } = await run(archiveOf('A'.repeat(1_200_000)), {});
    expect(result).toEqual({
      ok: false,
      error: {
        message:
          '--output-dir: unpack-widgets returned ~0.9 MB of inline images — too large to print to stdout (it would flood the context). Re-run with `--output-dir <dir>` to write the images to disk; each image then carries `savedTo` instead of `base64`.',
        code: 'inline_too_large',
      },
    });
  });
});

describe('a command that answers a media array', () => {
  it('writes the media to --output-dir and does not use --output-path when the call has the two flags', async () => {
    const media = { count: 1, media: [{ path: 'ppt/media/image1.png', base64: PNG_B64 }] };
    const { result, fs } = await run(media, { outputDir: '/work/media', outputPath: '/work/out.md' });
    expect(result).toEqual(ok({ count: 1, media: [{ path: 'ppt/media/image1.png', savedTo: '/work/media/ppt_media_image1.png' }] }));
    expect(fs.has('/work/out.md')).toBe(false);
  });
});
