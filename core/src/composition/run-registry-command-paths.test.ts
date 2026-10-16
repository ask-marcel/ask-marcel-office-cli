import { describe, expect, it } from 'bun:test';
import { err, ok } from '../domain/result.ts';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { describeOutputPathError, runRegistryCommand } from './run-registry-command.ts';

// Each --output-path refusal says what was wrong with the path or the answer,
// and what to do instead (core, package split phase 3).

const widgets = fakeCommandPackage();
const refusalFor = (error: Parameters<typeof describeOutputPathError>[0]): string => describeOutputPathError(error, 'get-widget-file', 'cli', widgets);

describe('an --output-path that cannot be written', () => {
  it('names an empty path as a quoting mistake, and a directory as not a file', () => {
    expect(refusalFor({ type: 'empty_path' })).toBe('--output-path: path argument is empty (likely a shell-quoting mistake — pass a real filesystem path)');
    expect(refusalFor({ type: 'is_directory' })).toBe('--output-path: must be a file path, not a directory (paths ending in `/` or `\\` look like a directory).');
  });

  it('refuses text under a binary name, and points at the commands that give the original bytes', () => {
    expect(refusalFor({ type: 'text_under_binary_extension', contentType: 'text/markdown', requestedExtension: '.pdf' })).toBe(
      '--output-path: get-widget-file returned text (`text/markdown`), not the bytes of a `.pdf`; saving it under that name would make a file no reader can open. Save it as `.md` or `.txt`, or fetch the original bytes with `get-mail-attachment` (a mail attachment), `download-drive-item-content` (a drive file) or `download-drive-item-version --format original` (a version).'
    );
  });

  it('refuses a PDF name for source bytes that were passed through unconverted', () => {
    expect(refusalFor({ type: 'passthrough_extension_mismatch', contentType: 'application/rtf', requestedExtension: '.pdf' })).toBe(
      "--output-path: response is passthrough source bytes (contentType: `application/rtf`), NOT a converted PDF. Save with the source extension matching that contentType, not `.pdf` — see the response's `note` field."
    );
  });

  it('gives the size of an inline payload too large for stdout in megabytes', () => {
    expect(refusalFor({ type: 'inline_too_large', base64Length: 4_000_000 })).toBe(
      '--output-path: get-widget-file returned a ~3.0 MB inline payload — too large to print to stdout (it would flood the context). Re-run with `--output-path <file>` to write the bytes to disk; the envelope then carries `savedTo` instead of `base64`.'
    );
  });

  it('names the missing parent of a failed write, and keeps any other write failure as the system said it', () => {
    expect(refusalFor({ type: 'write_failed', message: "ENOENT: no such file or directory, open '/x/y.pdf'" })).toBe(
      '--output-path: parent directory missing or not writable: /x/y.pdf'
    );
    expect(refusalFor({ type: 'write_failed', message: 'EACCES: permission denied' })).toBe('--output-path: write failed: EACCES: permission denied');
  });
});

describe('a command with no body to write', () => {
  it('names the bytes commands of the package in alphabetical order', () => {
    const files = fakeCommandPackage({
      registry: {
        'list-widgets': fakeCommand(),
        'get-widget-file': fakeCommand({ producesBytes: true }),
        'export-widget': fakeCommand({ producesBytes: true }),
      },
    });
    expect(describeOutputPathError({ type: 'no_inlined_bytes' }, 'list-widgets', 'mcp', files)).toBe(
      '--output-path: list-widgets did not return inlined bytes — this flag works only with commands that produce a body to write. Supported: export-widget, get-widget-file. Narrow the request instead (fewer items, a tighter query); there is no body to write.'
    );
  });
});

describe('an --output-dir that cannot be written', () => {
  const images = fakeCommand({ producesMedia: true });
  const run = (outputDir: string, fs = createFileSystemFake()): ReturnType<typeof runRegistryCommand> =>
    runRegistryCommand(
      { graph: fakeGraphClient(), fs, commandPackage: widgets },
      {
        name: 'get-widget-images',
        command: { ...images, execute: async () => ok({ count: 1, media: [{ path: 'a/b.png', base64: 'AQ==' }] }) },
        params: {},
        outputDir,
        surface: 'cli',
      }
    );

  it('names an empty directory as a quoting mistake', async () => {
    expect(await run('')).toEqual(err({ message: '--output-dir: directory argument is empty (likely a shell-quoting mistake — pass a real directory path)', code: 'empty_path' }));
  });

  it('keeps a failed image write as the system said it', async () => {
    const full = { ...createFileSystemFake(), writeBytes: async () => err({ type: 'io_failed' as const, message: 'ENOSPC: no space left on device' }) };
    expect(await run('out', full)).toEqual(err({ message: '--output-dir: write failed: ENOSPC: no space left on device', code: 'write_failed' }));
  });
});
