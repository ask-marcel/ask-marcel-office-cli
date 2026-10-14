import { describe, expect, it } from 'bun:test';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import type { RunRegistryCommandRequest } from './run-registry-command.ts';
import { describeOutputPathError, runRegistryCommand } from './run-registry-command.ts';

// run-registry-command is core: the commands it names in an error are those of
// the package it runs for, and a shell remedy names that package's bin.
const widgets = fakeCommandPackage();
const deps = { graph: fakeGraphClient(), fs: createFileSystemFake(), commandPackage: widgets };
const listWidgets = (over: Partial<RunRegistryCommandRequest>): RunRegistryCommandRequest => ({
  name: 'list-widgets',
  command: fakeCommand(),
  params: { top: '2' },
  surface: 'cli',
  ...over,
});

describe('a command run for a package', () => {
  it('returns what the command answered', async () => {
    const result = await runRegistryCommand(deps, listWidgets({}));
    expect(result).toEqual({ ok: true, value: { ranWith: { top: '2' } } });
  });

  it('refuses --output-path on a command with no body, naming the bytes commands of the package and a redirect through its bin', async () => {
    const result = await runRegistryCommand(deps, listWidgets({ outputPath: 'out/widgets.json' }));
    expect(result.ok).toBe(false);
    const message = result.ok ? '' : result.error.message;
    expect(message).toContain('Supported: get-widget-file.');
    expect(message).toContain('`widget-cli list-widgets ... > out.json`');
  });

  it('refuses --output-dir on a command with no media, naming the media commands of the package', async () => {
    const result = await runRegistryCommand(deps, listWidgets({ outputDir: 'out/images' }));
    expect(result).toEqual({
      ok: false,
      error: {
        message: '--output-dir: list-widgets did not return a media array — this flag works only with the image-extraction commands. Supported: get-widget-images.',
        code: 'no_media',
      },
    });
  });

  it('words the no-body refusal for an MCP client without a shell redirect', () => {
    const message = describeOutputPathError({ type: 'no_inlined_bytes' }, 'list-widgets', 'mcp', widgets);
    expect(message).toContain('Supported: get-widget-file.');
    expect(message).not.toContain('widget-cli');
  });
});
