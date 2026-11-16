import { describe, expect, it } from 'bun:test';
import { join } from 'node:path';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createModuleLoaderFake } from '../test-helpers/module-loader-fake.ts';
import { createPackageLocatorFake } from '../test-helpers/package-locator-fake.ts';
import { createStartPackageBin } from '../use-cases/start-package-bin.ts';
import { runBin } from './run-bin.ts';

/*
 * What a person sees when a bin of the facade runs (ADR 0005): nothing from
 * the facade when the bin of the package starts, and one stderr line with the
 * package, the cause and the next step when it cannot start.
 */

const READ = '@ask-marcel/office-read';
const MANIFEST = join('/global', 'node_modules', '@ask-marcel', 'office-read', 'package.json');
const READ_MANIFEST = JSON.stringify({ name: READ, bin: { 'ask-marcel-office-read': './dist/cli.js' } });

type Install = { readonly installed?: boolean; readonly manifest?: string; readonly loadFailure?: string };
type Seen = { readonly lines: ReadonlyArray<string>; readonly failures: number };

const runRead = async (install: Install): Promise<Seen> => {
  const fs = createFileSystemFake();
  if (install.manifest !== undefined) fs.seed(MANIFEST, install.manifest);
  const locator = createPackageLocatorFake(install.installed === false ? {} : { [READ]: MANIFEST });
  const start = createStartPackageBin({ locator, fs, loader: createModuleLoaderFake(install.loadFailure) });
  const lines: string[] = [];
  const seen = { failures: 0 };
  await runBin({ start, writeError: (line) => lines.push(line), fail: () => (seen.failures += 1) }, READ);
  return { lines, failures: seen.failures };
};

describe('a bin of the facade', () => {
  it('writes nothing and does not fail when the bin of the package starts', async () => {
    expect(await runRead({ manifest: READ_MANIFEST })).toEqual({ lines: [], failures: 0 });
  });

  it.each([
    ['the package is not installed', { installed: false }, `ask-marcel-office-cli: the package ${READ} is not installed. Install ask-marcel-office-cli again.\n`],
    ['its package.json cannot be read', {}, `ask-marcel-office-cli: the package.json of ${READ} cannot be read. Install ask-marcel-office-cli again.\n`],
    [
      'its package.json does not name one bin',
      { manifest: '{}' },
      `ask-marcel-office-cli: the package.json of ${READ} does not name exactly one bin. Install ask-marcel-office-cli again.\n`,
    ],
    ['its bin cannot be loaded', { manifest: READ_MANIFEST, loadFailure: 'Cannot find module' }, `ask-marcel-office-cli: the bin of ${READ} did not start: Cannot find module\n`],
  ])('writes one stderr line with the package, the cause and the next step, and fails, when %s', async (_label, install, line) => {
    expect(await runRead(install)).toEqual({ lines: [line], failures: 1 });
  });
});
