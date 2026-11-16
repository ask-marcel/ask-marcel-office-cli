import { describe, expect, it } from 'bun:test';
import { join } from 'node:path';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { createModuleLoaderFake } from '../test-helpers/module-loader-fake.ts';
import { createPackageLocatorFake } from '../test-helpers/package-locator-fake.ts';
import type { StartError } from '../domain/start-error.ts';
import type { Started } from './start-package-bin.ts';
import { createStartPackageBin } from './start-package-bin.ts';

/*
 * A bin of the facade starts the bin of one package (ADR 0005): it finds the
 * package.json of the installed package, reads its one `bin` entry and runs
 * that file.
 */

const AUTH = '@ask-marcel/office-auth';
// The separator is the one of the system that runs the bin: a backslash on Windows.
const FOLDER = join('/global', 'node_modules', '@ask-marcel', 'office-auth');
const MANIFEST = join(FOLDER, 'package.json');
const BIN = join(FOLDER, 'dist', 'cli.js');
const AUTH_MANIFEST = JSON.stringify({ name: AUTH, bin: { 'ask-marcel-office-auth': './dist/cli.js' } });

type Install = { readonly installed?: boolean; readonly manifest?: string; readonly loadFailure?: string };
type Outcome = { readonly result: Result<Started, StartError>; readonly loaded: ReadonlyArray<string> };

// Starts the auth bin on a computer where the package is installed (or not),
// its package.json holds `manifest` (or is absent), and a load fails or not.
const startAuth = async (install: Install): Promise<Outcome> => {
  const fs = createFileSystemFake();
  if (install.manifest !== undefined) fs.seed(MANIFEST, install.manifest);
  const loader = createModuleLoaderFake(install.loadFailure);
  const locator = createPackageLocatorFake(install.installed === false ? {} : { [AUTH]: MANIFEST });
  const result = await createStartPackageBin({ locator, fs, loader })(AUTH);
  return { result, loaded: loader.loaded() };
};

describe('starting the bin of an installed package', () => {
  it('starts the bin file that the manifest of the installed package names', async () => {
    const { result, loaded } = await startAuth({ manifest: AUTH_MANIFEST });
    expect(result).toEqual({ ok: true, value: { binFile: BIN } });
    expect(loaded).toEqual([BIN]);
  });

  it('gives package_not_installed, with the package name, when the package is not installed', async () => {
    const { result } = await startAuth({ installed: false });
    expect(result).toEqual({ ok: false, error: { type: 'package_not_installed', packageName: AUTH } });
  });

  it.each([
    ['is absent', undefined],
    ['is not JSON', '{ "name": '],
  ])('gives manifest_unreadable when the package.json of the package %s', async (_label, manifest) => {
    const { result } = await startAuth({ manifest });
    expect(result).toEqual({ ok: false, error: { type: 'manifest_unreadable', packageName: AUTH } });
  });

  it.each([
    ['names no bin', { name: AUTH }],
    ['names two bins', { bin: { one: './one.js', two: './two.js' } }],
    ['names a bin with no file', { bin: { one: 7 } }],
    ['names its bin as text, not as a map', { bin: './dist/cli.js' }],
    ['names an empty map of bins', { bin: {} }],
    ['is not an object', null],
  ])('gives no_single_bin when the manifest %s', async (_label, manifest) => {
    const { result } = await startAuth({ manifest: JSON.stringify(manifest) });
    expect(result).toEqual({ ok: false, error: { type: 'no_single_bin', packageName: AUTH } });
  });

  it('gives bin_failed, with the cause, when the bin file cannot be loaded', async () => {
    const { result, loaded } = await startAuth({ manifest: AUTH_MANIFEST, loadFailure: 'Cannot find module' });
    expect(result).toEqual({ ok: false, error: { type: 'bin_failed', packageName: AUTH, message: 'Cannot find module' } });
    expect(loaded).toEqual([BIN]);
  });

  it.each([
    ['the package is not installed', { installed: false }],
    ['the package.json cannot be read', {}],
    ['the manifest names no bin', { manifest: '{}' }],
  ])('loads no file when %s', async (_label, install) => {
    const { loaded } = await startAuth(install);
    expect(loaded).toEqual([]);
  });
});
