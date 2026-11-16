#!/usr/bin/env bun
/*
 * The local bundle of the facade (cli/, ADR 0005): one tarball that installs
 * `ask-marcel-office-cli` with the three packages inside it.
 *
 *   bun scripts/facade-bundle.ts <out-dir>           writes the tarball and prints its path
 *   bun scripts/facade-bundle.ts <out-dir> --global  then installs it with `bun add -g`
 *
 * Run it after `bun run build` in auth/, read/, write/ and cli/.
 *
 * Why a bundle: the tarball of `bun pm pack` names the three packages at their
 * versions, and a registry that does not have these versions cannot install
 * it. The bundle holds the three packed packages under its node_modules/
 * (`bundledDependencies`), and its `dependencies` are the registry
 * dependencies of the three. The bins are the ones of the facade, with no
 * change. scripts/check-packed-facade.ts installs the bundle in a clean
 * consumer and runs each bin.
 */

import { copyFileSync, existsSync, mkdirSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';

export type Manifest = {
  readonly name: string;
  readonly version: string;
  readonly bin?: Readonly<Record<string, string>>;
  readonly dependencies?: Readonly<Record<string, string>>;
  readonly devDependencies?: Readonly<Record<string, string>>;
  readonly [field: string]: unknown;
};

export const ROOT = resolve(import.meta.dir, '..');
export const FACADE = 'cli';
// The folders of the packages that the facade starts.
export const PACKAGES: ReadonlyArray<string> = ['auth', 'read', 'write'];

export const readManifest = async (folder: string): Promise<Manifest> => (await Bun.file(join(ROOT, folder, 'package.json')).json()) as Manifest;

// The registry dependencies of the packages, as one map. Two packages that
// name one dependency with two ranges have no one answer.
const mergedDependencies = (packages: ReadonlyArray<Manifest>): Readonly<Record<string, string>> => {
  const merged: Record<string, string> = {};
  for (const [name, range] of packages.flatMap((manifest) => Object.entries(manifest.dependencies ?? {}))) {
    if (merged[name] !== undefined && merged[name] !== range) throw new Error(`facade-bundle: ${name} has two ranges, ${merged[name]} and ${range}`);
    merged[name] = range;
  }
  return merged;
};

// The package.json of the bundle: the facade's, with the three packages inside
// it and their registry dependencies as its own.
export const bundleManifest = (facade: Manifest, packages: ReadonlyArray<Manifest>): Manifest => {
  const { devDependencies: _dev, ...kept } = facade;
  return { ...kept, dependencies: mergedDependencies(packages), bundledDependencies: packages.map((manifest) => manifest.name) };
};

const run = (cmd: ReadonlyArray<string>, cwd: string): string => {
  const proc = Bun.spawnSync([...cmd], { cwd, stdout: 'pipe', stderr: 'pipe' });
  if (proc.exitCode !== 0) throw new Error(`facade-bundle: \`${cmd.join(' ')}\` failed in ${cwd}: ${proc.stderr.toString().slice(0, 400)}`);
  return proc.stdout.toString();
};

// Packs one workspace folder into `dir` and gives the file name of its tarball.
// --quiet prints only the path (.claude/LESSONS.md, 2026-10-18).
export const packFolder = (folder: string, dir: string): string => {
  if (!existsSync(join(ROOT, folder, 'dist'))) throw new Error(`facade-bundle: ${folder}/dist is absent: run \`bun run build\` in ${folder}/`);
  const out = run(['bun', 'pm', 'pack', '--destination', dir, '--quiet', '--ignore-scripts'], join(ROOT, folder));
  return basename(out.trim().split('\n').at(-1) ?? '');
};

/*
 * Builds the bundle in a temporary folder and copies it to `outDir`. Each tar
 * call names its files relative to that folder: a GNU tar reads `C:` in an
 * absolute Windows path as a host name.
 */
export const buildBundle = async (outDir: string): Promise<string> => {
  const work = mkdtempSync(join(tmpdir(), 'facade-bundle-'));
  try {
    run(['tar', '-xzf', packFolder(FACADE, work)], work);
    const packages = await Promise.all(PACKAGES.map(readManifest));
    for (const [index, manifest] of packages.entries()) {
      const inside = `package/node_modules/${manifest.name}`;
      mkdirSync(join(work, inside), { recursive: true });
      run(['tar', '-xzf', packFolder(PACKAGES[index] ?? '', work), '-C', inside, '--strip-components=1'], work);
    }
    const facade = (await Bun.file(join(work, 'package', 'package.json')).json()) as Manifest;
    await Bun.write(join(work, 'package', 'package.json'), `${JSON.stringify(bundleManifest(facade, packages), null, 2)}\n`);
    const name = `${facade.name}-${facade.version}-local.tgz`;
    run(['tar', '-czf', name, 'package'], work);
    mkdirSync(outDir, { recursive: true });
    copyFileSync(join(work, name), join(outDir, name));
    return join(outDir, name);
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
};

const main = async (): Promise<number> => {
  const [outDir, ...flags] = process.argv.slice(2);
  if (outDir === undefined || outDir.startsWith('--')) {
    console.error('facade-bundle: name the folder for the tarball, for example `bun scripts/facade-bundle.ts dist-local`');
    return 1;
  }
  const tarball = await buildBundle(resolve(outDir));
  console.log(tarball);
  if (!flags.includes('--global')) return 0;
  const installed = Bun.spawnSync(['bun', 'add', '-g', tarball], { stdout: 'inherit', stderr: 'inherit' });
  return installed.exitCode;
};

if (import.meta.main) process.exit(await main());
