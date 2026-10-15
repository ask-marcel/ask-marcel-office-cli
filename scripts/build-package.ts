#!/usr/bin/env bun
/*
 * Builds the bundles of one workspace package (auth/, read/, write/). Run it
 * with the folder as the working directory, after `rm -rf dist`:
 *
 *   bun run ../scripts/build-package.ts
 *
 * The package.json decides everything, so the manifest and the bundles cannot
 * drift (docs/plans/2026-10-01-package-split.md, Target layout):
 *
 *   - the bundles are the files that `bin` and the `exports` import targets
 *     name: dist/cli.js from src/main.ts, dist/index.js from src/index.ts,
 *     dist/token.js from src/token.ts. A target with no known entry fails.
 *   - the modules left `external` are exactly the keys of `dependencies`. Every
 *     other import, core first of all, is inlined. A module that the bundle
 *     imports but `dependencies` does not name is inlined too, so a missing
 *     runtime dependency can only make a bundle bigger, never crash an install.
 *
 * The declarations come after, from scripts/build-types.ts.
 */

type Manifest = {
  readonly bin?: Readonly<Record<string, string>>;
  readonly exports?: Readonly<Record<string, string | { readonly import?: string }>>;
  readonly dependencies?: Readonly<Record<string, string>>;
};

const ENTRIES: Readonly<Record<string, string>> = {
  'dist/cli.js': 'src/main.ts',
  'dist/index.js': 'src/index.ts',
  'dist/token.js': 'src/token.ts',
};

const withoutDot = (path: string): string => (path.startsWith('./') ? path.slice(2) : path);

const bundleTargets = (manifest: Manifest): ReadonlyArray<string> => {
  const fromBin = Object.values(manifest.bin ?? {});
  const fromExports = Object.values(manifest.exports ?? {}).map((target) => (typeof target === 'string' ? target : (target.import ?? '')));
  return [...new Set([...fromBin, ...fromExports].map(withoutDot).filter((path) => path.endsWith('.js')))];
};

const build = async (target: string, external: ReadonlyArray<string>): Promise<boolean> => {
  const entry = ENTRIES[target];
  if (entry === undefined) {
    process.stderr.write(`build-package: package.json names ${target}, and no source entry is known for it (scripts/build-package.ts, ENTRIES)\n`);
    return false;
  }
  const result = await Bun.build({ entrypoints: [entry], outdir: 'dist', naming: target.slice('dist/'.length), target: 'node', format: 'esm', external: [...external] });
  for (const log of result.logs) process.stderr.write(`build-package: ${String(log.message)}\n`);
  if (result.success) process.stderr.write(`build-package: ${entry} -> ${target} (external: ${external.join(', ') || 'none'})\n`);
  return result.success;
};

const manifest = (await Bun.file('package.json').json()) as Manifest;
const external = Object.keys(manifest.dependencies ?? {});
const built = await Promise.all(bundleTargets(manifest).map((target) => build(target, external)));
process.exit(built.every(Boolean) ? 0 : 1);
