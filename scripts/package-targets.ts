/*
 * The bundles of one workspace package (auth/, read/, write/), in one place, so
 * that scripts/build-package.ts builds exactly what
 * scripts/check-packed-package.ts expects in the tarball:
 *
 *   - each .js file that `bin` and the `exports` import targets name;
 *   - each known entry below whose source exists in the folder. Only auth has
 *     src/token.ts, so only auth ships dist/token.js, the token helper that the
 *     locator names. It is not a bin (it would put a command on PATH that no
 *     person runs) and not an export (an import would run the helper).
 *
 * The facade (cli/, ADR 0005) has three bins and no library: dist/auth.js,
 * dist/read.js and dist/write.js, each from the source of the same name. They
 * are built only when `bin` names them (BIN_ENTRIES), never from the source
 * alone: a file named src/auth.ts in another folder is not a bundle.
 */

export type TargetManifest = {
  readonly bin?: Readonly<Record<string, string>>;
  readonly exports?: Readonly<Record<string, string | Readonly<Record<string, string>>>>;
};

export const ENTRIES: Readonly<Record<string, string>> = {
  'dist/cli.js': 'src/main.ts',
  'dist/index.js': 'src/index.ts',
  'dist/token.js': 'src/token.ts',
};

// The bins of the facade: a source for a target that `bin` names, and no more.
export const BIN_ENTRIES: Readonly<Record<string, string>> = {
  'dist/auth.js': 'src/auth.ts',
  'dist/read.js': 'src/read.ts',
  'dist/write.js': 'src/write.ts',
};

// The source of a bundle target, or undefined for a target with no known entry.
export const sourceOf = (target: string): string | undefined => ENTRIES[target] ?? BIN_ENTRIES[target];

const withoutDot = (path: string): string => (path.startsWith('./') ? path.slice(2) : path);

export const bundleTargets = (manifest: TargetManifest, hasSource: (path: string) => boolean): ReadonlyArray<string> => {
  const fromBin = Object.values(manifest.bin ?? {});
  const fromExports = Object.values(manifest.exports ?? {}).map((target) => (typeof target === 'string' ? target : (target['import'] ?? '')));
  const known = Object.entries(ENTRIES)
    .filter(([, source]) => hasSource(source))
    .map(([target]) => target);
  return [...new Set([...[...fromBin, ...fromExports].map(withoutDot).filter((path) => path.endsWith('.js')), ...known])];
};
