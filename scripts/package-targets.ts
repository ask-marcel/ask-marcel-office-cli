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

const withoutDot = (path: string): string => (path.startsWith('./') ? path.slice(2) : path);

export const bundleTargets = (manifest: TargetManifest, hasSource: (path: string) => boolean): ReadonlyArray<string> => {
  const fromBin = Object.values(manifest.bin ?? {});
  const fromExports = Object.values(manifest.exports ?? {}).map((target) => (typeof target === 'string' ? target : (target['import'] ?? '')));
  const known = Object.entries(ENTRIES)
    .filter(([, source]) => hasSource(source))
    .map(([target]) => target);
  return [...new Set([...[...fromBin, ...fromExports].map(withoutDot).filter((path) => path.endsWith('.js')), ...known])];
};
