/*
 * Heuristic detection of which package manager originally installed the CLI.
 * Used by the `ask-marcel-office update` command to invoke the matching install
 * tool (`npm i -g …` vs `bun add -g …`).
 *
 * Heuristic: if the bin path contains a `.bun` segment, it is a bun-installed
 * global; otherwise default to npm. This is good enough — the worst case is
 * the user installs via one tool and the update routes through the other,
 * which still produces an updated copy in the other tool's prefix; the user
 * can re-run with the original tool to clean up.
 */

export type PackageManager = 'npm' | 'bun';

export const detectPackageManager = (binPath: string): PackageManager => {
  const normalised = binPath.replaceAll('\\', '/');
  if (normalised.includes('/.bun/') || normalised.endsWith('/.bun')) return 'bun';
  return 'npm';
};
