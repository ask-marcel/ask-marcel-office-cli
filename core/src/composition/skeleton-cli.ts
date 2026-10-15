/*
 * The bin of a package in the workspace skeleton (package split, phase 2).
 *
 * auth, read and write exist as packages before any command moves out of the
 * root src/, so that their build, pack and install checks run from day one.
 * Until a package receives its commands (phases 4 to 6), its bin answers
 * `--version` and refuses everything else. The package's entry gives its own
 * name and version, and the two writers, so this module never touches stdout.
 */

export type SkeletonPackage = { readonly name: string; readonly version: string };

export type SkeletonOutput = {
  readonly out: (text: string) => void;
  readonly err: (text: string) => void;
};

export const runSkeletonCli = (pkg: SkeletonPackage, args: ReadonlyArray<string>, output: SkeletonOutput): number => {
  if (args[0] === '--version') {
    output.out(`${pkg.version}\n`);
    return 0;
  }
  output.err(`${pkg.name} ${pkg.version} has no commands yet: it is a skeleton of the package split, and its commands move in with a later phase.\n`);
  return 1;
};
