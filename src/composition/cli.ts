/*
 * The single package's CLI: the core CLI builder (`cli-shell.ts`) given the
 * single package (`single-package.ts`). It has no sign-in commands: `login`,
 * `logout` and `status` are commands of the auth package (`auth-program.ts`).
 * Its deps are the ones `main.ts` and the real-registry suites pass.
 */
import type { Command } from 'commander';
import type { ReadGraph } from '@ask-marcel/office-core/infra/read-graph.ts';
import type { CliShellDeps } from '@ask-marcel/office-core/composition/cli-shell.ts';
import { buildCliShell } from '@ask-marcel/office-core/composition/cli-shell.ts';
import { commandPackage } from './single-package.ts';

type BuildCliDeps = Omit<CliShellDeps<ReadGraph>, 'commandPackage' | 'addLifecycleCommands'>;

const buildCli = (deps: BuildCliDeps): Command => buildCliShell({ ...deps, commandPackage });

export { buildCli };
export type { BuildCliDeps };
