/*
 * The CLI of the write package: the core CLI builder (`buildCliShell`) given
 * the write package (`write-package.ts`). It runs on the write graph alone. It
 * has no sign-in command (the auth package signs in), and `mcp` is not a
 * commander command: the bin serves it before commander runs.
 */
import type { Command } from 'commander';
import type { CliShellDeps } from '@ask-marcel/office-core/composition/cli-shell.ts';
import { buildCliShell } from '@ask-marcel/office-core/composition/cli-shell.ts';
import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import { commandPackage } from './write-package.ts';

type BuildWriteCliDeps = Omit<CliShellDeps<WriteGraph>, 'commandPackage' | 'addLifecycleCommands'>;

const buildWriteCli = (deps: BuildWriteCliDeps): Command => buildCliShell({ ...deps, commandPackage });

export { buildWriteCli };
export type { BuildWriteCliDeps };
