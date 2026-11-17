/*
 * The CLI of the write package: the core CLI builder (`buildCliShell`) given
 * the write package (`write-package.ts`). It runs on the write graph alone. It
 * has no sign-in command (the auth package signs in), and `mcp` is not a
 * commander command: the bin serves it before commander runs. With the run
 * journal (`runs`, write-runs.ts), each draft command gets a run id and the CLI
 * has `undo <runId> [--apply]`.
 */
import type { Command } from 'commander';
import type { CliShellDeps } from '@ask-marcel/office-core/composition/cli-shell.ts';
import { buildCliShell } from '@ask-marcel/office-core/composition/cli-shell.ts';
import { addUndoCommand } from '@ask-marcel/office-core/composition/undo-shells.ts';
import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import { undoHandlers } from '../use-cases/commands/write-commands.ts';
import { commandPackage } from './write-package.ts';

type BuildWriteCliDeps = Omit<CliShellDeps<WriteGraph>, 'commandPackage' | 'addLifecycleCommands'>;

const buildWriteCli = (deps: BuildWriteCliDeps): Command => {
  const { runs, graph } = deps;
  if (runs === undefined) return buildCliShell({ ...deps, commandPackage });
  return buildCliShell({ ...deps, commandPackage, addLifecycleCommands: (program, output) => addUndoCommand(program, output, { runs, graph, handlers: undoHandlers }) });
};

export { buildWriteCli };
export type { BuildWriteCliDeps };
