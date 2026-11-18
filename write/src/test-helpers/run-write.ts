import { runRegistryCommand } from '@ask-marcel/office-core/composition/run-registry-command.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { RunRecorder } from '@ask-marcel/office-core/use-cases/commands/record-run.ts';
import { undoRun } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { UndoReport } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { RunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { commands, undoHandlers } from '../use-cases/commands/write-commands.ts';

/**
 * Runs a command of the write registry as the shells run it
 * (`runRegistryCommand`), with the run journal of `journal`, and the undo of
 * its run, for the tests of the commands that have a dry-run and an undo.
 * `value` is the answer of a run, or its error.
 */
export const RUN_ID = '20261118T090507Z-0a1b2c3d';

const NOW = new Date('2026-11-18T09:05:07.000Z');

export const recorderOf = (journal: RunJournalFake): RunRecorder => ({ journal, now: () => NOW, randomSuffix: () => '0a1b2c3d' });

export type RunOptions = { readonly dryRun?: boolean; readonly journal?: RunJournalFake; readonly files?: LocalFileReader };

export const runWrite = async (graph: WriteGraph, name: string, params: Record<string, string>, options: RunOptions = {}): Promise<Record<string, unknown>> => {
  const command = commands[name];
  if (command === undefined) throw new Error(`no command ${name}`);
  const deps = {
    graph,
    fs: createFileSystemFake(),
    commandPackage: { registry: commands, binName: 'ask-marcel-office-write' },
    runs: recorderOf(options.journal ?? createRunJournalFake()),
    ...(options.files === undefined ? {} : { files: options.files }),
  };
  const result = await runRegistryCommand(deps, { name, command, params, surface: 'cli', dryRun: options.dryRun === true });
  return result.ok ? (result.value as Record<string, unknown>) : { error: result.error };
};

export const undoWrite = async (graph: WriteGraph, journal: RunJournalFake, apply: boolean): Promise<Result<UndoReport, unknown>> =>
  undoRun({ journal, graph, handlers: undoHandlers, now: () => NOW }, RUN_ID, apply);

/** The status of each change of an undo, and its reason when it has one. */
export const statusesOf = (report: Result<UndoReport, unknown>): unknown =>
  report.ok ? report.value.changes.map((change) => (change.reason === undefined ? change.status : `${change.status}: ${change.reason}`)) : report.error;
