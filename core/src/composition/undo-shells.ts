/*
 * The `undo` of a recorded run on both surfaces (undo-run.ts): the CLI command
 * `undo <runId> [--apply]`, and the MCP tools `undo` (it only shows, so it is
 * read-only) and `undo-apply` (it reverses changes, a delete among them, so it
 * carries the destructive mark). A package with a run journal adds them: the
 * write package does, in write-program.ts and write-server.ts.
 */
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import type { Command } from 'commander';
import { z } from 'zod';
import type { WriteGraph } from '../infra/write-graph.ts';
import { renderErrorToString, renderToString } from '../presenter/render-to-string.ts';
import type { RunRecorder } from '../use-cases/commands/record-run.ts';
import type { UndoHandlers } from '../use-cases/commands/undo-run.ts';
import { undoRun } from '../use-cases/commands/undo-run.ts';
import type { CliOutput } from './cli-shell.ts';

/** What the undo needs: the run journal and its clock, the write graph, and the undo handler of each kind. */
export type UndoShellDeps = { readonly runs: RunRecorder; readonly graph: WriteGraph; readonly handlers: UndoHandlers };

const undo = async (deps: UndoShellDeps, runId: string, apply: boolean): ReturnType<typeof undoRun> =>
  undoRun({ journal: deps.runs.journal, graph: deps.graph, handlers: deps.handlers, now: deps.runs.now }, runId, apply);

const RUN_ID_HELP = 'The runId field of the output of the write to reverse, for example 20261118T090507Z-0a1b2c3d. The journal keeps a run 30 days.';

/** Adds `undo <runId> [--apply]` to a CLI. */
export const addUndoCommand = (program: Command, output: CliOutput, deps: UndoShellDeps): void => {
  program
    .command('undo')
    .description(
      'Show the undo of a write run, or do it with --apply. Each change is reversed only if its item did not change after the run; the others are passed over and reported. A run that is fully undone cannot be undone again.'
    )
    .argument('<runId>', RUN_ID_HELP)
    .option('--apply', 'Do the undo. Without it, the command only shows what the undo would do, and changes nothing.')
    .action(async (runId: string, opts: { readonly apply?: true }) => {
      const result = await undo(deps, runId, opts.apply === true);
      if (result.ok) {
        output.print(result.value);
        return;
      }
      output.fail(result.error.message, result.error.code, 'cli');
    });
};

const toolAnswer = (result: Awaited<ReturnType<typeof undoRun>>): CallToolResult =>
  result.ok
    ? { content: [{ type: 'text', text: renderToString(result.value, 'text') }] }
    : { content: [{ type: 'text', text: renderErrorToString(result.error.message, 'text', result.error.code, 'cli') }], isError: true };

/** Registers the MCP tools `undo` (read-only) and `undo-apply` (destructive). */
export const addUndoTools = (server: McpServer, deps: UndoShellDeps): void => {
  const inputSchema = { runId: z.string().describe(RUN_ID_HELP) };
  server.registerTool(
    'undo',
    {
      title: 'Show the undo of a write run',
      description:
        'Show what the undo of a write run would do, change by change, and change nothing. A change whose item changed after the run, or is gone, is passed over. Call undo-apply to do it.',
      inputSchema,
      annotations: { readOnlyHint: true, idempotentHint: true },
    },
    async ({ runId }): Promise<CallToolResult> => toolAnswer(await undo(deps, runId, false))
  );
  server.registerTool(
    'undo-apply',
    {
      title: 'Undo a write run',
      description:
        'Reverse the changes of a write run, for example move a draft that the run made to Deleted Items. Each change is reversed only if its item did not change after the run; the others are passed over and reported. Call undo first to see what it will do.',
      inputSchema,
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ runId }): Promise<CallToolResult> => toolAnswer(await undo(deps, runId, true))
  );
};
