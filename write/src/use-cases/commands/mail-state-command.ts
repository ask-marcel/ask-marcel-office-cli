import type { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { readsOf } from '@ask-marcel/office-core/use-cases/commands/command-graph.ts';
import type { CommandOptionMeta, WriteCommand } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { formatZodError } from '@ask-marcel/office-core/use-cases/commands/format-zod-error.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';

/*
 * What the mail state commands share: the check of their params, the options
 * that name the messages, and the `execute` of a library caller, which is the
 * plan and its writes in one call (the CLI and MCP add the run journal).
 */

/** The params checked with the schema of the command, or the refusal that names each wrong param. */
export const parseParams = <T>(schema: z.ZodType<T>, params: Record<string, string>): Result<T, GraphError> => {
  const parsed = schema.safeParse(params);
  return parsed.success ? ok(parsed.data) : err({ type: 'validation_error', message: formatZodError(parsed.error) });
};

/** The `execute` of a command with a plan: the plan, then its writes; the answer has the ids before and after each change. */
export const createExecute =
  (plan: PlanWrite): WriteCommand['execute'] =>
  async (graph, params, files) => {
    const planned = await plan(readsOf(graph), params, files);
    if (!planned.ok) return planned;
    const applied = await planned.value.apply(graph);
    if (!applied.ok) return applied;
    const changes = applied.value.changes.map(({ action, target, before, after }) => ({ action, target, before, after }));
    return ok({ ...(applied.value.value as object), changes });
  };

export const MESSAGE_ID_OPTION: CommandOptionMeta = {
  name: 'message-id',
  key: 'messageId',
  required: false,
  description:
    'The id of the message, or the ids of more messages separated by commas, as list-mail-folder-messages, search-mail-messages or get-mail-message return them. Give it, --ids-file, or both.',
  argumentHint: { kind: 'idOrName' },
};

export const IDS_FILE_OPTION: CommandOptionMeta = {
  name: 'ids-file',
  key: 'idsFile',
  required: false,
  description:
    'A local text file with one message id on each line (empty lines are passed over), for a batch. A relative path starts from the folder the CLI runs in. At most 1000 messages in one run.',
};

/** The words of the output that each mail state command shares. */
export const BATCH_RESPONSE =
  '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the `messageId`, the `folderId` and the `folder` name of each message before the change, and `after` the ids after it. A message that is no longer there, or that moved or changed between the plan and the write, is in `skipped`; a write that Graph refused is in `failed`, and the other messages still change. With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.';
