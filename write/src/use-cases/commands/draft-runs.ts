import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommand } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { UndoCheck, UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { AppliedChange, GraphReads, ItemIds, PlanWrite, WritePlan } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';

/*
 * The dry-run and the undo of the four mail draft commands (write-plan.ts).
 * Each command keeps its own `execute`: the plan checks the params as the
 * command does, before any write, and its writes are that `execute`, then one
 * read of the `changeKey` of the draft, which the undo compares. The undo of a
 * new draft moves it to Deleted Items (DELETE /me/messages/{id}); the undo of
 * an update writes back the fields as they were. Each undo runs only if the
 * message is still a draft and its changeKey is still the one the run left.
 */

/** What the undo of a new draft does, as the docs of the three create commands say it (`CommandMeta.undo`). */
export const DRAFT_CREATE_UNDO =
  'moves the draft to Deleted Items (`DELETE /me/messages/{id}`), only if the message is still a draft and its changeKey did not change after the run';

/** What the undo of `update-mail-draft` does. */
export const DRAFT_UPDATE_UNDO =
  'writes back the fields that the update changed, as they were before the run, only if the message is still a draft and its changeKey did not change after the run';

/** What a draft command writes to, and the ids before the write. */
export type DraftTarget = { readonly target: string; readonly before: ItemIds };

/** How a draft command that makes a new draft gets its dry-run and its undo. */
export type DraftCreateSpec = {
  /** The name of the change in the run, for example `create-draft`. */
  readonly action: string;
  /** Checks the params as `execute` does before its first write, and says what it writes to. */
  readonly target: (params: Record<string, string>) => Result<DraftTarget, GraphError>;
  readonly execute: WriteCommand['execute'];
};

const idSchema = z.object({ id: z.string() });

const changeKeySchema = z.object({ changeKey: z.string() });

// The changeKey of the message after the write. When it cannot be read, an empty
// key: the undo then passes over the draft, because it cannot prove it unchanged.
export const changeKeyAfter = async (reads: GraphReads, messageId: string): Promise<string> => {
  const read = await reads.get(`/me/messages/${messageId}?$select=changeKey`);
  const parsed = changeKeySchema.safeParse(read.ok ? read.value : undefined);
  return parsed.success ? parsed.data.changeKey : '';
};

/** One draft change with its undo step, after the write. */
export const draftChange = (action: string, where: DraftTarget, messageId: string, changeKey: string, undo: AppliedChange['undo']): AppliedChange => ({
  action,
  target: where.target,
  before: where.before,
  after: { messageId, changeKey },
  undo,
});

/** The plan of a draft command that makes a new draft: its undo moves the draft to Deleted Items. */
export const createDraftPlan =
  (spec: DraftCreateSpec): PlanWrite =>
  async (_reads, params) => {
    const where = spec.target(params);
    if (!where.ok) return where;
    return ok({
      changes: [{ action: spec.action, ...where.value }],
      apply: async (graph) => {
        const made = await spec.execute(graph, params);
        if (!made.ok) return made;
        // A 204 answer names no draft: there is nothing the undo could find.
        const draft = idSchema.safeParse(made.value);
        if (!draft.success) return ok({ value: made.value, changes: [] });
        const messageId = draft.data.id;
        const changeKey = await changeKeyAfter(graph, messageId);
        return ok({ value: made.value, changes: [draftChange(spec.action, where.value, messageId, changeKey, { kind: 'delete-draft', data: { messageId, changeKey } })] });
      },
    });
  };

const fieldsOf = (value: unknown, names: ReadonlyArray<string>): Readonly<Record<string, unknown>> => {
  const message: Readonly<Record<string, unknown>> = typeof value === 'object' && value !== null ? (value as Record<string, unknown>) : {};
  return Object.fromEntries(names.filter((name) => Object.hasOwn(message, name)).map((name) => [name, message[name]]));
};

/**
 * The plan of an update of the draft `messageId` that writes the fields
 * `names`: it reads them first, and its undo writes them back as they were.
 * `write` is the update itself.
 */
export const planDraftUpdate = async (
  reads: GraphReads,
  messageId: string,
  names: ReadonlyArray<string>,
  write: (graph: WriteGraph) => Promise<Result<unknown, GraphError>>
): Promise<Result<WritePlan, GraphError>> => {
  const read = await reads.get(`/me/messages/${messageId}?$select=${[...names, 'changeKey'].join(',')}`);
  if (!read.ok) return read;
  const version = changeKeySchema.safeParse(read.value);
  const where: DraftTarget = { target: `/me/messages/${messageId}`, before: { messageId, changeKey: version.success ? version.data.changeKey : '' } };
  const fields = fieldsOf(read.value, names);
  return ok({
    changes: [{ action: 'update-draft', ...where }],
    apply: async (graph) => {
      const made = await write(graph);
      if (!made.ok) return made;
      const changeKey = await changeKeyAfter(graph, messageId);
      return ok({ value: made.value, changes: [draftChange('update-draft', where, messageId, changeKey, { kind: 'restore-draft', data: { messageId, changeKey, fields } })] });
    },
  });
};

// A message id of the journal becomes a Graph path: only the characters of a Graph id pass.
const draftDataSchema = z.object({ messageId: z.string().regex(/^[\w=-]+$/), changeKey: z.string() });

const notADraftId: GraphError = { type: 'validation_error', message: 'The run journal names a message id that is not a Graph message id, so this undo does not run.' };

const versionSchema = z.object({ isDraft: z.boolean(), changeKey: z.string() });

// The draft is `ready` when it is still a draft with the changeKey the run left.
const checkDraft = async (reads: GraphReads, data: Readonly<Record<string, unknown>>): Promise<Result<UndoCheck, GraphError>> => {
  const step = draftDataSchema.safeParse(data);
  if (!step.success) return err(notADraftId);
  const read = await reads.get(`/me/messages/${step.data.messageId}?$select=isDraft,changeKey`);
  if (!read.ok) return read.error.type === 'api_error' && read.error.status === 404 ? ok('missing') : read;
  const version = versionSchema.safeParse(read.value);
  return ok(version.success && version.data.isDraft && version.data.changeKey === step.data.changeKey ? 'ready' : 'changed');
};

const restoreSchema = draftDataSchema.extend({ fields: z.record(z.string(), z.unknown()) });

/** The undo handlers of the draft commands, by kind. */
export const DRAFT_UNDO_HANDLERS: UndoHandlers = {
  // DELETE moves the message to Deleted Items: it is not a permanent delete.
  'delete-draft': {
    check: checkDraft,
    apply: async (graph, data) => {
      const step = draftDataSchema.safeParse(data);
      return step.success ? graph.delete(`/me/messages/${step.data.messageId}`) : err(notADraftId);
    },
  },
  'restore-draft': {
    check: checkDraft,
    apply: async (graph, data) => {
      const step = restoreSchema.safeParse(data);
      return step.success ? graph.patch(`/me/messages/${step.data.messageId}`, step.data.fields) : err(notADraftId);
    },
  },
};
