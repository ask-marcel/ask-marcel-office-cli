import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { AppliedChange, GraphReads, ItemIds, PlannedChange, SkippedItem, UndoStep, WritePlan } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { parseGraphId } from '../../domain/graph-id.ts';
import type { FolderReader } from './mail-folder-reader.ts';
import { isNotThere } from './mail-folder-reader.ts';

/*
 * The plan and the writes of a mail state command over one or more messages
 * (write-plan.ts). The plan reads each message: its folder, its changeKey and
 * the fields the command changes, and shows the current folder of each one, so
 * the user sees a draft that a search found in place of the received message.
 * The writes read each message again, and pass over one that moved, changed or
 * is gone since the plan. Each message is one write (no `$batch`). A write
 * that fails is reported, and the others go on: the changes made stay in the
 * run journal, so the undo can reverse them.
 */

/** A message as the plan read it: `fields` holds the fields the command asked for. */
export type MailMessage = {
  readonly id: GraphId;
  readonly parentFolderId: string;
  readonly changeKey: string;
  readonly fields: Readonly<Record<string, unknown>>;
};

/** What a command does with one message: change it, pass over it with a reason, or refuse the run. */
export type Decision = 'change' | { readonly skip: string } | { readonly refuse: GraphError };

/** The ids after a write, and how to reverse it. */
export type Written = { readonly after: ItemIds; readonly undo: UndoStep };

export type MailBatchSpec = {
  /** The name of the change in the run, for example `move-message`. */
  readonly action: string;
  /** The fields that the command reads and changes, beside the folder and the changeKey. */
  readonly select: ReadonlyArray<string>;
  /** The Graph path that the write of a message goes to. */
  readonly target: (id: GraphId) => string;
  readonly decide: (message: MailMessage) => Decision;
  /** The state of the message before the write, as the output shows it (strings). */
  readonly shown: (message: MailMessage) => ItemIds;
  readonly write: (graph: WriteGraph, message: MailMessage) => Promise<Result<Written, GraphError>>;
};

const GONE = 'The message is no longer there: it moved to another folder (a move gives it a new id), or it was deleted.';
const CHANGED = 'The message moved or changed after the plan, so this run passes over it. Run the command again to see it as it is now.';

const messageSchema = z.object({ parentFolderId: z.string(), changeKey: z.string() }).loose();

const readMessage = async (reads: GraphReads, id: GraphId, spec: MailBatchSpec): Promise<Result<MailMessage, GraphError>> => {
  const read = await reads.get(`/me/messages/${id}?$select=${['parentFolderId', 'changeKey', ...spec.select].join(',')}`);
  if (!read.ok) return read;
  const parsed = messageSchema.safeParse(read.value);
  if (!parsed.success) return err({ type: 'validation_error', message: `Graph gave the message ${id} in a shape that this CLI cannot read.` });
  return ok({ id, parentFolderId: parsed.data.parentFolderId, changeKey: parsed.data.changeKey, fields: parsed.data });
};

// The name of the folder of the message, empty when it cannot be read: the dry-run then shows the folder id only.
const folderName = async (folders: FolderReader, folderId: string): Promise<string> => {
  const id = parseGraphId(folderId, 'mail folder');
  const folder = id.ok ? await folders.folder(id.value) : id;
  return folder.ok ? folder.value.displayName : '';
};

const before = async (folders: FolderReader, message: MailMessage, spec: MailBatchSpec): Promise<ItemIds> => ({
  messageId: message.id,
  folderId: message.parentFolderId,
  folder: await folderName(folders, message.parentFolderId),
  ...spec.shown(message),
});

type Planned = { readonly message: MailMessage; readonly change: PlannedChange };

type Batch = { readonly planned: Planned[]; readonly skipped: SkippedItem[] };

// Reads one message of the plan and puts it in the batch; a refusal or a read error stops the plan.
const planOne = async (reads: GraphReads, folders: FolderReader, id: GraphId, spec: MailBatchSpec, batch: Batch): Promise<Result<true, GraphError>> => {
  const target = spec.target(id);
  const read = await readMessage(reads, id, spec);
  if (!read.ok) return isNotThere(read.error) ? skip(batch, { target, before: { messageId: id }, reason: GONE }) : read;
  const shown = await before(folders, read.value, spec);
  const decision = spec.decide(read.value);
  if (decision !== 'change') return 'refuse' in decision ? err(decision.refuse) : skip(batch, { target, before: shown, reason: decision.skip });
  batch.planned.push({ message: read.value, change: { action: spec.action, target, before: shown } });
  return ok(true);
};

const skip = (batch: Batch, item: SkippedItem): Result<true, GraphError> => {
  batch.skipped.push(item);
  return ok(true);
};

type Outcome = { readonly changes: AppliedChange[]; readonly skipped: SkippedItem[]; readonly failed: SkippedItem[]; firstError?: GraphError };

// Reads the message again, and writes it only if it is as the plan read it.
const applyOne = async (graph: WriteGraph, item: Planned, spec: MailBatchSpec, outcome: Outcome): Promise<void> => {
  const { change, message } = item;
  const now = await readMessage(graph, message.id, spec);
  const passedOver = reasonToPassOver(now, message);
  if (passedOver !== undefined) {
    outcome.skipped.push({ target: change.target, before: change.before, reason: passedOver });
    return;
  }
  const written = now.ok ? await spec.write(graph, message) : now;
  if (written.ok) {
    outcome.changes.push({ ...change, ...written.value });
    return;
  }
  outcome.firstError ??= written.error;
  outcome.failed.push({ target: change.target, before: change.before, reason: written.error.message });
};

// Why the write passes over the message as it is now: gone, or not as the plan read it.
const reasonToPassOver = (now: Result<MailMessage, GraphError>, planned: MailMessage): string | undefined => {
  if (!now.ok) return isNotThere(now.error) ? GONE : undefined;
  return now.value.parentFolderId === planned.parentFolderId && now.value.changeKey === planned.changeKey ? undefined : CHANGED;
};

const applyBatch = async (graph: WriteGraph, batch: Batch, spec: MailBatchSpec): Promise<ReturnType<WritePlan['apply']>> => {
  const outcome: Outcome = { changes: [], skipped: [...batch.skipped], failed: [] };
  for (const item of batch.planned) await applyOne(graph, item, spec, outcome);
  // A run that changed nothing and failed is an error; with one change made, the run is recorded so its undo can run.
  if (outcome.changes.length === 0 && outcome.firstError !== undefined) return err(outcome.firstError);
  return ok({ value: { done: outcome.changes.length, skipped: outcome.skipped, failed: outcome.failed }, changes: outcome.changes });
};

/** The plan of a mail state command over the messages `ids`; `folders` reads the folder names, and keeps them for the plan. */
export const planMailBatch = async (reads: GraphReads, folders: FolderReader, ids: ReadonlyArray<GraphId>, spec: MailBatchSpec): Promise<Result<WritePlan, GraphError>> => {
  const batch: Batch = { planned: [], skipped: [] };
  for (const id of ids) {
    const planned = await planOne(reads, folders, id, spec, batch);
    if (!planned.ok) return planned;
  }
  return ok({
    changes: batch.planned.map((item) => item.change),
    ...(batch.skipped.length === 0 ? {} : { skipped: batch.skipped }),
    apply: async (graph) => applyBatch(graph, batch, spec),
  });
};
