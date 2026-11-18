import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoCheck, UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { graphIdSchema } from '../../domain/graph-id.ts';
import { attachmentsOf, isDraftNow } from './mail-draft-reader.ts';
import { checkRead, createTodoUndoHandler, notThere } from './todo-undo-handler.ts';

/*
 * The undo of the attachment commands of a mail draft. Each undo runs only
 * while the message is still an unsent draft out of the Outbox. The undo of
 * an added attachment removes it, only if it is still there with its name.
 * The undo of a removed attachment attaches the file again from the bytes
 * that the run journal keeps: Graph gives the new attachment a new id.
 */

const addedSchema = z.object({ messageId: graphIdSchema, attachmentId: graphIdSchema, name: z.string() });
const reattachSchema = z.object({ messageId: graphIdSchema, file: z.object({ name: z.string(), contentType: z.string().optional(), contentBytes: z.string() }) });
const nameSchema = z.object({ name: z.string() });

// `changed` when the message is no longer an unsent draft, else the outcome of `then`.
const whileDraft = async (reads: GraphReads, messageId: GraphId, then: () => Promise<Result<UndoCheck, GraphError>>): Promise<Result<UndoCheck, GraphError>> => {
  const draft = await isDraftNow(reads, messageId);
  if (!draft.ok) return notThere(draft.error);
  return draft.value ? then() : ok('changed');
};

/** The handler of each kind of undo of the attachment commands of a mail draft. */
export const DRAFT_ATTACHMENT_UNDO_HANDLERS: UndoHandlers = {
  'delete-draft-attachment': createTodoUndoHandler(
    addedSchema,
    async (reads, data) =>
      whileDraft(reads, data.messageId, async () =>
        checkRead(await reads.get(`${attachmentsOf(data.messageId)}/${data.attachmentId}?$select=name`), (attachment) => nameSchema.safeParse(attachment).data?.name === data.name)
      ),
    async (graph, data) => graph.delete(`${attachmentsOf(data.messageId)}/${data.attachmentId}`)
  ),
  'reattach-draft-file': createTodoUndoHandler(
    reattachSchema,
    async (reads, data) => whileDraft(reads, data.messageId, async () => ok('ready')),
    async (graph, data) => {
      const { name, contentType, contentBytes } = data.file;
      const bytes = Uint8Array.from(Buffer.from(contentBytes, 'base64'));
      const added = await graph.uploadAttachment({ kind: 'mail-draft', messageId: data.messageId }, { name, bytes, ...(contentType === undefined ? {} : { contentType }) });
      if (!added.ok) return added;
      return ok(added.value.id === undefined ? {} : { now: { messageId: data.messageId, attachmentId: added.value.id } });
    }
  ),
};
