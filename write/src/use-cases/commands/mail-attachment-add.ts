import type { AttachmentFile } from '@ask-marcel/office-core/domain/attachment-upload.ts';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads, ItemIds, PlannedChange } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import type { DraftAttachment } from './mail-draft-reader.ts';
import { attachmentsOf, isDraftNow, readDraftAttachments } from './mail-draft-reader.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';

/*
 * The op that attaches a local file to a mail draft (on the op engine of
 * the To Do commands). Just before the upload it reads the draft again, and passes
 * over a draft that was sent after the plan. The upload gives the id of the
 * new attachment. An upload session can end with no id: then the op looks for
 * the new attachment by its name among the attachments that were not there
 * before. The undo removes the attachment.
 */

/** A mail above this size is refused by many mail servers: the dry-run warns of it. */
export const MAIL_WARNING_BYTES = 25_000_000;

const idUnknown: GraphError = {
  type: 'api_error',
  status: 502,
  code: 'attachment_id_unknown',
  message: 'Graph attached the file, but gave no id for it, so the undo cannot remove it. Remove it in Outlook if you do not want it.',
};

/** The warning of the dry-run for the file and the attachments of the draft, or undefined. */
export const attachWarning = (file: AttachmentFile, attachments: ReadonlyArray<DraftAttachment>): string | undefined => {
  const name = file.name.toLowerCase();
  const total = attachments.reduce((sum, attachment) => sum + attachment.size, file.bytes.byteLength);
  const warnings = [
    attachments.some((attachment) => attachment.name.toLowerCase() === name) ? 'The draft has an attachment with this name already.' : '',
    total > MAIL_WARNING_BYTES ? `The mail will be ${total} bytes with its attachments: many mail servers refuse a mail above ${MAIL_WARNING_BYTES} bytes.` : '',
  ].filter((warning) => warning !== '');
  return warnings.length === 0 ? undefined : warnings.join(' ');
};

// The id of the new attachment with the name `name`: one that was not in `before`.
const findNew = async (reads: GraphReads, messageId: GraphId, name: string, before: ReadonlyArray<DraftAttachment>): Promise<Result<GraphId, GraphError>> => {
  const all = await readDraftAttachments(reads, messageId);
  if (!all.ok) return all;
  const known = new Set(before.map((attachment) => attachment.id));
  const found = all.value.find((attachment) => attachment.name === name && !known.has(attachment.id));
  return found === undefined ? err(idUnknown) : ok(found.id);
};

/** Attaches `file` to the draft. `before` is the list of the attachments that the plan read. */
export const addDraftFileOp = (messageId: GraphId, file: AttachmentFile, shown: ItemIds, before: ReadonlyArray<DraftAttachment>): TodoOp => {
  const change: PlannedChange = { action: 'add-attachment', target: attachmentsOf(messageId), before: shown };
  return {
    change,
    recheck: async (reads) => isDraftNow(reads, messageId),
    write: async (graph) => {
      const added = await graph.uploadAttachment({ kind: 'mail-draft', messageId }, file);
      if (!added.ok) return outcomeOf(added);
      const id = added.value.id === undefined ? await findNew(graph, messageId, file.name, before) : ok(added.value.id);
      if (!id.ok) return outcomeOf(id);
      const ids = { messageId, attachmentId: id.value };
      return outcomeOf(ok({ ...change, after: ids, undo: { kind: 'delete-draft-attachment', data: { ...ids, name: file.name } } }));
    },
  };
};
