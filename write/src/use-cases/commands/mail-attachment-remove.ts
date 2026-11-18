import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads, PlannedChange } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { isNotThere } from './mail-folder-reader.ts';
import { attachmentsOf, isDraftNow } from './mail-draft-reader.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import { unreadable } from './todo-reader.ts';

/*
 * The op that removes a file attachment of a mail draft. The plan reads the
 * attachment and refuses one that the undo cannot make again. Just before the
 * delete, the op reads the draft and the attachment again, then the bytes of
 * the file, which the run journal keeps for the undo.
 */

/** The largest file that the run journal keeps for an undo, as for a To Do attachment. */
export const UNDO_FILE_MAX_BYTES = 25_000_000;

/** An attachment of a draft as the plan reads it: no bytes, and a size larger than the file. */
export type DraftFile = { readonly id: GraphId; readonly name: string; readonly size: number };

const FILE_TYPE = '#microsoft.graph.fileAttachment';

const metaSchema = z.object({
  '@odata.type': z.string().optional(),
  name: z.string(),
  size: z.number(),
  isInline: z.boolean().optional(),
});
const contentSchema = z.object({ contentBytes: z.string(), contentType: z.string().nullish() });
const oneContentSchema = z.union([contentSchema, z.object({ value: contentSchema }).transform((answer) => answer.value)]);

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message: `--attachment-id: ${message}` });

const notAttached = (): Result<never, GraphError> =>
  refused('mail_attachment_not_found', 'the draft has no attachment with this id. list-mail-attachments gives the id of each attachment of a message.');

// The refusal of an attachment that the undo cannot make again, or undefined.
const refusalFor = (attachment: z.infer<typeof metaSchema>): Result<never, GraphError> | undefined => {
  if ((attachment['@odata.type'] ?? FILE_TYPE) !== FILE_TYPE) {
    return refused('not_a_file_attachment', 'the attachment is an attached item or a link, not a file. The undo cannot make it again, so this command removes files only.');
  }
  if (attachment.isInline === true) {
    return refused('inline_attachment', 'the attachment is a picture in the body of the mail. The undo cannot put it back in the body, so this command does not remove it.');
  }
  if (attachment.size <= UNDO_FILE_MAX_BYTES) return undefined;
  return refused(
    'attachment_too_large_for_undo',
    `the attachment is ${attachment.size} bytes. The run journal keeps at most ${UNDO_FILE_MAX_BYTES} bytes of a file for the undo, so this command does not remove it. Remove it in Outlook.`
  );
};

const readMeta = async (reads: GraphReads, messageId: GraphId, attachmentId: GraphId): Promise<Result<z.infer<typeof metaSchema>, GraphError>> => {
  const read = await reads.get(`${attachmentsOf(messageId)}/${attachmentId}?$select=id,name,size,contentType,isInline`);
  if (!read.ok) return read;
  const meta = metaSchema.safeParse(read.value);
  return meta.success ? ok(meta.data) : err(unreadable('an attachment of a mail draft'));
};

/** The attachment of the draft that the undo can make again, or the refusal. */
export const readRemovableFile = async (reads: GraphReads, messageId: GraphId, attachmentId: GraphId): Promise<Result<DraftFile, GraphError>> => {
  const meta = await readMeta(reads, messageId, attachmentId);
  if (!meta.ok) return isNotThere(meta.error) ? notAttached() : meta;
  return refusalFor(meta.value) ?? ok({ id: attachmentId, name: meta.value.name, size: meta.value.size });
};

// True when the message is still a draft and the attachment is as the plan read it. A 404 says that it is gone.
const sameFile = async (reads: GraphReads, messageId: GraphId, file: DraftFile): Promise<Result<boolean, GraphError>> => {
  const draft = await isDraftNow(reads, messageId);
  if (!draft.ok || !draft.value) return draft;
  const meta = await readMeta(reads, messageId, file.id);
  return meta.ok ? ok(meta.value.name === file.name && meta.value.size === file.size) : meta;
};

/** Removes the attachment. The run journal keeps its bytes, so the undo can attach the file again. */
export const removeDraftFileOp = (messageId: GraphId, file: DraftFile): TodoOp => {
  const target = `${attachmentsOf(messageId)}/${file.id}`;
  const change: PlannedChange = { action: 'remove-attachment', target, before: { messageId, attachmentId: file.id, file: file.name, size: String(file.size) } };
  return {
    change,
    recheck: async (reads) => sameFile(reads, messageId, file),
    write: async (graph) => {
      const read = await graph.get(target);
      if (!read.ok) return outcomeOf(read);
      const content = oneContentSchema.safeParse(read.value);
      if (!content.success) return outcomeOf(err(unreadable('the bytes of the attachment')));
      const deleted = await graph.delete(target);
      if (!deleted.ok) return outcomeOf(deleted);
      const { contentBytes, contentType } = content.data;
      const saved = { name: file.name, contentBytes, ...(typeof contentType === 'string' ? { contentType } : {}) };
      return outcomeOf(ok({ ...change, after: {}, undo: { kind: 'reattach-draft-file', data: { messageId, file: saved } } }));
    },
  };
};
