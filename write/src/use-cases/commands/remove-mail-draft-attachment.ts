import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { parseGraphId } from '../../domain/graph-id.ts';
import { readRemovableFile, removeDraftFileOp, UNDO_FILE_MAX_BYTES } from './mail-attachment-remove.ts';
import { DRAFT_ATTACHMENT_RESPONSE, MESSAGE_ID_OF_DRAFT } from './mail-attachment-texts.ts';
import { requireDraft } from './mail-draft-reader.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Removes one file attachment of an unsent mail draft. Before the delete, the
 * run reads the bytes of the file, and the run journal keeps them, so the undo
 * can attach the file again. The plan refuses an attachment that the undo
 * cannot make again: an attached item, a link, an inline picture of the body,
 * or a file above the limit of the journal.
 */

const schema = z.object({ messageId: z.string().min(1), attachmentId: z.string().min(1) });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const messageId = parseGraphId(parsed.value.messageId, '--message-id');
  if (!messageId.ok) return messageId;
  const attachmentId = parseGraphId(parsed.value.attachmentId, '--attachment-id');
  if (!attachmentId.ok) return attachmentId;
  const draft = await requireDraft(reads, messageId.value);
  if (!draft.ok) return draft;
  const file = await readRemovableFile(reads, messageId.value, attachmentId.value);
  if (!file.ok) return file;
  return ok(todoPlan({ ops: [removeDraftFileOp(messageId.value, file.value)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: `Remove one file attachment of an unsent mail draft. The run journal keeps the bytes of the file (in ~/.ask-marcel/runs, for 30 days), so its undo can attach the file again. A message that is not a draft, a draft in the Outbox, an attached item, a link, a picture in the body and a file above ${UNDO_FILE_MAX_BYTES} bytes are refused before any write. Over MCP it runs on run-destructive-command.`,
  category: 'mail',
  graphMethod: 'DELETE',
  graphPathTemplate: '/me/messages/{message-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/attachment-delete',
  options: [
    MESSAGE_ID_OF_DRAFT,
    {
      name: 'attachment-id',
      key: 'attachmentId',
      required: true,
      description: 'The id of the attachment, as list-mail-attachments returns it.',
      argumentHint: { kind: 'idOrName' },
    },
  ],
  example: '{write-bin} remove-mail-draft-attachment --message-id "AAMkAD..." --attachment-id "AAMkAD..." --dry-run',
  effect: 'delete',
  undo: 'attaches the file again from the bytes in the run journal, only if the message is still an unsent draft out of the Outbox. It is a new attachment with a new id',
  scopesRequired: ['Mail.ReadWrite'],
  responseShape: DRAFT_ATTACHMENT_RESPONSE,
};

export { execute, meta, plan, schema };
