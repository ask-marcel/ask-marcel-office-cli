import { z } from 'zod';
import { ATTACHMENT_MAX_BYTES } from '@ask-marcel/office-core/domain/attachment-upload.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { readLocalFile } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { contentTypeOf } from '../../domain/attachment-content-type.ts';
import { parseGraphId } from '../../domain/graph-id.ts';
import { addDraftFileOp, attachWarning } from './mail-attachment-add.ts';
import { DRAFT_ATTACHMENT_RESPONSE, MESSAGE_ID_OF_DRAFT } from './mail-attachment-texts.ts';
import { readDraftAttachments, requireDraft } from './mail-draft-reader.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Attaches a local file to an unsent mail draft, up to 150 MB. A file under
 * 3 MB goes in one POST, a larger one through an upload session (core
 * attachment-upload.ts). The plan reads the file and the draft first, so a
 * dry-run refuses a missing file and a message that is not a draft.
 */

const schema = z.object({ messageId: z.string().min(1), file: z.string().min(1) });

const plan: PlanWrite = async (reads, params, files) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const messageId = parseGraphId(parsed.value.messageId, '--message-id');
  if (!messageId.ok) return messageId;
  const read = await readLocalFile(files, parsed.value.file, ATTACHMENT_MAX_BYTES['mail-draft']);
  if (!read.ok) return read;
  if (read.value.bytes.byteLength === 0) {
    return err({ type: 'validation_error', code: 'attachment_file_empty', message: `${parsed.value.file} is empty. Graph does not attach an empty file.` });
  }
  const draft = await requireDraft(reads, messageId.value);
  if (!draft.ok) return draft;
  const attachments = await readDraftAttachments(reads, messageId.value);
  if (!attachments.ok) return attachments;
  const file = { ...read.value, contentType: contentTypeOf(read.value.name) };
  const warning = attachWarning(file, attachments.value);
  const shown = { messageId: messageId.value, file: file.name, size: String(file.bytes.byteLength), contentType: file.contentType, ...(warning === undefined ? {} : { warning }) };
  return ok(todoPlan({ ops: [addDraftFileOp(messageId.value, file, shown, attachments.value)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Attach a local file to an unsent mail draft, up to 150 MB. A file of 3 MB or more goes through an upload session. A message that is not a draft, or a draft in the Outbox, is refused before any write. The dry-run warns of a name that the draft has already, and of a mail above 25 MB. This command cannot send the draft.',
  category: 'mail',
  graphMethod: 'POST',
  graphPathTemplate: '/me/messages/{message-id}/attachments (or attachments/createUploadSession for 3 MB or more)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-post-attachments',
  options: [
    MESSAGE_ID_OF_DRAFT,
    {
      name: 'file',
      key: 'file',
      required: true,
      description:
        'The local file to attach, at most 150 MB, not empty. A relative path starts from the folder the CLI runs in. The attachment has the name of the file, and the content type of its extension (application/octet-stream for an extension that the CLI does not know).',
    },
  ],
  example: '{write-bin} add-mail-draft-attachment --message-id "AAMkAD..." --file ./report.pdf --dry-run',
  bodyTemplate:
    "{ '@odata.type': '#microsoft.graph.fileAttachment', name: '<the name of {file}>', contentType: '<the type of its extension>', contentBytes: '<the bytes of {file}, base64>' }",
  effect: 'draft',
  undo: 'removes the attachment, only if the message is still an unsent draft out of the Outbox and still has the attachment with its name',
  scopesRequired: ['Mail.ReadWrite'],
  responseShape: DRAFT_ATTACHMENT_RESPONSE,
};

export { execute, meta, plan, schema };
