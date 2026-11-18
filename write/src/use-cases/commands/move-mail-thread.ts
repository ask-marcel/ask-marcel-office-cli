import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { GraphReads, PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { parseGraphId } from '../../domain/graph-id.ts';
import { planMailBatch } from './mail-batch.ts';
import type { MailFolder } from './mail-folder-reader.ts';
import { createFolderReader } from './mail-folder-reader.ts';
import { moveTarget, resolveMailFolder } from './mail-folder-target.ts';
import { MAX_BATCH_MESSAGES } from './mail-message-ids.ts';
import { BATCH_RESPONSE, createExecute, parseParams } from './mail-state-command.ts';
import { MOVE_UNDO, moveSpec } from './move-mail-message.ts';

/*
 * Moves the messages of one conversation that are in the folder `--from` to
 * the folder `--to`. The other messages of the conversation (in Sent Items, in
 * Drafts, in another folder) stay where they are. The target guard and the
 * batch are those of move-mail-message.
 */

const schema = z.object({ conversationId: z.string().min(1), from: z.string().min(1), to: z.string().min(1) });

const listSchema = z.object({ value: z.array(z.object({ id: z.string() })), '@odata.nextLink': z.string().optional() });

const tooLong: GraphError = {
  type: 'validation_error',
  code: 'invalid_batch',
  message: `The conversation has more than ${MAX_BATCH_MESSAGES} messages in that folder: one run changes at most ${MAX_BATCH_MESSAGES}. Move them with move-mail-message --ids-file.`,
};

// The ids of the messages of the conversation in the folder; the id check stands before the `$filter` too.
const threadIds = async (reads: GraphReads, conversationId: GraphId, from: MailFolder): Promise<Result<ReadonlyArray<GraphId>, GraphError>> => {
  const filter = encodeURIComponent(`conversationId eq '${conversationId}'`);
  const read = await reads.get(`/me/mailFolders/${from.id}/messages?$filter=${filter}&$select=id&$top=${MAX_BATCH_MESSAGES}`);
  if (!read.ok) return read;
  const list = listSchema.safeParse(read.value);
  if (!list.success) return err({ type: 'validation_error', message: 'Graph gave the messages of the conversation in a shape that this CLI cannot read.' });
  if (list.data['@odata.nextLink'] !== undefined) return err(tooLong);
  const ids: GraphId[] = [];
  for (const message of list.data.value) {
    const id = parseGraphId(message.id, 'message');
    if (!id.ok) return id;
    ids.push(id.value);
  }
  return ok(ids);
};

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const conversationId = parseGraphId(parsed.value.conversationId, '--conversation-id');
  if (!conversationId.ok) return conversationId;
  const folders = createFolderReader(reads);
  const to = await moveTarget(folders, parsed.value.to, '--to');
  if (!to.ok) return to;
  const from = await resolveMailFolder(folders, parsed.value.from, '--from');
  if (!from.ok) return from;
  const ids = await threadIds(reads, conversationId.value, from.value);
  return ids.ok ? planMailBatch(reads, folders, ids.value, moveSpec(to.value)) : ids;
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Move the messages of one conversation that are in the folder --from to the folder --to. The messages of the conversation in other folders (Sent Items, Drafts) stay. Each moved message gets a new id, which the output shows. Refuses a move into Deleted Items, Recoverable Items, Junk Email or Outbox, or a folder in one of them.',
  category: 'mail',
  graphMethod: 'POST',
  graphPathTemplate:
    "GET /me/mailFolders/{from}/messages?$filter=conversationId eq '{conversation-id}', then POST /me/messages/<id>/move for each message (one call each), to the folder {to}",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-move',
  options: [
    {
      name: 'conversation-id',
      key: 'conversationId',
      required: true,
      description: 'The conversationId of the thread, as get-mail-message or list-mail-folder-messages return it.',
    },
    {
      name: 'from',
      key: 'from',
      required: true,
      description:
        'The folder whose messages of the conversation move: a well-known name (inbox), a path of folder names (`Inbox/Projects`), a folder name at the top, or a folder id.',
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'to',
      key: 'to',
      required: true,
      description: 'The folder to move the messages to, named as --from is.',
      argumentHint: { kind: 'idOrName' },
    },
  ],
  example: '{write-bin} move-mail-thread --conversation-id "AAQkAD..." --from inbox --to archive --dry-run',
  bodyTemplate: "{ destinationId: '<the id of the folder {to}>' }",
  effect: 'mail-state',
  undo: MOVE_UNDO,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape: BATCH_RESPONSE,
};

export { execute, meta, plan, schema };
