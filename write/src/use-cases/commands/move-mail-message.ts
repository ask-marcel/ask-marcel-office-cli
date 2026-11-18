import { z } from 'zod';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { MailBatchSpec } from './mail-batch.ts';
import { planMailBatch } from './mail-batch.ts';
import type { MailFolder } from './mail-folder-reader.ts';
import { createFolderReader } from './mail-folder-reader.ts';
import { moveTarget } from './mail-folder-target.ts';
import { messageIdsOf } from './mail-message-ids.ts';
import { moveTo } from './mail-move-write.ts';
import { BATCH_RESPONSE, createExecute, IDS_FILE_OPTION, MESSAGE_ID_OPTION, parseParams } from './mail-state-command.ts';

/*
 * Moves one or more messages to a folder. A move into Deleted Items,
 * Recoverable Items, Junk Email or Outbox, or into a folder in one of them, is
 * refused (mail-folder-target.ts). A moved message gets a new id: the output
 * shows the id before and after.
 */

const schema = z.object({ messageId: z.string().optional(), idsFile: z.string().optional(), to: z.string().min(1) });

/** What the undo of a move does, as the docs say it. */
export const MOVE_UNDO = 'moves each message back to the folder it was in (the message gets a new id again), only if it is still in the folder that the run put it in';

/** The batch of a move to the folder `to`: a message already there is passed over. */
export const moveSpec = (to: MailFolder): MailBatchSpec => ({
  action: 'move-message',
  select: [],
  target: (id) => `/me/messages/${id}/move`,
  decide: (message) => (message.parentFolderId === to.id ? { skip: `The message is in "${to.displayName}" already.` } : 'change'),
  shown: () => ({}),
  write: async (graph, message) => moveTo(graph, message, to),
});

const plan: PlanWrite = async (reads, params, files) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const folders = createFolderReader(reads);
  const to = await moveTarget(folders, parsed.value.to, '--to');
  if (!to.ok) return to;
  const ids = await messageIdsOf(parsed.value, files);
  return ids.ok ? planMailBatch(reads, folders, ids.value, moveSpec(to.value)) : ids;
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Move one or more mail messages to a folder: --message-id for one message or a few, --ids-file for a batch. A moved message gets a new id, which the output shows. Refuses a move into Deleted Items, Recoverable Items, Junk Email or Outbox, or a folder in one of them. Do a dry-run first: it shows the current folder of each message, so you see a draft that a search found in place of the received message.',
  category: 'mail',
  graphMethod: 'POST',
  graphPathTemplate: '/me/messages/{message-id}/move (one call for each message; the ids of {ids-file} too), to the folder {to}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-move',
  options: [
    MESSAGE_ID_OPTION,
    IDS_FILE_OPTION,
    {
      name: 'to',
      key: 'to',
      required: true,
      description:
        'The folder to move the messages to: a well-known name (inbox, archive), a path of folder names from the top (`Inbox/Projects/Alpha`), a folder name at the top, or a folder id from list-mail-folders.',
      argumentHint: { kind: 'idOrName' },
    },
  ],
  example: '{write-bin} move-mail-message --message-id "AAMkAD..." --to archive --dry-run',
  bodyTemplate: "{ destinationId: '<the id of the folder {to}>' }",
  effect: 'mail-state',
  undo: MOVE_UNDO,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape: BATCH_RESPONSE,
};

export { execute, meta, plan, schema };
