import { z } from 'zod';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { planMailBatch } from './mail-batch.ts';
import { fieldForm, writeField } from './mail-field-write.ts';
import { createFolderReader } from './mail-folder-reader.ts';
import { messageIdsOf } from './mail-message-ids.ts';
import { BATCH_RESPONSE, createExecute, IDS_FILE_OPTION, MESSAGE_ID_OPTION, parseParams } from './mail-state-command.ts';

/*
 * Marks one or more messages read or unread. A message that is in that state
 * already is passed over. The undo gives each message its read state of before.
 */

const schema = z.object({ messageId: z.string().optional(), idsFile: z.string().optional(), read: z.enum(['true', 'false']) });

/** What the undo of a state change does, as the docs say it. */
export const FIELD_UNDO = 'gives each message the value of before, only if the message still has the value that the run set';

const plan: PlanWrite = async (reads, params, files) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const ids = await messageIdsOf(parsed.value, files);
  if (!ids.ok) return ids;
  const read = parsed.value.read === 'true';
  return planMailBatch(reads, createFolderReader(reads), ids.value, {
    action: 'mark-message',
    select: ['isRead'],
    target: (id) => `/me/messages/${id}`,
    decide: (message) => (message.fields['isRead'] === read ? { skip: `The message is ${read ? 'read' : 'unread'} already.` } : 'change'),
    shown: (message) => ({ isRead: fieldForm('isRead', message.fields['isRead']) }),
    write: async (graph, message) => writeField(graph, message, 'isRead', read),
  });
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Mark one or more mail messages read or unread (--read true or false): --message-id for one message or a few, --ids-file for a batch. A message that is in that state already is passed over.',
  category: 'mail',
  graphMethod: 'PATCH',
  graphPathTemplate: '/me/messages/{message-id} (one call for each message; the ids of {ids-file} too)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-update',
  options: [
    MESSAGE_ID_OPTION,
    IDS_FILE_OPTION,
    { name: 'read', key: 'read', required: true, description: '`true` marks the messages read, `false` unread.', argumentHint: { kind: 'magicValue', values: ['true', 'false'] } },
  ],
  example: '{write-bin} mark-mail-message --message-id "AAMkAD..." --read false',
  bodyTemplate: '{ isRead: {read} }',
  effect: 'mail-state',
  undo: FIELD_UNDO,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape: BATCH_RESPONSE,
};

export { execute, meta, plan, schema };
