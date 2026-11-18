import { z } from 'zod';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { parseGraphId } from '../../domain/graph-id.ts';
import { planMailBatch } from './mail-batch.ts';
import { fieldForm, writeField } from './mail-field-write.ts';
import { createFolderReader } from './mail-folder-reader.ts';
import { BATCH_RESPONSE, createExecute, parseParams } from './mail-state-command.ts';
import { FIELD_UNDO } from './mark-mail-message.ts';

/*
 * Replaces the categories of one message. An empty --categories removes them
 * all. The undo gives the message its categories of before.
 */

const schema = z.object({ messageId: z.string().min(1), categories: z.string() });

/** The category names of `--categories`: separated by commas, each once, in the order given. */
export const categoriesOf = (value: string): ReadonlyArray<string> => [
  ...new Set(
    value
      .split(',')
      .map((name) => name.trim())
      .filter((name) => name !== '')
  ),
];

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const id = parseGraphId(parsed.value.messageId, '--message-id');
  if (!id.ok) return id;
  const categories = categoriesOf(parsed.value.categories);
  const wanted = fieldForm('categories', categories);
  return planMailBatch(reads, createFolderReader(reads), [id.value], {
    action: 'set-categories',
    select: ['categories'],
    target: (messageId) => `/me/messages/${messageId}`,
    decide: (message) => (fieldForm('categories', message.fields['categories']) === wanted ? { skip: 'The message has these categories already.' } : 'change'),
    shown: (message) => ({ categories: fieldForm('categories', message.fields['categories']) }),
    write: async (graph, message) => writeField(graph, message, 'categories', categories),
  });
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Replace the categories of one mail message with the names of --categories, separated by commas. An empty --categories removes all the categories. A name that is not in the category list of the mailbox shows with no color in Outlook.',
  category: 'mail',
  graphMethod: 'PATCH',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-update',
  options: [
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      description: 'The id of the message, as get-mail-message or list-mail-folder-messages return it.',
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'categories',
      key: 'categories',
      required: true,
      description: 'The new category names, separated by commas, for example "Blue category,Project". Give "" to remove all the categories.',
    },
  ],
  example: '{write-bin} set-mail-categories --message-id "AAMkAD..." --categories "Blue category,Project"',
  bodyTemplate: "{ categories: ['{categories}', ...] }",
  effect: 'mail-state',
  undo: FIELD_UNDO,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape: BATCH_RESPONSE,
};

export { execute, meta, plan, schema };
