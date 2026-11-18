import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { parseGraphId } from '../../domain/graph-id.ts';
import { isDay } from '../../domain/todo-fields.ts';
import { planMailBatch } from './mail-batch.ts';
import { fieldForm, writeField } from './mail-field-write.ts';
import { createFolderReader } from './mail-folder-reader.ts';
import { BATCH_RESPONSE, createExecute, parseParams } from './mail-state-command.ts';
import { FIELD_UNDO } from './mark-mail-message.ts';

/*
 * Sets the follow-up flag of one message: flagged, complete or not flagged,
 * with an optional due date for a flag. Graph takes a due date only with a
 * start date, so the start date is the due date. The undo gives the message
 * its flag of before.
 */

const schema = z.object({
  messageId: z.string().min(1),
  status: z.enum(['flagged', 'complete', 'notFlagged']),
  due: z.string().refine(isDay, 'must be a day in the form YYYY-MM-DD, for example 2026-11-20').optional(),
});

type Params = z.infer<typeof schema>;

const dueOnlyWithFlag: GraphError = { type: 'validation_error', message: '--due: a due date goes with --status flagged only.' };

/** The flag that Graph takes for the params: `startDateTime` and `dueDateTime` go together. */
export const flagOf = (params: Params): Result<Record<string, unknown>, GraphError> => {
  if (params.due === undefined) return ok({ flagStatus: params.status });
  if (params.status !== 'flagged') return err(dueOnlyWithFlag);
  const day = { dateTime: `${params.due}T00:00:00`, timeZone: 'UTC' };
  return ok({ flagStatus: 'flagged', startDateTime: day, dueDateTime: day });
};

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const id = parseGraphId(parsed.value.messageId, '--message-id');
  if (!id.ok) return id;
  const flag = flagOf(parsed.value);
  if (!flag.ok) return flag;
  const wanted = fieldForm('flag', flag.value);
  return planMailBatch(reads, createFolderReader(reads), [id.value], {
    action: 'set-flag',
    select: ['flag'],
    target: (messageId) => `/me/messages/${messageId}`,
    decide: (message) => (fieldForm('flag', message.fields['flag']) === wanted ? { skip: `The message has the flag "${wanted}" already.` } : 'change'),
    shown: (message) => ({ flag: fieldForm('flag', message.fields['flag']) }),
    write: async (graph, message) => writeField(graph, message, 'flag', flag.value),
  });
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Set the follow-up flag of one mail message: flagged, complete or not flagged (--status), with an optional due date for a flag (--due).',
  category: 'mail',
  graphMethod: 'PATCH',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/resources/followupflag',
  options: [
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      description: 'The id of the message, as get-mail-message or list-mail-folder-messages return it.',
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'status',
      key: 'status',
      required: true,
      description: '`flagged` flags the message for follow-up, `complete` marks the follow-up done, `notFlagged` removes the flag.',
      argumentHint: { kind: 'magicValue', values: ['flagged', 'complete', 'notFlagged'] },
    },
    {
      name: 'due',
      key: 'due',
      required: false,
      description: 'The due day of the flag, YYYY-MM-DD (UTC), with --status flagged only. Graph takes a due date only with a start date, so the start date is the same day.',
      argumentHint: { kind: 'iso8601' },
    },
  ],
  example: '{write-bin} set-mail-flag --message-id "AAMkAD..." --status flagged --due 2026-11-20',
  bodyTemplate:
    "{ flag: { flagStatus: '{status}', startDateTime?: { dateTime: '{due}T00:00:00', timeZone: 'UTC' }, dueDateTime?: { dateTime: '{due}T00:00:00', timeZone: 'UTC' } } }",
  effect: 'mail-state',
  undo: FIELD_UNDO,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape: BATCH_RESPONSE,
};

export { execute, meta, plan, schema };
