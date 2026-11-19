import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { parseGraphId } from '../../domain/graph-id.ts';
import { CALENDAR_DRAFT_RESPONSE, EVENT_ID_OPTION } from './calendar-draft-texts.ts';
import { deleteDraftOp } from './calendar-draft-ops.ts';
import { readDraft } from './calendar-draft-reader.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Deletes a calendar draft that create-calendar-draft made. The plan reads the
 * event first and refuses an event with attendees (its delete would send them
 * a cancellation), with no mark or no prefix "[DRAFT] ", of a series, or of
 * another organizer, before any write. The run journal keeps a copy of the
 * event, so the undo can make it again: Graph gives it a new id, and a new
 * Teams link to an online meeting.
 */

const schema = z.object({ eventId: z.string().min(1) });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const id = parseGraphId(parsed.value.eventId, '--event-id');
  if (!id.ok) return id;
  const draft = await readDraft(reads, id.value);
  return draft.ok ? ok(todoPlan({ ops: [deleteDraftOp(draft.value)] })) : draft;
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Delete a DRAFT event that create-calendar-draft made. An event with attendees is refused before any write, so no cancellation is sent. The undo makes the event again from the run journal, with a new id and, for an online meeting, a new Teams link. Over MCP it runs on run-destructive-command.',
  category: 'calendar',
  graphMethod: 'DELETE',
  graphPathTemplate: '/me/events/{event-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-delete',
  options: [EVENT_ID_OPTION],
  example: '{write-bin} delete-calendar-draft --event-id "AAMkAD..." --dry-run',
  effect: 'delete',
  undo: 'makes the event again from the copy in the run journal: Graph gives it a new id, a new Teams link to an online meeting, and a plain-text body (the formatting of the body and the old Teams link are lost)',
  scopesRequired: ['Calendars.ReadWrite'],
  responseShape: CALENDAR_DRAFT_RESPONSE,
};

export { execute, meta, plan, schema };
