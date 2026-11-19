import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoCheck, UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { draftTextOf, sameText } from '../../domain/calendar-draft-body.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { graphIdSchema, isGraphId } from '../../domain/graph-id.ts';
import { comparedOf, DRAFT_FIELDS } from './calendar-draft-ops.ts';
import { DRAFT_MARK, eventPath, isNotDraft, readDraft } from './calendar-draft-reader.ts';
import { unreadable } from './todo-reader.ts';
import { canonical, createTodoUndoHandler, notThere } from './todo-undo-handler.ts';

/*
 * The undo of the calendar drafts. Each check reads the event with the guard
 * of the commands: an event that lost the mark or the prefix, or got
 * attendees, is `changed`, so the undo never deletes or changes a meeting. The
 * check compares the fields that the run set, as Graph gave them back, and the
 * body text that the run wrote. The undo of a delete makes the event again:
 * Graph gives it a new id, and a new Teams link to an online meeting.
 */

const fieldsSchema = z.record(z.string(), z.unknown()).refine((fields) => Object.keys(fields).every((key) => DRAFT_FIELDS.has(key)));
const createdSchema = z.object({ eventId: graphIdSchema, set: z.record(z.string(), z.unknown()), text: z.string() });
const restoreSchema = z.object({ eventId: graphIdSchema, before: fieldsSchema, set: z.record(z.string(), z.unknown()), text: z.string().optional() });
const recreateSchema = z.object({ copy: fieldsSchema });

const checkDraft = async (reads: GraphReads, eventId: GraphId, set: Readonly<Record<string, unknown>>, text: string | undefined): Promise<Result<UndoCheck, GraphError>> => {
  const read = await readDraft(reads, eventId);
  if (!read.ok) return isNotDraft(read.error) ? ok('changed') : notThere(read.error);
  const now = comparedOf(read.value.fields, Object.keys(set));
  const same = Object.keys(set).every((key) => canonical(now[key]) === canonical(set[key])) && (text === undefined || sameText(draftTextOf(read.value.bodyText), text));
  return ok(same ? 'ready' : 'changed');
};

const idSchema = z.object({ id: z.string() });

/** The handler of each kind of undo of a calendar draft. */
export const CALENDAR_UNDO_HANDLERS: UndoHandlers = {
  'delete-created-calendar-draft': createTodoUndoHandler(
    createdSchema,
    async (reads, data) => checkDraft(reads, data.eventId, data.set, data.text),
    async (graph, data) => graph.delete(eventPath(data.eventId))
  ),
  'restore-calendar-draft-fields': createTodoUndoHandler(
    restoreSchema,
    async (reads, data) => checkDraft(reads, data.eventId, data.set, data.text),
    async (graph, data) => graph.patch(eventPath(data.eventId), data.before)
  ),
  'recreate-calendar-draft': createTodoUndoHandler(
    recreateSchema,
    async () => ok('ready'),
    async (graph, data) => {
      const created = await graph.post('/me/events', { ...data.copy, singleValueExtendedProperties: [DRAFT_MARK] });
      if (!created.ok) return created;
      const event = idSchema.safeParse(created.value);
      return event.success && isGraphId(event.data.id) ? ok({ now: { eventId: event.data.id } }) : err(unreadable('the event it made again'));
    }
  ),
};
