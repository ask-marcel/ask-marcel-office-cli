import { z } from 'zod';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads, ItemIds, PlannedChange } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { draftTextOf } from '../../domain/calendar-draft-body.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { isGraphId } from '../../domain/graph-id.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import type { CalendarDraft } from './calendar-draft-reader.ts';
import { DRAFT_MARK, eventPath, readDraft, sameDraft } from './calendar-draft-reader.ts';
import { unreadable } from './todo-reader.ts';

/*
 * The ops of the calendar drafts, on the op engine `todoPlan` of the To Do
 * commands: create, update and delete. An update and a delete read the
 * event again just before the write, and pass over an event that changed or
 * got attendees after the plan. Each op records how to reverse it.
 */

/** The fields of an event that this CLI writes, and so the only ones that an undo can write back. */
export const DRAFT_FIELDS: ReadonlySet<string> = new Set([
  'subject',
  'start',
  'end',
  'isAllDay',
  'location',
  'body',
  'showAs',
  'categories',
  'isReminderOn',
  'reminderMinutesBeforeStart',
  'isOnlineMeeting',
  'onlineMeetingProvider',
]);

// The fields that Exchange changes on its own after a write (the Teams join block in the body), so no undo check compares them.
const NOT_COMPARED: ReadonlySet<string> = new Set(['body', 'isOnlineMeeting', 'onlineMeetingProvider']);

const timeSchema = z.object({ dateTime: z.string() }).loose();
const answerSchema = z.object({ id: z.string(), subject: z.string(), start: timeSchema, end: timeSchema }).loose();
const locationSchema = z.object({ displayName: z.string() });

/** The date and time of a Graph time, or an empty text. */
export const dateTimeOf = (value: unknown): string => {
  const time = timeSchema.safeParse(value);
  return time.success ? time.data.dateTime : '';
};

// A location compares by its name: Graph adds other fields to it.
const comparable = (key: string, value: unknown): unknown => {
  if (key !== 'location') return value ?? null;
  const location = locationSchema.safeParse(value);
  return location.success ? location.data.displayName : '';
};

/** The values to compare of the fields `keys`, as an event gives them. */
export const comparedOf = (fields: Readonly<Record<string, unknown>>, keys: ReadonlyArray<string>): Record<string, unknown> =>
  Object.fromEntries(keys.map((key) => [key, comparable(key, fields[key])]));

/** The keys of the written fields that an undo check compares. */
export const comparedKeys = (fields: Readonly<Record<string, unknown>>): ReadonlyArray<string> => Object.keys(fields).filter((key) => !NOT_COMPARED.has(key));

/**
 * The values that the undo check will compare, read back after the write:
 * the answer of a write gives the times in the zone of the request, and a
 * read in UTC. When the read fails, the answer is the best value, and the
 * undo then says `changed` rather than undo a wrong event.
 */
const settled = async (reads: GraphReads, id: GraphId, answer: Readonly<Record<string, unknown>>, fields: Readonly<Record<string, unknown>>): Promise<Record<string, unknown>> => {
  const now = await readDraft(reads, id);
  return comparedOf(now.ok ? now.value.fields : answer, comparedKeys(fields));
};

/** The ids and the times of an event, for `before` and `after`. */
export const draftIds = (id: string, fields: Readonly<Record<string, unknown>>): ItemIds => ({
  eventId: id,
  subject: typeof fields['subject'] === 'string' ? fields['subject'] : '',
  start: dateTimeOf(fields['start']),
  end: dateTimeOf(fields['end']),
});

/** Creates the event with the mark of this CLI. `text` is the body text that the undo check compares. */
export const createDraftOp = (fields: Readonly<Record<string, unknown>>, text: string, before: ItemIds): TodoOp => {
  const change: PlannedChange = { action: 'create-calendar-draft', target: '/me/events', before };
  return {
    change,
    write: async (graph) => {
      const created = await graph.post(change.target, { ...fields, singleValueExtendedProperties: [DRAFT_MARK] });
      if (!created.ok) return outcomeOf(created);
      const event = answerSchema.safeParse(created.value);
      if (!event.success || !isGraphId(event.data.id)) return outcomeOf(err(unreadable('the event it created')));
      const set = await settled(graph, event.data.id, event.data, fields);
      return outcomeOf(ok({ ...change, after: draftIds(event.data.id, event.data), undo: { kind: 'delete-created-calendar-draft', data: { eventId: event.data.id, set, text } } }));
    },
  };
};

/** Sets the fields of the event. `before` has the value of each field before the run; `text` is the body text that it writes, if any. */
export const updateDraftOp = (draft: CalendarDraft, fields: Readonly<Record<string, unknown>>, before: Readonly<Record<string, unknown>>, text?: string): TodoOp => {
  const change: PlannedChange = { action: 'update-calendar-draft', target: eventPath(draft.id), before: draftIds(draft.id, draft.fields) };
  return {
    change,
    recheck: async (reads) => sameDraft(reads, draft),
    write: async (graph) => {
      const written = await graph.patch(change.target, fields);
      if (!written.ok) return outcomeOf(written);
      const event = answerSchema.safeParse(written.value);
      if (!event.success) return outcomeOf(err(unreadable('the event it changed')));
      const set = await settled(graph, draft.id, event.data, fields);
      const data = { eventId: draft.id, before, set, ...(text === undefined ? {} : { text }) };
      return outcomeOf(
        ok({ ...change, after: { ...draftIds(draft.id, event.data), fields: Object.keys(fields).join(', ') }, undo: { kind: 'restore-calendar-draft-fields', data } })
      );
    },
  };
};

const onlineOf = (draft: CalendarDraft): Record<string, unknown> => (draft.isOnlineMeeting ? { isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' } : {});

/** The fields that make the event again: its text body without the old Teams block, and a new Teams link if it had one. */
export const copyOf = (draft: CalendarDraft): Record<string, unknown> => {
  const kept = ['subject', 'start', 'end', 'isAllDay', 'showAs', 'categories', 'isReminderOn', 'reminderMinutesBeforeStart'].filter((key) => draft.fields[key] !== undefined);
  return {
    ...Object.fromEntries(kept.map((key) => [key, draft.fields[key]])),
    location: { displayName: comparable('location', draft.fields['location']) },
    body: { contentType: 'text', content: draftTextOf(draft.bodyText) },
    ...onlineOf(draft),
  };
};

/** Deletes the event. The run journal keeps its copy, so the undo can make it again (new id). */
export const deleteDraftOp = (draft: CalendarDraft): TodoOp => {
  const change: PlannedChange = { action: 'delete-calendar-draft', target: eventPath(draft.id), before: draftIds(draft.id, draft.fields) };
  return {
    change,
    recheck: async (reads) => sameDraft(reads, draft),
    write: async (graph) => {
      const deleted = await graph.delete(change.target);
      return outcomeOf(deleted.ok ? ok({ ...change, after: {}, undo: { kind: 'recreate-calendar-draft', data: { copy: copyOf(draft) } } }) : deleted);
    },
  };
};
