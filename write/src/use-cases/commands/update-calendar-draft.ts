import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { GraphReads, PlanWrite, WritePlan } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { draftBodyText, htmlBodyWith, ownTextOf, plannedInBody, teamsBlockOf } from '../../domain/calendar-draft-body.ts';
import { parseGraphId } from '../../domain/graph-id.ts';
import { CALENDAR_DRAFT_RESPONSE, DRAFT_BODY_FILE, EVENT_ID_OPTION } from './calendar-draft-texts.ts';
import { updateDraftOp } from './calendar-draft-ops.ts';
import type { DraftParams } from './calendar-draft-options.ts';
import { DRAFT_FIELD_OPTIONS, draftFieldShape, plainFieldsOf, plannedParam, timeFieldsOf } from './calendar-draft-options.ts';
import type { CalendarDraft } from './calendar-draft-reader.ts';
import { readDraft, readHtmlBody } from './calendar-draft-reader.ts';
import { createBodyFilePlan } from './draft-body-file.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { readTimeZone } from './todo-reader.ts';

/*
 * Changes the fields of a calendar draft that create-calendar-draft made. The
 * plan reads the event first and refuses an event with attendees, with no
 * mark, with no prefix "[DRAFT] ", of a series, or of another organizer,
 * before any write. A new body text keeps the planned attendees and the Teams
 * link of before, and new planned attendees keep the text. The undo writes
 * the fields of before back, only if the event did not change after the run.
 */

const schema = z.object({ eventId: z.string().min(1), ...draftFieldShape });

type BodyChange = { readonly field: Record<string, unknown>; readonly before: Record<string, unknown>; readonly text: string };

// The new body, the body of before (HTML, with its Teams block) and the text that the undo check compares.
const bodyChange = async (reads: GraphReads, draft: CalendarDraft, params: DraftParams): Promise<Result<BodyChange | undefined, GraphError>> => {
  const planned = plannedParam(params);
  if (params.body === undefined && planned === undefined) return ok(undefined);
  const html = await readHtmlBody(reads, draft.id);
  if (!html.ok) return html;
  const text = draftBodyText(params.body ?? ownTextOf(draft.bodyText, draft.isOnlineMeeting), planned ?? plannedInBody(draft.bodyText, draft.isOnlineMeeting));
  const block = draft.isOnlineMeeting ? teamsBlockOf(html.value) : '';
  const field = block === '' ? { contentType: 'text', content: text } : { contentType: 'html', content: htmlBodyWith(text, block) };
  return ok({ field: { body: field }, before: { body: { contentType: 'html', content: html.value } }, text });
};

const nothing: GraphError = { type: 'validation_error', code: 'nothing_to_update', message: 'Give at least one field to change, for example --subject or --start.' };

// The time fields of the options, checked with the times and the all-day form of the event.
const timesOf = (params: DraftParams, draft: CalendarDraft, timeZone: string): Result<Record<string, unknown>, GraphError> => {
  const allDay = params.allDay === undefined ? draft.isAllDay : params.allDay === 'true';
  return timeFieldsOf(params, allDay, timeZone, draft);
};

// The value of each field before the run: a field that Graph did not give is empty; the body is the HTML of before.
const beforeOf = (draft: CalendarDraft, fields: Readonly<Record<string, unknown>>, body: BodyChange | undefined): Record<string, unknown> => ({
  ...Object.fromEntries(Object.keys(fields).map((key) => [key, draft.fields[key] ?? null])),
  ...body?.before,
});

const planOf = async (reads: GraphReads, params: z.infer<typeof schema>, timeZone: string, draft: CalendarDraft): Promise<Result<WritePlan, GraphError>> => {
  const times = timesOf(params, draft, timeZone);
  if (!times.ok) return times;
  const body = await bodyChange(reads, draft, params);
  if (!body.ok) return body;
  const fields = { ...plainFieldsOf(params), ...times.value, ...body.value?.field };
  if (Object.keys(fields).length === 0) return err(nothing);
  return ok(todoPlan({ ops: [updateDraftOp(draft, fields, beforeOf(draft, fields, body.value), body.value?.text)] }));
};

const plan: PlanWrite = createBodyFilePlan(async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const id = parseGraphId(parsed.value.eventId, '--event-id');
  if (!id.ok) return id;
  const zone = await readTimeZone(reads);
  if (!zone.ok) return zone;
  const draft = await readDraft(reads, id.value, zone.value);
  return draft.ok ? planOf(reads, parsed.value, zone.value, draft.value) : draft;
}, DRAFT_BODY_FILE);

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Change the fields of a DRAFT event that create-calendar-draft made: the subject (it keeps the prefix "[DRAFT] "), the times, the location, the body, the planned attendees (as text) and the other options. An event with attendees is refused before any write, so no update or invitation is sent.',
  category: 'calendar',
  graphMethod: 'PATCH',
  graphPathTemplate: '/me/events/{event-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-update',
  options: [EVENT_ID_OPTION, ...DRAFT_FIELD_OPTIONS],
  example: '{write-bin} update-calendar-draft --event-id "AAMkAD..." --start 2026-11-20T15:00 --end 2026-11-20T16:00 --dry-run',
  bodyTemplate:
    "{ subject?: '[DRAFT] {subject}', start?: { dateTime: '{start}:00', timeZone: <time zone of the mailbox> }, end?: { dateTime: '{end}:00', timeZone: <time zone of the mailbox> }, isAllDay?: {all-day}, location?: { displayName: '{location}' }, body?: { contentType: 'text', content: '{body}' (or the text of {body-file}), then a line 'Planned attendees (NOT invited): {planned-attendees}' } (HTML with the Teams block of before on an online meeting), showAs?: '{show-as}', categories?: [{categories}], isReminderOn?: true or false, reminderMinutesBeforeStart?: {reminder} }",
  effect: 'calendar-draft',
  undo: 'writes the fields of before back, only if the event has no attendees and the fields and the body text that the run set did not change',
  scopesRequired: ['Calendars.ReadWrite', 'MailboxSettings.Read'],
  responseShape: CALENDAR_DRAFT_RESPONSE,
};

export { execute, meta, plan, schema };
