import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { draftBodyText, draftSubject } from '../../domain/calendar-draft-body.ts';
import { CALENDAR_DRAFT_RESPONSE, DRAFT_BODY_FILE } from './calendar-draft-texts.ts';
import { createDraftOp } from './calendar-draft-ops.ts';
import { DRAFT_FIELD_OPTIONS, draftFieldShape, plainFieldsOf, plannedParam, timeFieldsOf } from './calendar-draft-options.ts';
import { createBodyFilePlan } from './draft-body-file.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { readTimeZone } from './todo-reader.ts';

/*
 * Creates a draft of a meeting in the calendar of the user: an event with no
 * attendees, so Graph sends no invitation. The subject gets the prefix
 * `[DRAFT] `, the planned attendees go into the body as text, and `--teams
 * true` adds a Teams link. The event gets the hidden mark of this CLI, which
 * update-calendar-draft and delete-calendar-draft ask for. The undo deletes
 * the event if it did not change and has no attendees.
 */

const schema = z.object({
  ...draftFieldShape,
  subject: z.string().trim().min(1),
  start: z.string(),
  end: z.string(),
  teams: z.enum(['true', 'false']).optional(),
});

const TEAMS = { isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' };

const plan: PlanWrite = createBodyFilePlan(async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const zone = await readTimeZone(reads);
  if (!zone.ok) return zone;
  const times = timeFieldsOf(parsed.value, parsed.value.allDay === 'true', zone.value, parsed.value);
  if (!times.ok) return times;
  const planned = plannedParam(parsed.value) ?? [];
  const text = draftBodyText(parsed.value.body ?? '', planned);
  const teams = parsed.value.teams === 'true';
  const fields = { ...plainFieldsOf(parsed.value), ...times.value, body: { contentType: 'text', content: text }, ...(teams ? TEAMS : {}) };
  const before = {
    subject: draftSubject(parsed.value.subject),
    start: parsed.value.start,
    end: parsed.value.end,
    timeZone: zone.value,
    ...(planned.length === 0 ? {} : { plannedAttendees: planned.join('; ') }),
    ...(teams ? { teams: 'true' } : {}),
  };
  return ok(todoPlan({ ops: [createDraftOp(fields, text, before)] }));
}, DRAFT_BODY_FILE);

const execute = createExecute(plan);

const required = new Set(['subject', 'start', 'end']);

const meta: WriteCommandMeta = {
  summary:
    'Create a DRAFT event in your calendar: the subject gets the prefix "[DRAFT] ", and the event has NO attendees, so no invitation is sent. The planned attendees go into the body as text only. --teams true adds a Teams link. You send the invitation yourself from Outlook.',
  category: 'calendar',
  graphMethod: 'POST',
  graphPathTemplate: '/me/events',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-post-events',
  options: [
    ...DRAFT_FIELD_OPTIONS.map((option) => (required.has(option.name) ? { ...option, required: true } : option)),
    {
      name: 'teams',
      key: 'teams',
      required: false,
      description: '`true` adds a Teams link to the event. With no attendees, Teams sends nobody an invitation.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example:
    '{write-bin} create-calendar-draft --subject "Budget review" --start 2026-11-20T14:00 --end 2026-11-20T15:00 --planned-attendees "a@example.com; b@example.com" --teams true --dry-run',
  bodyTemplate:
    "{ subject: '[DRAFT] {subject}', start: { dateTime: '{start}:00', timeZone: <time zone of the mailbox> }, end: { dateTime: '{end}:00', timeZone: <time zone of the mailbox> }, body: { contentType: 'text', content: '{body}' (or the text of {body-file}), then a line 'Planned attendees (NOT invited): {planned-attendees}' }, isAllDay?: {all-day}, location?: { displayName: '{location}' }, showAs?: '{show-as}', categories?: [{categories}], isReminderOn?: true, reminderMinutesBeforeStart?: {reminder}, isOnlineMeeting?: true, onlineMeetingProvider?: 'teamsForBusiness' (with {teams} true), singleValueExtendedProperties: [<the mark of this CLI>] }",
  effect: 'calendar-draft',
  undo: 'deletes the event, only if it has no attendees and the fields and the body text that the run set did not change',
  scopesRequired: ['Calendars.ReadWrite', 'MailboxSettings.Read'],
  responseShape: CALENDAR_DRAFT_RESPONSE,
};

export { execute, meta, plan, schema };
