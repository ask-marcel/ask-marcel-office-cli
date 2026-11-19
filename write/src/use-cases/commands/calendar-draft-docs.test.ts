import { describe, expect, it } from 'bun:test';
import type { CommandOptionMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { meta as createMeta } from './create-calendar-draft.ts';
import { meta as deleteMeta } from './delete-calendar-draft.ts';
import { meta as updateMeta } from './update-calendar-draft.ts';

/*
 * The docs of the calendar draft commands, as the help, the docs pages and
 * the MCP tools show them: each option, the Graph path and body, the effect,
 * what the undo does and the scopes.
 */

const EVENT_ID: CommandOptionMeta = {
  name: 'event-id',
  key: 'eventId',
  required: true,
  description:
    'The id of the event, as create-calendar-draft, list-calendar-view or get-calendar-event return it. The command runs only on an event that create-calendar-draft made, that has no attendees and the prefix "[DRAFT] ", and of which you are the organizer.',
  argumentHint: { kind: 'idOrName' },
};

const FIELD_OPTIONS: ReadonlyArray<CommandOptionMeta> = [
  { name: 'subject', key: 'subject', required: false, description: 'The subject. It gets the prefix "[DRAFT] " (not a second time if it starts with it).' },
  {
    name: 'start',
    key: 'start',
    required: false,
    description: 'The start: YYYY-MM-DDTHH:mm in the time zone of the mailbox, or YYYY-MM-DD for an all-day event.',
    argumentHint: { kind: 'iso8601' },
  },
  {
    name: 'end',
    key: 'end',
    required: false,
    description: 'The end, in the form of --start. The end of an all-day event is the day after its last day. It must be after the start.',
    argumentHint: { kind: 'iso8601' },
  },
  {
    name: 'all-day',
    key: 'allDay',
    required: false,
    description: '`true` for an all-day event, whose times are days. A change of it needs --start and --end.',
    argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
  },
  { name: 'location', key: 'location', required: false, description: 'The location, as text. An empty value removes it.' },
  { name: 'body', key: 'body', required: false, description: 'The text of the body, as plain text. It replaces the text of before and keeps the Teams link.' },
  {
    name: 'body-file',
    key: 'bodyFile',
    required: false,
    description:
      'A local UTF-8 text file with the body text, in place of --body, for a long body. At most 1000000 bytes. A byte order mark is removed. A relative path starts from the folder the CLI runs in. Mutually exclusive with --body.',
  },
  {
    name: 'planned-attendees',
    key: 'plannedAttendees',
    required: false,
    description:
      'The mail addresses of the planned attendees, separated by ";" or ",". They go into the body as text only ("Planned attendees (NOT invited): ..."): the event gets no attendees, so nobody gets an invitation. An empty value removes the line.',
  },
  {
    name: 'show-as',
    key: 'showAs',
    required: false,
    description: 'How the time shows in your calendar.',
    argumentHint: { kind: 'magicValue', values: ['free', 'tentative', 'busy', 'oof', 'workingElsewhere'] },
  },
  {
    name: 'categories',
    key: 'categories',
    required: false,
    description: 'The categories, separated by commas: they replace the categories of before. An empty value removes all categories.',
  },
  { name: 'reminder', key: 'reminder', required: false, description: 'The minutes of the reminder before the start, for example 15. `none` turns the reminder off.' },
];

const REQUIRED = new Set(['subject', 'start', 'end']);

const RESPONSE =
  '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the `eventId`, the `subject`, the `start` and the `end` of the event before the change (for a create: the subject, the times and the time zone that the run will give it), and `after` the same ids after it: a new event shows its new `eventId` there. An event that changed or got attendees between the plan and the write is in `skipped`; a write that Graph refused is in `failed`. With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes. No invitation is sent: the event has no attendees.';

describe('the docs of the calendar draft commands', () => {
  it('give create-calendar-draft its required fields and --teams', () => {
    expect(createMeta).toEqual({
      ...{
        summary:
          'Create a DRAFT event in your calendar: the subject gets the prefix "[DRAFT] ", and the event has NO attendees, so no invitation is sent. The planned attendees go into the body as text only. --teams true adds a Teams link. You send the invitation yourself from Outlook.',
        category: 'calendar',
        graphMethod: 'POST',
        graphPathTemplate: '/me/events',
        graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-post-events',
        example:
          '{write-bin} create-calendar-draft --subject "Budget review" --start 2026-11-20T14:00 --end 2026-11-20T15:00 --planned-attendees "a@example.com; b@example.com" --teams true --dry-run',
        bodyTemplate:
          "{ subject: '[DRAFT] {subject}', start: { dateTime: '{start}:00', timeZone: <time zone of the mailbox> }, end: { dateTime: '{end}:00', timeZone: <time zone of the mailbox> }, body: { contentType: 'text', content: '{body}' (or the text of {body-file}), then a line 'Planned attendees (NOT invited): {planned-attendees}' }, isAllDay?: {all-day}, location?: { displayName: '{location}' }, showAs?: '{show-as}', categories?: [{categories}], isReminderOn?: true, reminderMinutesBeforeStart?: {reminder}, isOnlineMeeting?: true, onlineMeetingProvider?: 'teamsForBusiness' (with {teams} true), singleValueExtendedProperties: [<the mark of this CLI>] }",
        effect: 'calendar-draft',
        undo: 'deletes the event, only if it has no attendees and the fields and the body text that the run set did not change',
        scopesRequired: ['Calendars.ReadWrite', 'MailboxSettings.Read'],
      },
      options: [
        ...FIELD_OPTIONS.map((option) => (REQUIRED.has(option.name) ? { ...option, required: true } : option)),
        {
          name: 'teams',
          key: 'teams',
          required: false,
          description: '`true` adds a Teams link to the event. With no attendees, Teams sends nobody an invitation.',
          argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
        },
      ],
      responseShape: RESPONSE,
    });
  });

  it('give update-calendar-draft the event and each field, none of them required', () => {
    expect(updateMeta).toEqual({
      ...{
        summary:
          'Change the fields of a DRAFT event that create-calendar-draft made: the subject (it keeps the prefix "[DRAFT] "), the times, the location, the body, the planned attendees (as text) and the other options. An event with attendees is refused before any write, so no update or invitation is sent.',
        category: 'calendar',
        graphMethod: 'PATCH',
        graphPathTemplate: '/me/events/{event-id}',
        graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-update',
        example: '{write-bin} update-calendar-draft --event-id "AAMkAD..." --start 2026-11-20T15:00 --end 2026-11-20T16:00 --dry-run',
        bodyTemplate:
          "{ subject?: '[DRAFT] {subject}', start?: { dateTime: '{start}:00', timeZone: <time zone of the mailbox> }, end?: { dateTime: '{end}:00', timeZone: <time zone of the mailbox> }, isAllDay?: {all-day}, location?: { displayName: '{location}' }, body?: { contentType: 'text', content: '{body}' (or the text of {body-file}), then a line 'Planned attendees (NOT invited): {planned-attendees}' } (HTML with the Teams block of before on an online meeting), showAs?: '{show-as}', categories?: [{categories}], isReminderOn?: true or false, reminderMinutesBeforeStart?: {reminder} }",
        effect: 'calendar-draft',
        undo: 'writes the fields of before back, only if the event has no attendees and the fields and the body text that the run set did not change',
        scopesRequired: ['Calendars.ReadWrite', 'MailboxSettings.Read'],
      },
      options: [EVENT_ID, ...FIELD_OPTIONS],
      responseShape: RESPONSE,
    });
  });

  it('put delete-calendar-draft on the delete class, with the limits of its undo', () => {
    expect(deleteMeta).toEqual({
      ...{
        summary:
          'Delete a DRAFT event that create-calendar-draft made. An event with attendees is refused before any write, so no cancellation is sent. The undo makes the event again from the run journal, with a new id and, for an online meeting, a new Teams link. Over MCP it runs on run-destructive-command.',
        category: 'calendar',
        graphMethod: 'DELETE',
        graphPathTemplate: '/me/events/{event-id}',
        graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-delete',
        example: '{write-bin} delete-calendar-draft --event-id "AAMkAD..." --dry-run',
        effect: 'delete',
        undo: 'makes the event again from the copy in the run journal: Graph gives it a new id, a new Teams link to an online meeting, and a plain-text body (the formatting of the body and the old Teams link are lost)',
        scopesRequired: ['Calendars.ReadWrite'],
      },
      options: [EVENT_ID],
      responseShape: RESPONSE,
    });
  });
});
