import type { CommandOptionMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { BodyFileSpec } from './draft-body-file.ts';

/*
 * What the three calendar draft commands share: the option of the event, how
 * --body-file stands in for --body, and the words of their output.
 */

export const EVENT_ID_OPTION: CommandOptionMeta = {
  name: 'event-id',
  key: 'eventId',
  required: true,
  description:
    'The id of the event, as create-calendar-draft, list-calendar-view or get-calendar-event return it. The command runs only on an event that create-calendar-draft made, that has no attendees and the prefix "[DRAFT] ", and of which you are the organizer.',
  argumentHint: { kind: 'idOrName' },
};

/** --body-file gives the text of --body; a calendar draft can have an empty body. */
export const DRAFT_BODY_FILE: BodyFileSpec = { key: 'body', required: false, exclusive: ['body'] };

/** The words of the output that each calendar draft command shares. */
export const CALENDAR_DRAFT_RESPONSE =
  '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the `eventId`, the `subject`, the `start` and the `end` of the event before the change (for a create: the subject, the times and the time zone that the run will give it), and `after` the same ids after it: a new event shows its new `eventId` there. An event that changed or got attendees between the plan and the write is in `skipped`; a write that Graph refused is in `failed`. With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes. No invitation is sent: the event has no attendees.';
