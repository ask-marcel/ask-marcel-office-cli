import { describe, expect, it } from 'bun:test';
import { createCalendarStoreFake, seededDraft } from '../../test-helpers/calendar-store-fake.ts';
import { runWrite } from '../../test-helpers/run-write.ts';

/*
 * No invitation and no cancellation, ever. update-calendar-draft and
 * delete-calendar-draft read the event first and refuse, before any write, an
 * event that this CLI did not make (no mark), an event with attendees, an
 * event of another organizer, a part of a series, and an event with no prefix
 * "[DRAFT] ". The read again just before the write passes over an event that
 * got attendees after the plan. No request of the three commands has
 * attendees.
 */

const ATTENDEE = { emailAddress: { address: 'a@example.com' }, type: 'required' };

const COMMANDS: ReadonlyArray<[string, Record<string, string>]> = [
  ['update-calendar-draft', { eventId: 'ev-a', subject: 'New' }],
  ['delete-calendar-draft', { eventId: 'ev-a' }],
];

const WHY = 'so this command does not change it.';

const RULES: ReadonlyArray<[string, Record<string, unknown>, string, string]> = [
  [
    'no mark',
    { singleValueExtendedProperties: undefined },
    'not_a_calendar_draft',
    `--event-id: create-calendar-draft did not create this event (it has no mark of this CLI), ${WHY}`,
  ],
  [
    'a mark of another value',
    { singleValueExtendedProperties: [{ id: 'String {6b1f3c2e-9d47-4e8a-a5c1-3f20d7b9e461} Name askMarcelCalendarDraft', value: 'x' }] },
    'not_a_calendar_draft',
    `--event-id: create-calendar-draft did not create this event (it has no mark of this CLI), ${WHY}`,
  ],
  [
    'another property',
    { singleValueExtendedProperties: [{ id: 'String {00000000-0000-0000-0000-000000000000} Name other', value: 'create-calendar-draft' }] },
    'not_a_calendar_draft',
    `--event-id: create-calendar-draft did not create this event (it has no mark of this CLI), ${WHY}`,
  ],
  [
    'one attendee',
    { attendees: [ATTENDEE] },
    'calendar_draft_has_attendees',
    `--event-id: the event has 1 attendee(s). A change can send them an update and a delete a cancellation, ${WHY} Change it in Outlook.`,
  ],
  ['another organizer', { isOrganizer: false }, 'not_calendar_organizer', `--event-id: you are not the organizer of this event, ${WHY}`],
  ['a series', { type: 'seriesMaster' }, 'calendar_series', `--event-id: the event is a part of a series (seriesMaster). create-calendar-draft makes no series, ${WHY}`],
  ['an occurrence', { type: 'occurrence' }, 'calendar_series', `--event-id: the event is a part of a series (occurrence). create-calendar-draft makes no series, ${WHY}`],
  ['an exception', { type: 'exception' }, 'calendar_series', `--event-id: the event is a part of a series (exception). create-calendar-draft makes no series, ${WHY}`],
  [
    'no prefix',
    { subject: 'Budget review' },
    'calendar_draft_prefix_missing',
    `--event-id: the subject does not start with "[DRAFT] ", so the event is no longer a draft and ${WHY.replace('so ', '')}`,
  ],
  [
    'a prefix in lower case',
    { subject: '[draft] Budget review' },
    'calendar_draft_prefix_missing',
    `--event-id: the subject does not start with "[DRAFT] ", so the event is no longer a draft and ${WHY.replace('so ', '')}`,
  ],
];

const writes = (calls: ReadonlyArray<string>): ReadonlyArray<string> => calls.filter((call) => !call.startsWith('GET'));

describe.each(COMMANDS)('%s refuses, before any write, an event with', (name, params) => {
  it.each(RULES)('%s', async (_rule, fields, code, message) => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a', fields)]);
    expect((await runWrite(calendar.graph, name, params))['error']).toMatchObject({ code, message });
    expect(writes(calendar.calls)).toEqual([]);
  });

  it('two attendees, with their count', async () => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a', { attendees: [ATTENDEE, ATTENDEE] })]);
    expect((await runWrite(calendar.graph, name, params, { dryRun: true }))['error']).toMatchObject({ message: expect.stringContaining('the event has 2 attendee(s).') });
  });

  it('an id that is not a Graph id, with no request', async () => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a')]);
    expect((await runWrite(calendar.graph, name, { ...params, eventId: 'ev-a/../x' }))['error']).toMatchObject({
      code: 'invalid_id',
      message: expect.stringContaining('--event-id: "ev-a/../x"'),
    });
    expect(calendar.calls).toEqual([]);
  });

  it('attendees given after the plan: the read before the write passes over it', async () => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a')]);
    calendar.hooks.afterRead = (count) => (count === 1 ? calendar.edit('ev-a', { attendees: [ATTENDEE] }) : undefined);
    const run = await runWrite(calendar.graph, name, params);
    expect(run).toMatchObject({
      done: 0,
      failed: [],
      skipped: [{ target: '/me/events/ev-a', reason: 'The item changed after the plan, so this run passes over it. Run the command again to see it as it is now.' }],
    });
    expect(writes(calendar.calls)).toEqual([]);
  });

  it('a change after the plan, or a delete: passed over as changed or gone', async () => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a')]);
    calendar.hooks.afterRead = (count) => (count === 1 ? calendar.edit('ev-a', { showAs: 'free' }) : undefined);
    expect((await runWrite(calendar.graph, name, params))['skipped']).toMatchObject([{ reason: expect.stringContaining('changed after the plan') }]);
    calendar.hooks.afterRead = (count) => (count === 3 ? calendar.events.delete('ev-a') : undefined);
    expect((await runWrite(calendar.graph, name, params))['skipped']).toMatchObject([{ reason: 'The item is no longer there: it was moved or deleted after the plan.' }]);
    expect(writes(calendar.calls)).toEqual([]);
  });
});

describe('the requests of the calendar draft commands', () => {
  it('never have attendees, in a full cycle of create, update and delete', async () => {
    const calendar = createCalendarStoreFake();
    const bodies: unknown[] = [];
    const graph = {
      ...calendar.graph,
      post: async (path: string, body: unknown) => (bodies.push(body), calendar.graph.post(path, body)),
      patch: async (path: string, body: unknown) => (bodies.push(body), calendar.graph.patch(path, body)),
    };
    const meeting = { subject: 'Plan', start: '2026-11-20T14:00', end: '2026-11-20T15:00', plannedAttendees: 'a@example.com', teams: 'true' };
    await runWrite(graph, 'create-calendar-draft', meeting);
    await runWrite(graph, 'update-calendar-draft', { eventId: 'event-1', plannedAttendees: 'b@example.com', body: 'x' });
    await runWrite(graph, 'delete-calendar-draft', { eventId: 'event-1' });
    expect(bodies).toHaveLength(2);
    expect(JSON.stringify(bodies).toLowerCase()).not.toContain('"attendees"');
    expect(calendar.calls.filter((call) => !call.startsWith('GET'))).toEqual(['POST /me/events', 'PATCH /me/events/event-1', 'DELETE /me/events/event-1']);
  });
});
