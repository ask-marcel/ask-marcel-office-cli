import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createCalendarStoreFake } from '../../test-helpers/calendar-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The checks of create-calendar-draft, before any write: a time that is not a
 * day of the calendar or not in the form of the event, and an end that is not
 * after the start, are refused with no exception. Then the undo of a create:
 * it deletes the event only if the fields and the body text of the run are
 * still there and the event has no attendees.
 */

const MEETING = { subject: 'Budget review', start: '2026-11-20T14:00', end: '2026-11-20T15:00' };

const TIMED = 'in the form YYYY-MM-DDTHH:mm, for example 2026-11-20T14:00, in the time zone of the mailbox.';
const ALL_DAY = 'in the form YYYY-MM-DD, for example 2026-11-20 (an all-day event), in the time zone of the mailbox.';
const ORDER = { code: 'calendar_end_before_start', message: 'The end of the event must be after its start. The end of an all-day event is the day after its last day.' };

const REFUSALS: ReadonlyArray<[string, Record<string, string>, Record<string, string>]> = [
  ['a start that is not a day of the calendar', { start: '2026-02-30T14:00' }, { code: 'invalid_event_time', message: `--start must be a day of the calendar ${TIMED}` }],
  ['an end with an hour 24', { end: '2026-11-20T24:00' }, { code: 'invalid_event_time', message: `--end must be a day of the calendar ${TIMED}` }],
  ['a day with no time', { start: '2026-11-20' }, { code: 'invalid_event_time', message: `--start must be a day of the calendar ${TIMED}` }],
  ['a time on an all-day event', { allDay: 'true' }, { code: 'invalid_event_time', message: `--start must be a day of the calendar ${ALL_DAY}` }],
  [
    'an all-day end that is not a day',
    { allDay: 'true', start: '2026-11-20', end: '2026-13-01' },
    { code: 'invalid_event_time', message: `--end must be a day of the calendar ${ALL_DAY}` },
  ],
  ['an end equal to the start', { end: '2026-11-20T14:00' }, ORDER],
  ['an end before the start', { end: '2026-11-19T15:00' }, ORDER],
  ['an all-day event of no day', { allDay: 'true', start: '2026-11-20', end: '2026-11-20' }, ORDER],
  [
    'a reminder that is not minutes',
    { reminder: '15m' },
    { message: '--reminder must be the minutes before the start, a whole number such as 15, or none to turn the reminder off' },
  ],
  ['a show-as that Graph does not take', { showAs: 'away' }, { message: '--show-as Invalid option: expected one of "free"|"tentative"|"busy"|"oof"|"workingElsewhere"' }],
];

describe('create-calendar-draft refuses, before any write', () => {
  it.each(REFUSALS)('%s', async (_name, change, refusal) => {
    const calendar = createCalendarStoreFake();
    const run = await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, ...change });
    expect(run['error']).toMatchObject(refusal);
    expect(calendar.calls.filter((call) => !call.startsWith('GET'))).toEqual([]);
  });

  it('a subject of white space only, and no start', async () => {
    const calendar = createCalendarStoreFake();
    const noStart = { subject: MEETING.subject, end: MEETING.end };
    expect((await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, subject: '   ' }))['error']).toMatchObject({ message: '--subject is empty' });
    expect((await runWrite(calendar.graph, 'create-calendar-draft', noStart))['error']).toMatchObject({ message: '--start is missing' });
    expect(calendar.calls).toEqual([]);
  });
});

const created = async (): Promise<{ calendar: ReturnType<typeof createCalendarStoreFake>; journal: ReturnType<typeof createRunJournalFake> }> => {
  const calendar = createCalendarStoreFake();
  const journal = createRunJournalFake();
  await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, body: 'Agenda', location: 'Room 3', reminder: '10' }, { journal });
  return { calendar, journal };
};

const CHANGED = 'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.';

describe('the undo of create-calendar-draft', () => {
  it('shows the delete on a dry-run of the undo, and deletes nothing', async () => {
    const { calendar, journal } = await created();
    expect(statusesOf(await undoWrite(calendar.graph, journal, false))).toEqual(['will-undo']);
    expect(calendar.events.has('event-1')).toBe(true);
  });

  it.each([
    ['the subject', { subject: '[DRAFT] Budget review, moved' }],
    ['the start', { start: { dateTime: '2026-11-20T15:00:00.0000000', timeZone: 'UTC' } }],
    ['the location', { location: { displayName: 'Room 4' } }],
    ['the reminder', { reminderMinutesBeforeStart: 30 }],
    ['the body text', { body: { contentType: 'text', content: 'Agenda, and more' } }],
    ['attendees', { attendees: [{ emailAddress: { address: 'a@example.com' } }] }],
    ['no prefix', { subject: 'Budget review' }],
  ])('keeps the event when %s changed after the run', async (_name, fields) => {
    const { calendar, journal } = await created();
    calendar.edit('event-1', fields);
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual([CHANGED]);
    expect(calendar.calls.filter((call) => call.startsWith('DELETE'))).toEqual([]);
  });

  it('deletes the event when only Graph changed the location object, not its name', async () => {
    const { calendar, journal } = await created();
    calendar.edit('event-1', { location: { displayName: 'Room 3', locationType: 'conferenceRoom' } });
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual(['undone']);
  });

  it('says missing when the event is gone', async () => {
    const { calendar, journal } = await created();
    calendar.events.delete('event-1');
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual(['missing: The item is no longer there, so there is nothing to undo.']);
  });
});
