import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { CALENDAR_ZONE, createCalendarStoreFake, seededDraft, TEAMS_RULE } from '../../test-helpers/calendar-store-fake.ts';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * update-calendar-draft changes the fields of a calendar draft. It reads the
 * event in the time zone of the mailbox, writes only the fields given, keeps
 * the prefix, the planned attendees, the text and the Teams link that the
 * options do not change, and its undo writes the fields of before back if the
 * event did not change after the run.
 */

const BEFORE = { eventId: 'ev-a', subject: '[DRAFT] Budget review', start: '2026-11-20T14:00:00.0000000', end: '2026-11-20T15:00:00.0000000' };

const seeded = (): ReturnType<typeof createCalendarStoreFake> => createCalendarStoreFake([seededDraft('ev-a')]);

const bodyOf = (calendar: ReturnType<typeof createCalendarStoreFake>, id: string): string => String((calendar.events.get(id)?.['body'] as { content?: string }).content);

describe('update-calendar-draft', () => {
  it('moves the event in the time zone of the mailbox, and the undo moves it back', async () => {
    const calendar = seeded();
    const journal = createRunJournalFake();
    const run = await runWrite(
      calendar.graph,
      'update-calendar-draft',
      { eventId: 'ev-a', subject: 'Budget review 2', start: '2026-11-20T15:00', end: '2026-11-20T16:00' },
      { journal }
    );
    expect(run['changes']).toEqual([
      {
        action: 'update-calendar-draft',
        target: '/me/events/ev-a',
        before: BEFORE,
        after: { ...BEFORE, subject: '[DRAFT] Budget review 2', start: '2026-11-20T15:00:00.0000000', end: '2026-11-20T16:00:00.0000000', fields: 'subject, start, end' },
      },
    ]);
    expect(calendar.calls).toEqual(['GET /me/mailboxSettings/timeZone', 'GET /me/events/ev-a', 'GET /me/events/ev-a', 'PATCH /me/events/ev-a', 'GET /me/events/ev-a']);
    expect(calendar.prefers.slice(1, 3)).toEqual([`outlook.timezone="${CALENDAR_ZONE}", outlook.body-content-type="text"`, 'outlook.body-content-type="text"']);
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual(['undone']);
    expect(calendar.events.get('ev-a')).toMatchObject({ subject: '[DRAFT] Budget review', start: { dateTime: '2026-11-20T14:00:00.0000000' } });
  });

  it('shows the change on a dry-run and writes nothing', async () => {
    const calendar = seeded();
    const dry = await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', showAs: 'free' }, { dryRun: true });
    expect(dry).toEqual({ dryRun: true, command: 'update-calendar-draft', changes: [{ action: 'update-calendar-draft', target: '/me/events/ev-a', before: BEFORE }] });
    expect(calendar.calls).toEqual(['GET /me/mailboxSettings/timeZone', 'GET /me/events/ev-a']);
  });

  it('compares one new time with the other time of the event', async () => {
    const calendar = seeded();
    const early = await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', end: '2026-11-20T13:59' });
    expect(early['error']).toMatchObject({ code: 'calendar_end_before_start' });
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', end: '2026-11-20T16:30' });
    expect(calendar.events.get('ev-a')?.['end']).toEqual({ dateTime: '2026-11-20T16:30:00.0000000', timeZone: 'UTC' });
  });

  it('keeps the planned attendees with a new text, and the text with new planned attendees', async () => {
    const calendar = seeded();
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', body: 'New agenda' });
    expect(bodyOf(calendar, 'ev-a')).toBe('New agenda\n\nPlanned attendees (NOT invited): a@example.com; b@example.com');
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', plannedAttendees: 'c@example.com' });
    expect(bodyOf(calendar, 'ev-a')).toBe('New agenda\n\nPlanned attendees (NOT invited): c@example.com');
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', plannedAttendees: '' });
    expect(bodyOf(calendar, 'ev-a')).toBe('New agenda');
  });

  it('keeps the Teams link of an online meeting with a new body, and the undo gives the body of before back', async () => {
    const calendar = createCalendarStoreFake();
    await runWrite(calendar.graph, 'create-calendar-draft', { subject: 'Plan', start: '2026-11-20T14:00', end: '2026-11-20T15:00', teams: 'true', body: 'Old' });
    const old = bodyOf(calendar, 'event-1');
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'event-1', body: 'New <text> & more' }, { journal });
    expect(calendar.events.get('event-1')).toMatchObject({ isOnlineMeeting: true });
    expect(bodyOf(calendar, 'event-1')).toStartWith(`<html><body><p>New &lt;text&gt; &amp; more</p><div style="width:100%">${TEAMS_RULE}`);
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual(['undone']);
    expect(bodyOf(calendar, 'event-1')).toBe(old);
  });

  it('takes the text from --body-file', async () => {
    const calendar = seeded();
    const files = createLocalBytesFake({ 'agenda.txt': utf8Bytes('Long agenda') });
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', bodyFile: 'agenda.txt' }, { files });
    expect(bodyOf(calendar, 'ev-a')).toStartWith('Long agenda\n\nPlanned attendees');
  });

  it('makes an all-day event with both days, removes the location, and refuses --all-day with one time', async () => {
    const calendar = seeded();
    const one = await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', allDay: 'true', start: '2026-11-21' });
    expect(one['error']).toMatchObject({ code: 'calendar_all_day_needs_times', message: '--all-day changes the form of the times, so give --start and --end with it.' });
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', allDay: 'true', start: '2026-11-21', end: '2026-11-22', location: '' });
    expect(calendar.events.get('ev-a')).toMatchObject({ isAllDay: true, start: { dateTime: '2026-11-21T00:00:00.0000000' }, location: { displayName: '' } });
  });

  it('refuses a run with no field to change, before any write', async () => {
    const calendar = seeded();
    expect((await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a' }))['error']).toMatchObject({
      code: 'nothing_to_update',
      message: 'Give at least one field to change, for example --subject or --start.',
    });
    expect(calendar.calls).toEqual(['GET /me/mailboxSettings/timeZone', 'GET /me/events/ev-a']);
  });
});

describe('the undo of update-calendar-draft', () => {
  const updated = async (): Promise<{ calendar: ReturnType<typeof createCalendarStoreFake>; journal: ReturnType<typeof createRunJournalFake> }> => {
    const calendar = seeded();
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', subject: 'Moved', body: 'New agenda' }, { journal });
    return { calendar, journal };
  };

  it('restores the fields when another field changed after the run', async () => {
    const { calendar, journal } = await updated();
    calendar.edit('ev-a', { location: { displayName: 'Room 9' } });
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual(['undone']);
    expect(calendar.events.get('ev-a')).toMatchObject({ subject: '[DRAFT] Budget review', location: { displayName: 'Room 9' } });
  });

  it.each([
    ['a field that the run set', { subject: '[DRAFT] Moved again' }],
    ['the body text that the run set', { body: { contentType: 'text', content: 'Other agenda' } }],
    ['attendees', { attendees: [{ emailAddress: { address: 'a@example.com' } }] }],
  ])('passes over the event when %s changed after the run', async (_name, fields) => {
    const { calendar, journal } = await updated();
    calendar.edit('ev-a', fields);
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual([
      'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.',
    ]);
    expect(calendar.calls.filter((call) => call.startsWith('PATCH'))).toEqual(['PATCH /me/events/ev-a']);
  });
});
