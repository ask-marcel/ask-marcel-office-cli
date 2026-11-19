import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { CALENDAR_ZONE, createCalendarStoreFake } from '../../test-helpers/calendar-store-fake.ts';
import { createLocalBytesFake, UTF8_BOM, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * create-calendar-draft makes a draft of a meeting: an event with the prefix
 * "[DRAFT] ", the hidden mark of this CLI, and no attendees, so Graph sends no
 * invitation. The planned attendees go into the body as text. The times are
 * in the time zone of the mailbox. The undo deletes the event if it did not
 * change.
 */

const MEETING = { subject: 'Budget review', start: '2026-11-20T14:00', end: '2026-11-20T15:00' };

const MARK = [{ id: 'String {6b1f3c2e-9d47-4e8a-a5c1-3f20d7b9e461} Name askMarcelCalendarDraft', value: 'create-calendar-draft' }];

describe('create-calendar-draft', () => {
  it('creates an event with the prefix, the mark, the planned attendees as text and no attendees, and the undo deletes it', async () => {
    const calendar = createCalendarStoreFake();
    const journal = createRunJournalFake();
    const params = { ...MEETING, plannedAttendees: 'a@example.com, b@example.com;', body: '  Agenda: the figures  ', location: 'Room 3' };
    const run = await runWrite(calendar.graph, 'create-calendar-draft', params, { journal });
    expect(run).toMatchObject({ done: 1, skipped: [], failed: [], runId: RUN_ID });
    expect(calendar.calls).toEqual(['GET /me/mailboxSettings/timeZone', 'POST /me/events']);
    const event = calendar.events.get('event-1');
    expect(event).toMatchObject({
      subject: '[DRAFT] Budget review',
      start: { dateTime: '2026-11-20T14:00:00.0000000', timeZone: 'UTC' },
      body: { contentType: 'text', content: 'Agenda: the figures\n\nPlanned attendees (NOT invited): a@example.com; b@example.com' },
      singleValueExtendedProperties: MARK,
      attendees: [],
    });
    expect(run['changes']).toEqual([
      {
        action: 'create-calendar-draft',
        target: '/me/events',
        before: { subject: '[DRAFT] Budget review', start: '2026-11-20T14:00', end: '2026-11-20T15:00', timeZone: CALENDAR_ZONE, plannedAttendees: 'a@example.com; b@example.com' },
        after: { eventId: 'event-1', subject: '[DRAFT] Budget review', start: '2026-11-20T14:00:00.0000000', end: '2026-11-20T15:00:00.0000000' },
      },
    ]);
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual(['undone']);
    expect(calendar.events.size).toBe(0);
  });

  it('sends the times in the time zone of the mailbox and only the fields given', async () => {
    const calendar = createCalendarStoreFake();
    const sent: unknown[] = [];
    const graph = { ...calendar.graph, post: async (path: string, body: unknown) => (sent.push(body), calendar.graph.post(path, body)) };
    await runWrite(graph, 'create-calendar-draft', MEETING);
    expect(sent).toEqual([
      {
        subject: '[DRAFT] Budget review',
        start: { dateTime: '2026-11-20T14:00:00', timeZone: CALENDAR_ZONE },
        end: { dateTime: '2026-11-20T15:00:00', timeZone: CALENDAR_ZONE },
        body: { contentType: 'text', content: '' },
        singleValueExtendedProperties: MARK,
      },
    ]);
  });

  it('shows the event on a dry-run, with --teams, and writes nothing', async () => {
    const calendar = createCalendarStoreFake();
    const dry = await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, teams: 'true' }, { dryRun: true });
    expect(dry).toEqual({
      dryRun: true,
      command: 'create-calendar-draft',
      changes: [
        {
          action: 'create-calendar-draft',
          target: '/me/events',
          before: { subject: '[DRAFT] Budget review', start: '2026-11-20T14:00', end: '2026-11-20T15:00', timeZone: CALENDAR_ZONE, teams: 'true' },
        },
      ],
    });
    expect(calendar.calls).toEqual(['GET /me/mailboxSettings/timeZone']);
  });

  it('gives the prefix once to a subject that has it in any case', async () => {
    const calendar = createCalendarStoreFake();
    await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, subject: ' [draft]   Budget review ' });
    expect(calendar.events.get('event-1')?.['subject']).toBe('[DRAFT] Budget review');
  });

  it('adds a Teams link with --teams true, and the undo still deletes the event that Exchange gave a join block', async () => {
    const calendar = createCalendarStoreFake();
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, teams: 'true', body: 'Agenda' }, { journal });
    expect(calendar.events.get('event-1')).toMatchObject({ isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' });
    expect(String((calendar.events.get('event-1')?.['body'] as { content: string }).content)).toContain('Microsoft Teams meeting');
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual(['undone']);
  });

  it('makes an all-day event of days, and sets the location, the categories, the show-as and the reminder', async () => {
    const calendar = createCalendarStoreFake();
    const params = { ...MEETING, start: '2026-11-20', end: '2026-11-22', allDay: 'true', categories: 'Red, Blue', showAs: 'oof', reminder: 'none', teams: 'false' };
    await runWrite(calendar.graph, 'create-calendar-draft', params);
    expect(calendar.events.get('event-1')).toMatchObject({
      start: { dateTime: '2026-11-20T00:00:00.0000000' },
      end: { dateTime: '2026-11-22T00:00:00.0000000' },
      isAllDay: true,
      categories: ['Red', 'Blue'],
      showAs: 'oof',
      isReminderOn: false,
      isOnlineMeeting: false,
    });
    await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, reminder: '30' });
    expect(calendar.events.get('event-2')).toMatchObject({ isReminderOn: true, reminderMinutesBeforeStart: 30 });
  });

  it('takes the body from --body-file, with no byte order mark, and refuses it with --body', async () => {
    const calendar = createCalendarStoreFake();
    const files = createLocalBytesFake({ 'agenda.txt': utf8Bytes('Agenda: é and ü', UTF8_BOM) });
    await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, bodyFile: 'agenda.txt', plannedAttendees: 'a@example.com' }, { files });
    expect(calendar.events.get('event-1')?.['body']).toEqual({ contentType: 'text', content: 'Agenda: é and ü\n\nPlanned attendees (NOT invited): a@example.com' });
    const both = await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, bodyFile: 'agenda.txt', body: 'x' }, { files });
    expect(both['error']).toMatchObject({
      code: 'body_file_conflict',
      message: '--body-file and --body are mutually exclusive: give the body text in one of them.',
    });
  });
});
