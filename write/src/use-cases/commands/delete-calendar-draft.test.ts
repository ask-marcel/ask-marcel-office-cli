import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createCalendarStoreFake, seededDraft } from '../../test-helpers/calendar-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * delete-calendar-draft deletes a calendar draft, and the run journal keeps a
 * copy of it. The undo makes the event again from the copy: Graph gives it a
 * new id, the mark of this CLI, a plain-text body (the text of the user and
 * the planned attendees), and a new Teams link to an online meeting.
 */

const MARK = [{ id: 'String {6b1f3c2e-9d47-4e8a-a5c1-3f20d7b9e461} Name askMarcelCalendarDraft', value: 'create-calendar-draft' }];

const BEFORE = { eventId: 'ev-a', subject: '[DRAFT] Budget review', start: '2026-11-20T14:00:00.0000000', end: '2026-11-20T15:00:00.0000000' };

describe('delete-calendar-draft', () => {
  it('deletes the event, and the undo makes it again with a new id, the mark and its fields', async () => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a', { location: { displayName: 'Room 3', uniqueId: 'r3' }, categories: ['Red'], isReminderOn: false })]);
    const journal = createRunJournalFake();
    const run = await runWrite(calendar.graph, 'delete-calendar-draft', { eventId: 'ev-a' }, { journal });
    expect(run).toEqual({ done: 1, skipped: [], failed: [], runId: RUN_ID, changes: [{ action: 'delete-calendar-draft', target: '/me/events/ev-a', before: BEFORE, after: {} }] });
    expect(calendar.events.has('ev-a')).toBe(false);
    const undo = await undoWrite(calendar.graph, journal, true);
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', { eventId: 'event-1' }]]);
    expect(calendar.events.get('event-1')).toMatchObject({
      subject: '[DRAFT] Budget review',
      start: { dateTime: '2026-11-20T14:00:00.0000000', timeZone: 'UTC' },
      location: { displayName: 'Room 3' },
      categories: ['Red'],
      isReminderOn: false,
      isOnlineMeeting: false,
      body: { contentType: 'text', content: 'Agenda: the figures\n\nPlanned attendees (NOT invited): a@example.com; b@example.com' },
      singleValueExtendedProperties: MARK,
      attendees: [],
    });
  });

  it('shows the delete on a dry-run and deletes nothing', async () => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a')]);
    const dry = await runWrite(calendar.graph, 'delete-calendar-draft', { eventId: 'ev-a' }, { dryRun: true });
    expect(dry).toEqual({ dryRun: true, command: 'delete-calendar-draft', changes: [{ action: 'delete-calendar-draft', target: '/me/events/ev-a', before: BEFORE }] });
    expect(calendar.calls).toEqual(['GET /me/events/ev-a']);
  });

  it('makes an online meeting again with a new Teams link and without the old join block', async () => {
    const calendar = createCalendarStoreFake();
    await runWrite(calendar.graph, 'create-calendar-draft', { subject: 'Plan', start: '2026-11-20T14:00', end: '2026-11-20T15:00', teams: 'true', body: 'Agenda' });
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'delete-calendar-draft', { eventId: 'event-1' }, { journal });
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual(['undone']);
    const again = calendar.events.get('event-2');
    expect(again).toMatchObject({ isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' });
    expect(String((again?.['body'] as { content: string }).content)).toContain('Join: https://teams.example/join/event-2');
    expect(String((again?.['body'] as { content: string }).content)).not.toContain('join/event-1');
  });

  it('refuses a journal whose copy has a field that this CLI does not write, before any request', async () => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a')]);
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'delete-calendar-draft', { eventId: 'ev-a' }, { journal });
    const record = journal.records.get(RUN_ID) as { changes: Array<{ undo: { data: { copy: Record<string, unknown> } } }> };
    const change = record.changes[0];
    if (change !== undefined) change.undo.data.copy['attendees'] = [{ emailAddress: { address: 'a@example.com' } }];
    journal.records.set(RUN_ID, record);
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual([
      'failed: The run journal names an id that is not a Graph id, or data that this CLI did not write, so this undo does not run.',
    ]);
    expect(calendar.calls.filter((call) => call.startsWith('POST'))).toEqual([]);
  });
});
