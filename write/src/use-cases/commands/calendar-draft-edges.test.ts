import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createCalendarStoreFake, seededDraft } from '../../test-helpers/calendar-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The edges of the calendar drafts: the subject, the body text around the
 * planned attendees and the Teams join block, the times at their limits, the
 * read just before the write, and the exact reads that the commands send.
 */

const MEETING = { subject: 'Budget review', start: '2026-11-20T14:00', end: '2026-11-20T15:00' };

const MARK_ID = 'String {6b1f3c2e-9d47-4e8a-a5c1-3f20d7b9e461} Name askMarcelCalendarDraft';

const MARK_FILTER = encodeURI(`singleValueExtendedProperties($filter=id eq '${MARK_ID}')`);

type Calendar = ReturnType<typeof createCalendarStoreFake>;

const bodyOf = (calendar: Calendar, id: string): unknown => calendar.events.get(id)?.['body'];

const online = (content: string): Record<string, unknown> => ({ isOnlineMeeting: true, body: { contentType: 'html', content } });

describe('the subject and the body of a calendar draft', () => {
  it('keeps a "[DRAFT]" inside the subject, and trims an updated subject', async () => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a')]);
    await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, subject: 'Notes on [DRAFT] plans' });
    expect(calendar.events.get('event-1')?.['subject']).toBe('[DRAFT] Notes on [DRAFT] plans');
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', subject: '  Moved  ' });
    expect(calendar.events.get('ev-a')?.['subject']).toBe('[DRAFT] Moved');
    expect((await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', subject: '  ' }))['error']).toMatchObject({ message: '--subject is empty' });
  });

  it('keeps the text before the planned attendees, an underscore in it, and only the line of the mark', async () => {
    const content = 'Read file_a first\n\nPlanned attendees (NOT invited): a@example.com\nBring the slides';
    const calendar = createCalendarStoreFake([seededDraft('ev-a', { body: { contentType: 'text', content } })]);
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', plannedAttendees: 'b@example.com' });
    expect(bodyOf(calendar, 'ev-a')).toEqual({ contentType: 'text', content: 'Read file_a first\n\nPlanned attendees (NOT invited): b@example.com' });
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', body: 'New' });
    expect(bodyOf(calendar, 'ev-a')).toEqual({ contentType: 'text', content: 'New\n\nPlanned attendees (NOT invited): b@example.com' });
  });

  it.each([
    [
      'from its title when it has no line of underscores',
      '<html><body><p>Old</p>Microsoft Teams meeting<br>Join: x</body></html>',
      'Microsoft Teams meeting<br>Join: x</body></html>',
    ],
    ['as no block when the line of underscores is not in a div', '<html><body><p>Old</p>________________________________</body></html>', undefined],
    ['as no block when the body has none', '<html><body><p>Old</p></body></html>', undefined],
  ])('takes the Teams block of an online meeting %s', async (_name, html, block) => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a', online(html))]);
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', body: 'a\nb' });
    const sent = block === undefined ? undefined : `<html><body><p>a<br>b</p>${block}`;
    expect(bodyOf(calendar, 'ev-a')).toEqual(sent === undefined ? { contentType: 'text', content: 'a\nb' } : { contentType: 'html', content: sent });
  });
});

describe('the Teams join block and the text of the user', () => {
  it('keeps each word of the text of an event that is not an online meeting', async () => {
    const text = 'Agenda for the Microsoft Teams meeting of next week\n____________\nNotes';
    const calendar = createCalendarStoreFake([seededDraft('ev-a', { body: { contentType: 'text', content: `${text}\n\nPlanned attendees (NOT invited): a@example.com` } })]);
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', plannedAttendees: 'b@example.com' }, { journal });
    expect(bodyOf(calendar, 'ev-a')).toEqual({ contentType: 'text', content: `${text}\n\nPlanned attendees (NOT invited): b@example.com` });
    expect(statusesOf(await undoWrite(calendar.graph, journal, false))).toEqual(['will-undo']);
  });

  it.each([
    ['the last title with its line of ten underscores', '_'.repeat(10), 'Microsoft Teams meeting<br>Join: x'],
    ['the last title alone after nine underscores, which stay in the text', '_'.repeat(9), 'Microsoft Teams meeting<br>Join: x'],
  ])('cuts the block of an online meeting at %s', async (_name, rule, block) => {
    const own = 'See the Microsoft Teams meeting notes';
    const html = `<html><body><p>${own}</p><div>${rule}</div><div>${block}</div></body></html>`;
    const calendar = createCalendarStoreFake([seededDraft('ev-a', online(html))]);
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', plannedAttendees: 'b@example.com' });
    const kept = rule.length === 10 ? own : `${own}<br>${rule}`;
    const tail = rule.length === 10 ? `<div>${rule}</div><div>${block}</div></body></html>` : `${block}</div></body></html>`;
    expect(bodyOf(calendar, 'ev-a')).toEqual({ contentType: 'html', content: `<html><body><p>${kept}<br><br>Planned attendees (NOT invited): b@example.com</p>${tail}` });
  });
});

describe('an online meeting with no text of the user', () => {
  it('gets the planned attendees before its whole Teams block', async () => {
    const block = `<div>${'_'.repeat(10)}</div><div>Microsoft Teams meeting<br>Join: x</div></body></html>`;
    const calendar = createCalendarStoreFake([seededDraft('ev-a', online(`<html><body>${block}`))]);
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', plannedAttendees: 'b@example.com' });
    expect(bodyOf(calendar, 'ev-a')).toEqual({ contentType: 'html', content: `<html><body><p>Planned attendees (NOT invited): b@example.com</p>${block}` });
  });
});

describe('the times of a calendar draft at their limits', () => {
  it('refuses a new start at the end of the event, and an all-day change with the end alone', async () => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a')]);
    expect((await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', start: '2026-11-20T15:00' }))['error']).toMatchObject({ code: 'calendar_end_before_start' });
    expect((await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', allDay: 'false', end: '2026-11-20T16:00' }))['error']).toMatchObject({
      code: 'calendar_all_day_needs_times',
    });
  });

  it('changes another field of an event of no time, and makes an event that is not all-day with --all-day false', async () => {
    const at = { dateTime: '2026-11-20T14:00:00.0000000', timeZone: 'UTC' };
    const calendar = createCalendarStoreFake([seededDraft('ev-a', { start: at, end: at })]);
    expect(await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', showAs: 'free' })).toMatchObject({ done: 1 });
    await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, allDay: 'false' });
    expect(calendar.events.get('event-1')?.['isAllDay']).toBe(false);
  });

  it('refuses a reminder of six digits', async () => {
    const calendar = createCalendarStoreFake();
    expect((await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, reminder: '123456' }))['error']).toMatchObject({
      message: '--reminder must be the minutes before the start, a whole number such as 15, or none to turn the reminder off',
    });
  });
});

describe('the read of a calendar draft', () => {
  it('asks for the fields, the mark and the HTML body with these exact queries', async () => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a')]);
    await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', body: 'x' }, { dryRun: true });
    expect(calendar.paths).toEqual([
      '/me/mailboxSettings/timeZone',
      `/me/events/ev-a?$select=id,subject,start,end,isAllDay,location,showAs,categories,isReminderOn,reminderMinutesBeforeStart,isOnlineMeeting,onlineMeetingProvider,attendees,isOrganizer,type,body,lastModifiedDateTime&$expand=${MARK_FILTER}`,
      '/me/events/ev-a?$select=body',
    ]);
  });

  it.each([
    ['loses its mark', { singleValueExtendedProperties: undefined }],
    ['gets another organizer', { isOrganizer: false }],
    ['becomes a part of a series', { type: 'occurrence' }],
    ['loses its prefix', { subject: 'Budget review' }],
  ])('passes over an event that %s after the plan', async (_name, fields) => {
    const calendar = createCalendarStoreFake([seededDraft('ev-a')]);
    calendar.hooks.afterRead = (count) => (count === 1 ? calendar.edit('ev-a', fields) : undefined);
    expect(await runWrite(calendar.graph, 'delete-calendar-draft', { eventId: 'ev-a' })).toMatchObject({ done: 0, failed: [], skipped: [{ target: '/me/events/ev-a' }] });
    expect(calendar.events.has('ev-a')).toBe(true);
  });
});

describe('the undo of the calendar drafts at their edges', () => {
  it('deletes an online meeting after Exchange or Outlook changed only its Teams link', async () => {
    const calendar = createCalendarStoreFake();
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, teams: 'true', body: 'Agenda' }, { journal });
    calendar.edit('event-1', { ...online('<html><body><p>Agenda</p>Microsoft Teams meeting<br>Join: y</body></html>'), onlineMeetingProvider: 'unknown' });
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual(['undone']);
  });

  it('keeps an online meeting whose Teams link was turned off after the run', async () => {
    const calendar = createCalendarStoreFake();
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'create-calendar-draft', { ...MEETING, teams: 'true', body: 'Agenda' }, { journal });
    calendar.edit('event-1', { isOnlineMeeting: false });
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual([
      'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.',
    ]);
  });

  it('keeps the whole text of a deleted event that is not an online meeting', async () => {
    const content = 'Agenda for the Microsoft Teams meeting of next week\n____________\nNotes';
    const calendar = createCalendarStoreFake([seededDraft('ev-a', { body: { contentType: 'text', content } })]);
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'delete-calendar-draft', { eventId: 'ev-a' }, { journal });
    await undoWrite(calendar.graph, journal, true);
    expect(calendar.events.get('event-1')?.['body']).toEqual({ contentType: 'text', content });
  });

  it('makes a deleted event again with each of its fields, a trimmed text and an empty location', async () => {
    const fields = {
      location: null,
      showAs: 'free',
      isReminderOn: true,
      reminderMinutesBeforeStart: 5,
      categories: ['Red'],
      body: { contentType: 'text', content: '\n\n  Agenda  \n' },
    };
    const calendar = createCalendarStoreFake([seededDraft('ev-a', fields)]);
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'delete-calendar-draft', { eventId: 'ev-a' }, { journal });
    await undoWrite(calendar.graph, journal, true);
    expect(calendar.events.get('event-1')).toMatchObject({
      subject: '[DRAFT] Budget review',
      start: { dateTime: '2026-11-20T14:00:00.0000000' },
      end: { dateTime: '2026-11-20T15:00:00.0000000' },
      isAllDay: false,
      location: { displayName: '' },
      showAs: 'free',
      reminderMinutesBeforeStart: 5,
      categories: ['Red'],
      body: { contentType: 'text', content: 'Agenda' },
    });
  });
});
