import { describe, expect, it } from 'bun:test';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createCalendarStoreFake, seededDraft } from '../../test-helpers/calendar-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * Each Graph call of the calendar draft commands that fails, and each answer
 * that this CLI cannot read, stops the command with the error and makes no
 * later write: a run reports it, and an undo marks the change failed.
 */

const BUSY: GraphError = { type: 'api_error', status: 503, code: 'ServiceUnavailable', message: 'Try again later.' };

const MEETING = { subject: 'Plan', start: '2026-11-20T14:00', end: '2026-11-20T15:00' };

const unreadable = (what: string): Record<string, string> => ({ message: `Graph gave ${what} in a shape that this CLI cannot read.` });

const seeded = (): ReturnType<typeof createCalendarStoreFake> => createCalendarStoreFake([seededDraft('ev-a')]);

const FAILED_CALLS: ReadonlyArray<[string, Record<string, string>, string]> = [
  ['create-calendar-draft', MEETING, 'GET /me/mailboxSettings/timeZone'],
  ['create-calendar-draft', MEETING, 'POST /me/events'],
  ['update-calendar-draft', { eventId: 'ev-a', body: 'x' }, 'GET /me/mailboxSettings/timeZone'],
  ['update-calendar-draft', { eventId: 'ev-a', body: 'x' }, 'GET /me/events/ev-a'],
  ['update-calendar-draft', { eventId: 'ev-a', body: 'x' }, 'PATCH /me/events/ev-a'],
  ['delete-calendar-draft', { eventId: 'ev-a' }, 'GET /me/events/ev-a'],
  ['delete-calendar-draft', { eventId: 'ev-a' }, 'DELETE /me/events/ev-a'],
];

describe('a failed Graph call', () => {
  it.each(FAILED_CALLS)('stops %s with the error of %s', async (name, params, call) => {
    const calendar = seeded();
    calendar.failures.set(call, BUSY);
    expect((await runWrite(calendar.graph, name, params))['error']).toMatchObject({ code: 'ServiceUnavailable' });
    expect(calendar.calls.at(-1)).toBe(call);
  });

  it('stops an update when the read of the HTML body fails, after the read of the event', async () => {
    const calendar = seeded();
    calendar.hooks.afterRead = (count) => (count === 1 ? calendar.failures.set('GET /me/events/ev-a', BUSY) : undefined);
    expect((await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', body: 'x' }))['error']).toMatchObject({ code: 'ServiceUnavailable' });
  });

  it('fails the write when the read before it fails with another error than a missing event', async () => {
    const calendar = seeded();
    calendar.hooks.afterRead = (count) => (count === 1 ? calendar.failures.set('GET /me/events/ev-a', BUSY) : undefined);
    expect((await runWrite(calendar.graph, 'delete-calendar-draft', { eventId: 'ev-a' }))['error']).toMatchObject({ code: 'ServiceUnavailable' });
    expect(calendar.events.has('ev-a')).toBe(true);
  });
});

const UNREADABLE: ReadonlyArray<[string, Record<string, string>, string, unknown, Record<string, string>]> = [
  ['create-calendar-draft', MEETING, 'GET /me/mailboxSettings/timeZone', { value: '' }, unreadable('the time zone of the mailbox')],
  ['create-calendar-draft', MEETING, 'POST /me/events', { id: 'event 1', subject: 'x', start: { dateTime: 'a' }, end: { dateTime: 'b' } }, unreadable('the event it created')],
  ['create-calendar-draft', MEETING, 'POST /me/events', { id: 'event-1', subject: 'x', start: {}, end: { dateTime: 'b' } }, unreadable('the event it created')],
  ['update-calendar-draft', { eventId: 'ev-a', showAs: 'free' }, 'GET /me/events/ev-a', { id: 'ev-a', subject: 'x' }, unreadable('a calendar event')],
  ['update-calendar-draft', { eventId: 'ev-a', showAs: 'free' }, 'PATCH /me/events/ev-a', { id: 'ev-a' }, unreadable('the event it changed')],
  ['delete-calendar-draft', { eventId: 'ev-a' }, 'GET /me/events/ev-a', null, unreadable('a calendar event')],
];

describe('an answer that this CLI cannot read', () => {
  it.each(UNREADABLE)('stops %s at %s', async (name, params, call, answer, refusal) => {
    const calendar = seeded();
    calendar.answers.set(call, answer);
    expect((await runWrite(calendar.graph, name, params))['error']).toMatchObject(refusal);
  });

  it('stops an update when the HTML body has no content', async () => {
    const calendar = seeded();
    calendar.hooks.afterRead = (count) => (count === 1 ? calendar.answers.set('GET /me/events/ev-a', { body: {} }) : undefined);
    expect((await runWrite(calendar.graph, 'update-calendar-draft', { eventId: 'ev-a', body: 'x' }))['error']).toMatchObject(unreadable('the body of a calendar event'));
  });
});

describe('a failed read just after a write', () => {
  it('keeps the change in the run, and its undo says changed rather than delete the event', async () => {
    const calendar = createCalendarStoreFake();
    const journal = createRunJournalFake();
    calendar.failures.set('GET /me/events/event-1', BUSY);
    expect(await runWrite(calendar.graph, 'create-calendar-draft', MEETING, { journal })).toMatchObject({ done: 1, failed: [] });
    calendar.failures.clear();
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual([
      'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.',
    ]);
    expect(calendar.events.has('event-1')).toBe(true);
  });
});

describe('a failed call in an undo', () => {
  it('marks the undo failed when the read of its check fails', async () => {
    const calendar = createCalendarStoreFake();
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'create-calendar-draft', MEETING, { journal });
    calendar.failures.set('GET /me/events/event-1', BUSY);
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual(['failed: Try again later.']);
    expect(calendar.events.has('event-1')).toBe(true);
  });

  it.each([
    ['fails', BUSY, undefined, 'failed: Try again later.'],
    ['gives no id', undefined, { subject: 'x' }, 'failed: Graph gave the event it made again in a shape that this CLI cannot read.'],
    ['gives an id that is not a Graph id', undefined, { id: 'a/b' }, 'failed: Graph gave the event it made again in a shape that this CLI cannot read.'],
  ])('marks the undo of a delete failed when the create %s', async (_name, failure, answer, status) => {
    const calendar = seeded();
    const journal = createRunJournalFake();
    await runWrite(calendar.graph, 'delete-calendar-draft', { eventId: 'ev-a' }, { journal });
    if (failure !== undefined) calendar.failures.set('POST /me/events', failure);
    if (answer !== undefined) calendar.answers.set('POST /me/events', answer);
    expect(statusesOf(await undoWrite(calendar.graph, journal, true))).toEqual([status]);
  });
});
