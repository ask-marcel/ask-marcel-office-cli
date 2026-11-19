import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeWriteGraph } from '../../test-helpers/write-graph-fake.ts';
import type { WriteGraph } from '../ports/graph.ts';
import { writeGraphOf } from './command-graph.ts';

/*
 * No command sends mail (product rule), and a calendar event can send mail
 * too: `cancel` sends a cancellation to the attendees, `accept`, `decline` and
 * `tentativelyAccept` send an answer to the organizer, and `forward` sends the
 * invitation. The write view refuses a change to each of these actions before
 * any request, in any case and with the namespace. A second guard refuses an
 * event write whose body has attendees: Graph sends an invitation to each
 * attendee of an event that it creates or changes. An empty list of
 * attendees sends nothing, and a body with attendees to a path that is not an
 * event is not the business of this guard.
 */

type Sent = { method: string; path: string };

const recordingGraph = (sent: Sent[]): WriteGraph =>
  fakeWriteGraph({
    post: async (path) => {
      sent.push({ method: 'POST', path });
      return ok({ id: 'event-1' });
    },
    patch: async (path) => {
      sent.push({ method: 'PATCH', path });
      return ok({ id: 'event-1' });
    },
  });

const CALENDAR_SENDING_PATHS = [
  '/me/events/e1/cancel',
  '/me/events/e1/accept',
  '/me/events/e1/decline',
  '/me/events/e1/tentativelyAccept',
  '/me/calendar/events/e1/TentativelyAccept',
  '/me/events/e1/microsoft.graph.cancel',
  '/me/events/e1/Decline/',
  '/me/events/e1/forward',
];

describe('a write command that asks to answer, cancel or forward a calendar event', () => {
  it.each(CALENDAR_SENDING_PATHS)('is refused before any request when it POSTs to %s', async (path) => {
    const sent: Sent[] = [];
    const result = await writeGraphOf(recordingGraph(sent)).post(path, { comment: 'x' });
    expect(sent).toEqual([]);
    expect(result).toEqual({
      ok: false,
      error: {
        type: 'validation_error',
        code: 'send_refused',
        message: `POST ${path} was not sent: it would send mail, and this CLI never sends mail. A draft stays in Drafts until you send it yourself from Outlook.`,
      },
    });
  });
});

const ATTENDEE = { emailAddress: { address: 'someone@example.com' }, type: 'required' };

const INVITING_WRITES: ReadonlyArray<[string, string, Record<string, unknown>]> = [
  ['POST', '/me/events', { subject: 'Plan', attendees: [ATTENDEE] }],
  ['POST', '/me/calendar/events', { attendees: [ATTENDEE] }],
  ['POST', '/me/calendars/c1/events', { attendees: [ATTENDEE] }],
  ['POST', '/users/u1/events?$select=id', { Attendees: [ATTENDEE] }],
  ['PATCH', '/me/events/e1', { attendees: [ATTENDEE] }],
  ['PATCH', '/me/calendar/events/e1/', { ATTENDEES: [ATTENDEE, ATTENDEE] }],
  ['POST', '/me/Events', { attendees: [ATTENDEE] }],
];

const writeOf = async (view: WriteGraph, method: string, path: string, body: Record<string, unknown>): Promise<unknown> =>
  method === 'POST' ? view.post(path, body) : view.patch(path, body);

describe('a write command that gives a calendar event attendees', () => {
  it.each(INVITING_WRITES)('is refused before any request on %s %s', async (method, path, body) => {
    const sent: Sent[] = [];
    const result = await writeOf(writeGraphOf(recordingGraph(sent)), method, path, body);
    expect(sent).toEqual([]);
    expect(result).toEqual({
      ok: false,
      error: {
        type: 'validation_error',
        code: 'invitation_refused',
        message: `${method} ${path} was not sent: the event has attendees, so Graph would send them an invitation, and this CLI never sends one. Give the planned attendees as text in the body.`,
      },
    });
  });

  it('is sent when the event has no attendees, an empty list of attendees, or a body that is not an object', async () => {
    const sent: Sent[] = [];
    const view = writeGraphOf(recordingGraph(sent));
    await view.post('/me/events', { subject: 'Plan', body: { content: 'attendees: none' } });
    await view.patch('/me/events/e1', { attendees: [] });
    await view.patch('/me/events/e1', { attendees: 'none' });
    await view.post('/me/events', null);
    await view.post('/me/events', 'attendees');
    expect(sent).toEqual([
      { method: 'POST', path: '/me/events' },
      { method: 'PATCH', path: '/me/events/e1' },
      { method: 'PATCH', path: '/me/events/e1' },
      { method: 'POST', path: '/me/events' },
      { method: 'POST', path: '/me/events' },
    ]);
  });

  it('is sent when the path is not an event, whatever its body has', async () => {
    const sent: Sent[] = [];
    const view = writeGraphOf(recordingGraph(sent));
    await view.post('/me/messages', { attendees: [ATTENDEE] });
    await view.post('/me/eventsArchive', { attendees: [ATTENDEE] });
    await view.patch('/me/todo/lists/l1/tasks/t1', { attendees: [ATTENDEE] });
    expect(sent.map((item) => item.path)).toEqual(['/me/messages', '/me/eventsArchive', '/me/todo/lists/l1/tasks/t1']);
  });
});
