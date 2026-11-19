import type { GraphError } from './graph-error.ts';
import type { Result } from './result.ts';
import { err } from './result.ts';

/*
 * The no-invitation check of the write graph (product rule: no command sends
 * a meeting invitation). Graph sends an invitation to each attendee of an
 * event that a POST creates or a PATCH changes, so an event write whose body
 * has attendees is refused before any request. An event path ends in
 * `events` (a create) or in `events/{id}` (a change), under any mailbox or
 * calendar. An empty list of attendees sends nothing.
 */

// The segments of the path as the request sends it: the URL parser drops the query and resolves dot segments.
const segmentsOf = (path: string): ReadonlyArray<string> =>
  new URL(`https://graph.invalid/v1.0${path}`).pathname
    .split('/')
    .filter((segment) => segment !== '')
    .map((segment) => segment.toLowerCase());

const isEventPath = (path: string): boolean => segmentsOf(path).slice(-2).includes('events');

// Graph reads a property name with no case, so `Attendees` counts too.
const hasAttendees = (body: unknown): boolean =>
  body !== null && typeof body === 'object' && Object.entries(body).some(([key, value]) => key.toLowerCase() === 'attendees' && Array.isArray(value) && value.length > 0);

/** Whether a write of `body` to `path` would send a meeting invitation. */
const invites = (path: string, body: unknown): boolean => isEventPath(path) && hasAttendees(body);

/** Runs `change` only when the write sends no invitation; refuses it otherwise, before any request. */
const unlessInvites = async <T>(method: string, path: string, body: unknown, change: () => Promise<Result<T, GraphError>>): Promise<Result<T, GraphError>> =>
  invites(path, body)
    ? err({
        type: 'validation_error',
        code: 'invitation_refused',
        message: `${method} ${path} was not sent: the event has attendees, so Graph would send them an invitation, and this CLI never sends one. Give the planned attendees as text in the body.`,
      })
    : change();

export { unlessInvites };
