import type { SessionEnd } from '../ports/auth-session.ts';

// The end of the sign-in session, as `status` and the answer of `login` give it:
// a time and the seconds left when the token cache holds the time of the
// browser sign-in, else a reason that says the end is not known and how to
// record it. A plain `login` with a good session opens no browser, so it does
// not record the time: `login --renew` does.
type SessionReport = SessionEnd | { readonly reason: string };

const SESSION_UNKNOWN_REASON =
  'not known: the token cache does not hold the time of the browser sign-in (a cache written before this version). The next browser sign-in records it: run `{auth-bin} login --renew` to start a new session now.';

const sessionReport = (session: SessionEnd | undefined): SessionReport => session ?? { reason: SESSION_UNKNOWN_REASON };

// Past its predicted end, a session cannot renew itself. The prediction never
// stops a redemption: it only chooses the texts and the sign-in of `login`.
const sessionHasEnded = (session: SessionEnd | undefined): boolean => session !== undefined && session.endsInSeconds <= 0;

export { sessionHasEnded, sessionReport };
export type { SessionReport };
