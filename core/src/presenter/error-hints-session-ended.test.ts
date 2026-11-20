import { describe, expect, it } from 'bun:test';
import { findErrorHint } from './error-hints.ts';

// A read or write command whose token helper says that the sign-in session
// ended gets one hint: it names the limit of about 24 hours and a plain
// `login`. The hint of an expired or missing token does not come too, so the
// remedy is not given two times.

const ended = (tier: string, code: string): string =>
  `The token helper gave no ${tier} token: exit code 1, errorCode ${code}, cause session_ended, 640 bytes on stdout. Run it by hand (\`{auth-bin} token --tier ${tier}\`) to read its message and remedy.`;

describe('the hint for a sign-in session that ended', () => {
  it('names the end of the session, the limit of about 24 hours and a plain login, for the basic token', () => {
    const hint = findErrorHint(ended('basic', 'not_authenticated'), 'not_authenticated');
    expect(hint?.source).toBe('cli');
    expect(hint?.hint).toContain('sign-in session ended');
    expect(hint?.hint).toContain('about 24 hours');
    expect(hint?.hint).toContain('`{auth-bin} login`');
  });

  it('gives the same hint for a Teams chat token, and not the hint of a token that was not captured', () => {
    const hint = findErrorHint(ended('chatsvcagg', 'secondary_token_unavailable'), 'secondary_token_unavailable');
    expect(hint?.hint).toContain('sign-in session ended');
    expect(hint?.hint).not.toContain('was not captured at login');
  });

  it('says the login only one time', () => {
    const hint = findErrorHint(ended('ic3', 'secondary_token_unavailable'), 'secondary_token_unavailable')?.hint ?? '';
    expect(hint.split('{auth-bin} login').length - 1).toBe(1);
  });
});
