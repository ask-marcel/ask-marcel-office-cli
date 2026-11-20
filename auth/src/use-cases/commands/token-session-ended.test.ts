import { describe, expect, it } from 'bun:test';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { AuthError } from '../ports/auth-session.ts';
import type { TokenIssuer } from '../ports/token-issuer.ts';
import { execute } from './token.ts';
import { parseTokenArgs } from './token-args.ts';

// When Entra refused the refresh token because the sign-in session ended, the
// helper line keeps the code of its tier and adds `cause: "session_ended"`, so
// the read and write bins can name the end of the session. The message of the
// ladder already tells to run a plain `login`, so the remedy does not say it a
// second time.

const ENDED_MESSAGE = 'The sign-in session ended, so Entra refused to renew the token. Run `{auth-bin} login` in a terminal: a plain login starts a new session.';

const refusing = (error: AuthError): TokenIssuer => ({
  issueToken: async () => err(error),
  cachedRegion: async () => ok('emea'),
});

const lineFor = async (tier: string, error: AuthError): Promise<unknown> => {
  const args = parseTokenArgs(['--tier', tier]);
  if (!args.ok) throw new Error('the test passed arguments the helper refuses');
  const line = await execute(refusing(error), args.value);
  return line.ok ? line.value : line.error;
};

describe('token helper: a sign-in session that ended', () => {
  it('gives the basic tier the code not_authenticated, the cause session_ended and the message of the ladder', async () => {
    const line = await lineFor('basic', { type: 'auth_failed', message: ENDED_MESSAGE, code: 'not_authenticated', cause: 'session_ended' });
    expect(line).toMatchObject({ errorCode: 'not_authenticated', tier: 'basic', message: ENDED_MESSAGE, cause: 'session_ended' });
  });

  it('gives a Teams chat tier the code secondary_token_unavailable and the cause session_ended', async () => {
    const line = await lineFor('chatsvcagg', { type: 'auth_failed', message: ENDED_MESSAGE, code: 'secondary_token_unavailable', cause: 'session_ended' });
    expect(line).toMatchObject({ errorCode: 'secondary_token_unavailable', tier: 'chatsvcagg', cause: 'session_ended' });
  });

  it('gives a remedy that does not tell to run login a second time, and names status for the end of the new session', async () => {
    for (const tier of ['basic', 'chatsvcagg', 'ic3']) {
      const line = (await lineFor(tier, { type: 'auth_failed', message: ENDED_MESSAGE, code: 'secondary_token_unavailable', cause: 'session_ended' })) as { remedy: string };
      expect(line.remedy).not.toContain('{auth-bin} login');
      expect(line.remedy).toContain('{auth-bin} status');
    }
  });

  it('gives no cause to a failure that has none', async () => {
    const line = await lineFor('ic3', { type: 'auth_failed', message: 'ic3 token is expired.', code: 'secondary_token_unavailable' });
    expect(Object.keys(line as object)).toEqual(['errorCode', 'tier', 'message', 'remedy']);
  });
});
