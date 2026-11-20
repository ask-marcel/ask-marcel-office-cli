import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { AuthSession, CachedTierInfo, SessionEnd } from '../ports/auth-session.ts';
import { execute } from './login.ts';
import { buildLoginSummary } from './login-status.ts';

// `login --renew` starts a new sign-in session of about 24 hours before the
// current one ends: the browser capture of every token, with the cookies kept.
// A plain `login` does not change while the session is good, but a session past
// its predicted end is not good: a plain `login` then starts a new one. The
// answer of a successful `login` gives the end of the session.

type SignInOptions = { readonly force?: boolean; readonly renew?: boolean };

const AVAILABLE: CachedTierInfo = { available: true, expiresInSeconds: 3600, scopes: [] };

// A session that records each sign-in it is asked for. Every read says that
// the session is good, except the end of the session that the test gives.
const sessionEndingIn = (endsInSeconds: number | undefined): AuthSession & { readonly signIns: Array<SignInOptions | undefined> } => {
  const signIns: Array<SignInOptions | undefined> = [];
  const end: SessionEnd | undefined = endsInSeconds === undefined ? undefined : { endsAt: '2026-11-21T20:09:00.000Z', endsInSeconds };
  return {
    signIns,
    getAccessToken: async (options) => {
      signIns.push(options);
      return ok(accessTokenUnsafe('tok'));
    },
    getElevatedAccessToken: async () => ok(accessTokenUnsafe('elevated')),
    warmSubstrateTokens: async () => {},
    getCachedElevatedInfo: async () => AVAILABLE,
    hasRefreshToken: async () => true,
    getCachedChatsvcaggInfo: async () => AVAILABLE,
    getCachedIc3Info: async () => AVAILABLE,
    getSessionEnd: async () => end,
    logout: async () => ok(undefined),
  };
};

describe('login --renew', () => {
  it('runs the capture with the cookies kept one time, and nothing after it', async () => {
    const auth = sessionEndingIn(43_200);
    expect((await execute(auth, { force: false, renew: true })).ok).toBe(true);
    expect(auth.signIns).toEqual([{ force: false, renew: true }]);
  });

  it('with --force runs only the forced capture, which also clears the cookies', async () => {
    const auth = sessionEndingIn(43_200);
    await execute(auth, { force: true, renew: true });
    expect(auth.signIns).toEqual([{ force: true, renew: true }]);
  });
});

describe('a plain login', () => {
  it('opens no browser while the sign-in session did not end', async () => {
    const auth = sessionEndingIn(43_200);
    await execute(auth, { force: false, renew: false });
    expect(auth.signIns).toEqual([{ force: false, renew: false }]);
  });

  it('opens no browser when the end of the session is not known', async () => {
    const auth = sessionEndingIn(undefined);
    await execute(auth, { force: false, renew: false });
    expect(auth.signIns).toEqual([{ force: false, renew: false }]);
  });

  it('starts a new session with the cookies kept when the session is past its predicted end', async () => {
    const auth = sessionEndingIn(-60);
    await execute(auth, { force: false, renew: false });
    expect(auth.signIns).toEqual([{ force: false, renew: false }, { renew: true }]);
  });
});

describe('the answer of a successful login', () => {
  const tiers = { elevatedAvailable: true, chatsvcaggAvailable: true, ic3Available: true };

  it('gives the end of the sign-in session as a time and the seconds left, before the hint', () => {
    const summary = buildLoginSummary({ ...tiers, session: { endsAt: '2026-11-21T20:09:00.000Z', endsInSeconds: 86_390 } });
    expect(summary.session).toEqual({ endsAt: '2026-11-21T20:09:00.000Z', endsInSeconds: 86_390 });
    expect(Object.keys(summary)).toEqual(['status', 'available', 'unlocked', 'missing', 'session', 'hint']);
  });

  it('says that the end is not known when the cache holds no time of the sign-in', () => {
    const summary = buildLoginSummary(tiers);
    expect(summary.session).toEqual({ reason: expect.stringContaining('not known') });
  });

  it('names login --renew in its hint', () => {
    expect(buildLoginSummary(tiers).hint).toContain('`{auth-bin} login --renew`');
  });
});
