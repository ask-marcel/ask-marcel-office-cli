import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../../domain/access-token.ts';
import { accessTokenUnsafe } from '../../domain/access-token.ts';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import { tenantIdUnsafe } from '../../domain/tenant-id.ts';
import type { TokenFingerprint } from '../../domain/token-fingerprint.ts';
import { tokenFingerprint } from '../../domain/token-fingerprint.ts';
import type { TokenIssuer, TokenRequest } from '../ports/token-issuer.ts';
import type { TokenError } from '../ports/token-source.ts';
import { execute } from './token.ts';
import { parseTokenArgs } from './token-args.ts';

const GUID = '8f2c1a4e-3b6d-4c9a-9e1f-2a7b5c8d0e3f';
const EXP = 1_900_000_000;

const jwt = (claims: Record<string, unknown>): AccessToken => accessTokenUnsafe(`${btoa(JSON.stringify({ alg: 'RS256' }))}.${btoa(JSON.stringify(claims))}.sig`);
const TOKEN = jwt({ exp: EXP, aud: 'https://graph.microsoft.com' });

// A hand-written issuer: answers what the test says and records what it was asked.
const fakeIssuer = (
  answer: Result<AccessToken, TokenError>,
  region: string | Result<string, TokenError> = 'emea'
): TokenIssuer & { readonly asked: Array<{ request: TokenRequest; rejected?: TokenFingerprint }> } => {
  const asked: Array<{ request: TokenRequest; rejected?: TokenFingerprint }> = [];
  return {
    asked,
    issueToken: async (request, rejected) => {
      asked.push(rejected === undefined ? { request } : { request, rejected });
      return answer;
    },
    cachedRegion: async () => (typeof region === 'string' ? ok(region) : region),
  };
};

const issue = async (argv: ReadonlyArray<string>, issuer: TokenIssuer): Promise<unknown> => {
  const args = parseTokenArgs(argv);
  if (!args.ok) throw new Error(`the test passed arguments the helper refuses: ${args.error.message}`);
  const line = await execute(issuer, args.value);
  return line.ok ? line.value : line.error;
};

describe('token helper: the line it prints', () => {
  it('prints the basic token with its expiry in seconds since the epoch, and no region', async () => {
    expect(await issue(['--tier', 'basic'], fakeIssuer(ok(TOKEN)))).toEqual({ accessToken: TOKEN, expiresOn: EXP });
  });

  it('adds the Teams region to a chatsvcagg or ic3 token, since both route by it', async () => {
    expect(await issue(['--tier', 'chatsvcagg'], fakeIssuer(ok(TOKEN), 'amer'))).toEqual({ accessToken: TOKEN, expiresOn: EXP, region: 'amer' });
    expect(await issue(['--tier', 'ic3'], fakeIssuer(ok(TOKEN), 'apac'))).toEqual({ accessToken: TOKEN, expiresOn: EXP, region: 'apac' });
  });

  it('reports a token that carries no expiry as already expired', async () => {
    expect(await issue(['--tier', 'elevated'], fakeIssuer(ok(jwt({ aud: 'https://graph.microsoft.com' }))))).toMatchObject({ expiresOn: 0 });
  });

  it('asks the issuer for the guest token of the named tenant, and passes a refused fingerprint on', async () => {
    const fingerprint = await tokenFingerprint('the-refused-one');
    const issuer = fakeIssuer(ok(TOKEN));
    expect(await issue(['--tier', 'guest', '--tenant', GUID, '--reject', fingerprint], issuer)).toEqual({ accessToken: TOKEN, expiresOn: EXP });
    expect(issuer.asked).toEqual([{ request: { tier: 'guest', tenant: tenantIdUnsafe(GUID) }, rejected: fingerprint }]);
  });

  // The region is pasted into every chat URL the caller builds, so a cache that
  // names something else is a failure, never a path to send.
  it('fails rather than hand out a region that is not a region name, and prints no token', async () => {
    const line = await issue(['--tier', 'ic3'], fakeIssuer(ok(TOKEN), 'emea/../admin'));
    expect(line).toMatchObject({ errorCode: 'secondary_token_unavailable', tier: 'ic3', message: expect.stringContaining('not a region name') });
    expect(JSON.stringify(line)).not.toContain(TOKEN);
    expect((line as { remedy: string }).remedy).toContain('ask-marcel-office login --force');
  });

  // The region is read after the token is issued; a cache gone or unreadable
  // by then (a logout in between) must not turn into the default region.
  it('fails a chat token whose region cannot be read back, rather than hand it out with a guessed region', async () => {
    const unreadable = err({ type: 'auth_failed', message: 'The token cache could not be read for the Teams region.', code: 'secondary_token_unavailable' } as const);
    const line = await issue(['--tier', 'chatsvcagg'], fakeIssuer(ok(TOKEN), unreadable));
    expect(line).toMatchObject({ errorCode: 'secondary_token_unavailable', tier: 'chatsvcagg', message: 'The token cache could not be read for the Teams region.' });
    expect(JSON.stringify(line)).not.toContain(TOKEN);
  });

  it('reports a missing sign-in as not_authenticated, with the remedy naming login and status', async () => {
    const line = await issue(['--tier', 'basic'], fakeIssuer(err({ type: 'auth_failed', message: 'Not signed in.', code: 'not_authenticated' })));
    expect(line).toEqual({
      errorCode: 'not_authenticated',
      tier: 'basic',
      message: 'Not signed in.',
      remedy: 'Run `ask-marcel-office login` in a terminal to sign in, then check every token with `ask-marcel-office status`.',
    });
  });

  it('reports a closed sign-in window as auth_cancelled', async () => {
    expect(await issue(['--tier', 'basic'], fakeIssuer(err({ type: 'auth_cancelled' })))).toMatchObject({
      errorCode: 'auth_cancelled',
      tier: 'basic',
      message: 'The sign-in was cancelled.',
    });
  });

  it('passes the ladder codes a caller acts on through unchanged', async () => {
    for (const code of ['not_authenticated', 'sign_in_in_progress', 'token_cache_unwritable']) {
      expect(await issue(['--tier', 'elevated'], fakeIssuer(err({ type: 'auth_failed', message: 'm', code })))).toMatchObject({ errorCode: code });
    }
    for (const code of ['secondary_token_unavailable', 'sign_in_in_progress', 'token_cache_unwritable']) {
      expect(await issue(['--tier', 'basic'], fakeIssuer(err({ type: 'auth_failed', message: 'm', code })))).toMatchObject({ errorCode: code });
    }
  });

  // A failure with no code (a browser that crashed, a code this version does not
  // know) still gets a code: the basic tier's means sign in again, the other
  // tiers' means that token is out of reach.
  it('gives an uncoded failure the code of its tier', async () => {
    expect(await issue(['--tier', 'basic'], fakeIssuer(err({ type: 'auth_failed', message: 'browser crashed' })))).toMatchObject({ errorCode: 'not_authenticated' });
    expect(await issue(['--tier', 'guest', '--tenant', GUID], fakeIssuer(err({ type: 'auth_failed', message: 'm', code: 'new_code' })))).toMatchObject({
      errorCode: 'secondary_token_unavailable',
    });
  });

  // A caller pastes the remedy in front of a person, who may not know which
  // program to run: every remedy names the bin and the command that checks.
  it('names ask-marcel-office and its status command in the remedy of every failure', async () => {
    const remedies: string[] = [];
    const remedyOf = async (argv: ReadonlyArray<string>, issuer: TokenIssuer): Promise<void> => {
      remedies.push(((await issue(argv, issuer)) as { remedy: string }).remedy);
    };
    const tiers = [
      ['--tier', 'basic'],
      ['--tier', 'elevated'],
      ['--tier', 'chatsvcagg'],
      ['--tier', 'ic3'],
      ['--tier', 'guest', '--tenant', GUID],
    ];
    for (const argv of tiers) {
      for (const code of ['not_authenticated', 'secondary_token_unavailable', 'sign_in_in_progress', 'token_cache_unwritable', undefined]) {
        await remedyOf(argv, fakeIssuer(err({ type: 'auth_failed', message: 'm', code })));
      }
      await remedyOf(argv, fakeIssuer(err({ type: 'auth_cancelled' })));
    }
    await remedyOf(['--tier', 'ic3'], fakeIssuer(ok(TOKEN), 'emea/../admin'));
    expect(remedies).toHaveLength(31);
    for (const remedy of remedies) {
      expect(remedy).toContain('`ask-marcel-office');
      expect(remedy).toContain('ask-marcel-office status');
    }
  });

  it('tells each failure what to do next', async () => {
    const remedyOf = async (argv: ReadonlyArray<string>, code: string): Promise<string> =>
      ((await issue(argv, fakeIssuer(err({ type: 'auth_failed', message: 'm', code })))) as { remedy: string }).remedy;
    expect(await remedyOf(['--tier', 'elevated'], 'secondary_token_unavailable')).toContain('with a browser');
    expect(await remedyOf(['--tier', 'ic3'], 'secondary_token_unavailable')).toContain('Teams chat tokens');
    expect(await remedyOf(['--tier', 'chatsvcagg'], 'secondary_token_unavailable')).toContain('Teams chat tokens');
    expect(await remedyOf(['--tier', 'guest', '--tenant', GUID], 'secondary_token_unavailable')).toContain('guest in that tenant');
    expect(await remedyOf(['--tier', 'basic'], 'sign_in_in_progress')).toContain('Wait for');
    expect(await remedyOf(['--tier', 'basic'], 'token_cache_unwritable')).toContain('writable');
    expect(((await issue(['--tier', 'basic'], fakeIssuer(err({ type: 'auth_cancelled' })))) as { remedy: string }).remedy).toContain('finish the sign-in');
  });
});
