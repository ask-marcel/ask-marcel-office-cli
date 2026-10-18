import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import type { ProcessRunnerCall } from '@ask-marcel/office-core/test-helpers/process-runner-fake.ts';
import { createProcessRunnerFake } from '@ask-marcel/office-core/test-helpers/process-runner-fake.ts';
import { buildWriteTokenSource, createEnvThenHelperTokenSource, WRITE_NEEDED_BY } from './write-token-source.ts';

// The write package signs only through a tier variable or the token helper of
// the auth package (E3, LESSONS 2026-10-18): it holds no token cache and opens
// no browser. The tokens here are built at run time, with no signature.

const HOME = '/virtual/home';
const HELPER = '/opt/ask-marcel/auth/dist/token.js';
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const graphToken = (from: string): string => unsignedJwt({ aud: 'https://graph.microsoft.com', exp: inAnHour(), from });

// The helper answers each tier with its token, the way the auth package does.
const answeringHelper = (token: string): ReturnType<typeof createProcessRunnerFake> =>
  createProcessRunnerFake((_call: ProcessRunnerCall) => ok({ exitCode: 0, stdout: `${JSON.stringify({ accessToken: token, expiresOn: inAnHour() })}\n` }));

// The helper refuses the elevated token, the way the auth package does when no browser sign-in got it.
const refusingHelper = (): ReturnType<typeof createProcessRunnerFake> =>
  createProcessRunnerFake(() =>
    ok({ exitCode: 1, stdout: `${JSON.stringify({ errorCode: 'secondary_token_unavailable', tier: 'elevated', message: 'no token', remedy: 'run login' })}\n` })
  );

describe('the token source of the write package', () => {
  it('names no write command after a missing secondary token: no write command needs one', () => {
    expect(WRITE_NEEDED_BY).toEqual({ elevated: [], chatsvcagg: [], ic3: [] });
  });

  it('takes the tier variable ahead of the helper', async () => {
    const own = graphToken('env');
    const runner = answeringHelper(graphToken('helper'));
    const source = buildWriteTokenSource({
      env: { ASKMARCEL_TOKEN_BASIC: own, ASKMARCEL_TOKEN_COMMAND: HELPER },
      home: HOME,
      fs: createFileSystemFake(),
      runner,
      interactive: false,
    });
    expect(await source.graphToken('basic')).toEqual(ok(accessTokenUnsafe(own)));
    expect(runner.calls).toEqual([]);
  });

  it('runs the helper at ASKMARCEL_TOKEN_COMMAND when no tier variable is set, with no stdin in a session with no person', async () => {
    const token = graphToken('helper');
    const runner = answeringHelper(token);
    const source = buildWriteTokenSource({ env: { ASKMARCEL_TOKEN_COMMAND: HELPER }, home: HOME, fs: createFileSystemFake(), runner, interactive: false });
    expect(await source.graphToken('basic')).toEqual(ok(accessTokenUnsafe(token)));
    expect(runner.calls.map((call) => [call.command, ...call.args, call.options.stdin])).toEqual([[HELPER, '--tier', 'basic', 'ignore']]);
  });

  it('keeps the message of auth for a refused secondary token, and adds no command list', async () => {
    const source = buildWriteTokenSource({ env: { ASKMARCEL_TOKEN_COMMAND: HELPER }, home: HOME, fs: createFileSystemFake(), runner: refusingHelper(), interactive: false });
    const result = await source.graphToken('elevated');
    const error = result.ok ? undefined : result.error;
    expect(error?.type).toBe('auth_failed');
    if (error?.type === 'auth_failed') expect([error.code, error.message.includes('Commands that need')]).toEqual(['secondary_token_unavailable', false]);
  });

  it('gives a library caller a token source with no parameter to pass', () => {
    const source = createEnvThenHelperTokenSource();
    expect(Object.keys(source).toSorted((a, b) => a.localeCompare(b))).toEqual(['graphToken', 'guestToken', 'substrateRegion', 'substrateToken']);
  });
});
