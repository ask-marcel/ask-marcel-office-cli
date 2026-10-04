import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import { tokenFingerprint } from '../domain/token-fingerprint.ts';
import type { ProcessRunnerCall } from '../test-helpers/process-runner-fake.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import type { ProcessRunnerError, ProcessRunResult } from '../use-cases/ports/process-runner.ts';
import type { TokenError } from '../use-cases/ports/token-source.ts';
import type { TokenHelperCommand, TokenHelperLocateError } from './token-helper-locator.ts';
import type { TokenHelperRequest } from './token-helper-run.ts';
import { runTokenHelper } from './token-helper-run.ts';

const MINUTE = 60_000;
const HELPER: TokenHelperCommand = { command: '/usr/local/bin/node', args: ['/opt/ask-marcel/dist/token.js'] };
const TENANT = tenantIdUnsafe('11111111-2222-3333-4444-555555555555');
const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const jwt = (name: string): AccessToken => accessTokenUnsafe(`${segment({ alg: 'RS256' })}.${segment({ name })}.sig`);
const inAnHour = 1_800_003_600;

type Answer = Result<ProcessRunResult, ProcessRunnerError>;
type Setup = { readonly answer?: Answer; readonly interactive?: boolean; readonly located?: Result<TokenHelperCommand, TokenHelperLocateError> };

// One helper run over a fake runner that hands back `answer`.
const runWith = async (
  request: TokenHelperRequest,
  setup: Setup = {}
): Promise<{ result: Awaited<ReturnType<typeof runTokenHelper>>; calls: ReadonlyArray<ProcessRunnerCall> }> => {
  const runner = createProcessRunnerFake(() => setup.answer ?? ok({ exitCode: 0, stdout: JSON.stringify({ accessToken: jwt('t'), expiresOn: inAnHour }) }));
  const result = await runTokenHelper({ locate: async () => setup.located ?? ok(HELPER), runner, interactive: setup.interactive ?? false }, request);
  return { result, calls: runner.calls };
};

const failure = (result: Result<unknown, TokenError>): { message: string; code?: string } => {
  if (result.ok || result.error.type !== 'auth_failed') throw new Error('expected an auth_failed error');
  return result.error;
};

describe('one run of the token helper', () => {
  it('runs the located helper with --tier, with no stdin, a short deadline and capped stdout, and reads the token it prints', async () => {
    const { result, calls } = await runWith({ tier: 'basic' });
    expect(result).toEqual(ok({ token: jwt('t'), expiresOn: inAnHour }));
    expect(calls).toEqual([{ command: HELPER.command, args: [...HELPER.args, '--tier', 'basic'], options: { stdin: 'ignore', timeoutMs: 90_000, maxStdoutBytes: 64 * 1024 } }]);
  });

  it('hands a person at the terminal its stdin and a deadline long enough to wait for a held lock and finish a browser sign-in', async () => {
    const { calls } = await runWith({ tier: 'elevated' }, { interactive: true });
    expect(calls[0]?.options).toEqual({ stdin: 'inherit', timeoutMs: 13 * MINUTE, maxStdoutBytes: 64 * 1024 });
  });

  it('names the partner tenant of a guest token, and the fingerprint of a refused token', async () => {
    const rejected = await tokenFingerprint(jwt('refused'));
    expect((await runWith({ tier: 'guest', tenant: TENANT })).calls[0]?.args.slice(1)).toEqual(['--tier', 'guest', '--tenant', TENANT]);
    expect((await runWith({ tier: 'ic3', rejected })).calls[0]?.args.slice(1)).toEqual(['--tier', 'ic3', '--reject', rejected]);
  });

  it('reports a helper that outlived its deadline, one killed by a signal and one that printed too much as failures of the tier', async () => {
    const outcomes: ReadonlyArray<[ProcessRunnerError, string]> = [
      [{ type: 'timed_out', timeoutMs: 90_000 }, 'did not answer within 90 s'],
      [{ type: 'killed', signal: 'SIGTERM' }, 'stopped by SIGTERM'],
      [{ type: 'output_too_large', maxStdoutBytes: 65_536 }, 'more than 65536 bytes'],
    ];
    for (const [outcome, said] of outcomes) {
      const error = failure((await runWith({ tier: 'basic' }, { answer: err(outcome) })).result);
      expect(error.code).toBe('not_authenticated');
      expect(error.message).toContain(said);
    }
  });

  it('reports a helper that could not be started as token_helper_unavailable', async () => {
    const error = failure((await runWith({ tier: 'basic' }, { answer: err({ type: 'spawn_failed', message: 'ENOENT' }) })).result);
    expect(error.code).toBe('token_helper_unavailable');
    expect(error.message).toContain('could not be started (ENOENT)');
    expect(error.message).toContain('ASKMARCEL_TOKEN_BASIC');
  });

  it('reports no helper found as token_helper_unavailable, naming the tier variable that would stand in for it, and runs nothing', async () => {
    const { result, calls } = await runWith({ tier: 'chatsvcagg' }, { located: err({ type: 'not_found' }) });
    const error = failure(result);
    expect(error.code).toBe('token_helper_unavailable');
    expect(error.message).toContain('install @ask-marcel/office-auth');
    expect(error.message).toContain('ASKMARCEL_TOKEN_CHATSVCAGG');
    expect(calls).toEqual([]);
  });

  it('says that guest tokens need the auth package when no helper is found, since they have no variable', async () => {
    const error = failure((await runWith({ tier: 'guest', tenant: TENANT }, { located: err({ type: 'not_found' }) })).result);
    expect(error.code).toBe('token_helper_unavailable');
    expect(error.message).toContain('Guest tokens come only from the token helper');
  });

  it('refuses an ASKMARCEL_TOKEN_COMMAND that is not an absolute path, without running it', async () => {
    const error = failure((await runWith({ tier: 'elevated' }, { located: err({ type: 'not_absolute' }) })).result);
    expect(error.code).toBe('token_helper_unavailable');
    expect(error.message).toContain('ASKMARCEL_TOKEN_COMMAND must be an absolute path');
  });
});
