import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { ok } from '../domain/result.ts';
import type { ProcessRunResult } from '../use-cases/ports/process-runner.ts';
import type { TokenError } from '../use-cases/ports/token-source.ts';
import { readHelperAnswer } from './token-helper-answer.ts';

const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const jwt = (name: string): AccessToken => accessTokenUnsafe(`${segment({ alg: 'RS256' })}.${segment({ name })}.sig`);
const inAnHour = 1_800_003_600;

const printed = (exitCode: number, line: unknown): ProcessRunResult => ({ exitCode, stdout: `${JSON.stringify(line)}\n` });
const bytesOf = (run: ProcessRunResult): string => `${new TextEncoder().encode(run.stdout).byteLength} bytes`;

const failure = (result: Result<unknown, TokenError>): { message: string; code?: string } => {
  if (result.ok || result.error.type !== 'auth_failed') throw new Error('expected an auth_failed error');
  return result.error;
};

describe("reading the token helper's answer", () => {
  it('takes the token, its expiry and the region of a token line printed with exit 0', () => {
    const token = jwt('ic3');
    expect(readHelperAnswer('ic3', printed(0, { accessToken: token, expiresOn: inAnHour, region: 'emea' }))).toEqual(ok({ token, expiresOn: inAnHour, region: 'emea' }));
    expect(readHelperAnswer('basic', printed(0, { accessToken: token, expiresOn: inAnHour, region: 7 }))).toEqual(ok({ token, expiresOn: inAnHour }));
  });

  it('reports a helper that found no session with its exit code, its stdout length and its errorCode, and never its message or what it printed', () => {
    const run = printed(1, { errorCode: 'not_authenticated', tier: 'basic', message: 'helper text with secret-ish detail', remedy: 'helper remedy' });
    const error = failure(readHelperAnswer('basic', run));
    expect(error.code).toBe('not_authenticated');
    expect(error.message).toContain('exit code 1');
    expect(error.message).toContain(bytesOf(run));
    expect(error.message).toContain('errorCode not_authenticated');
    expect(error.message).not.toContain('secret-ish');
    expect(error.message).not.toContain('helper remedy');
  });

  it('counts the stdout length in UTF-8 bytes, not in characters', () => {
    const run = printed(1, { errorCode: 'not_authenticated', message: 'é'.repeat(10) });
    expect(new TextEncoder().encode(run.stdout).byteLength).not.toBe(run.stdout.length);
    expect(failure(readHelperAnswer('basic', run)).message).toContain(`, ${bytesOf(run)} on stdout`);
  });

  it('passes every code the helper may answer through as the error code', () => {
    for (const errorCode of ['secondary_token_unavailable', 'auth_cancelled', 'sign_in_in_progress', 'token_cache_unwritable']) {
      expect(failure(readHelperAnswer('elevated', printed(1, { errorCode, tier: 'elevated', message: 'm', remedy: 'r' }))).code).toBe(errorCode);
    }
  });

  it('reports arguments the helper refused (exit 2) as invalid_arguments, and names the command to run by hand', () => {
    const error = failure(readHelperAnswer('guest', printed(2, { errorCode: 'invalid_arguments', tier: null, message: 'm', remedy: 'r' })));
    expect(error.code).toBe('invalid_arguments');
    expect(error.message).toContain('exit code 2');
    expect(error.message).toContain('`ask-marcel-office token --tier guest --tenant <guid>`');
    expect(failure(readHelperAnswer('basic', printed(1, {}))).message).toContain('`ask-marcel-office token --tier basic`');
  });

  it('never takes a token line printed with a failing exit code: only exit 0 carries a token', () => {
    for (const exitCode of [1, 2]) {
      const error = failure(readHelperAnswer('basic', printed(exitCode, { accessToken: jwt('x'), expiresOn: inAnHour })));
      expect(error.code).toBe('not_authenticated');
      expect(error.message).toContain(`exit code ${exitCode}`);
    }
  });

  it('gives a failure line with a code the caller does not know the code of its tier, and does not repeat the unknown one', () => {
    const error = failure(readHelperAnswer('ic3', printed(1, { errorCode: 'eyJ-looks-like-a-token', tier: 'ic3', message: 'm', remedy: 'r' })));
    expect(error.code).toBe('secondary_token_unavailable');
    expect(error.message).not.toContain('eyJ-looks-like-a-token');
    expect(error.message).toContain('no errorCode it knows');
  });

  it('gives an answer that is not JSON the code of its tier and never quotes it', () => {
    const error = failure(readHelperAnswer('basic', { exitCode: 0, stdout: 'eyJ.raw.bearer printed by mistake' }));
    expect(error.code).toBe('not_authenticated');
    expect(error.message).toContain('exit code 0');
    expect(error.message).toContain('33 bytes');
    expect(error.message).not.toContain('raw.bearer');
  });

  it('refuses a token line whose token is not a JWT, holds a byte no HTTP header can carry, or whose expiry is not a number', () => {
    const lines = [
      { accessToken: 'opaque', expiresOn: inAnHour },
      { accessToken: 'abc.def.ghi', expiresOn: inAnHour },
      { accessToken: `${jwt('x')}SECRETSIG\nPART2`, expiresOn: inAnHour },
      { accessToken: jwt('x'), expiresOn: 'soon' },
      null,
      [jwt('x')],
    ];
    for (const line of lines) {
      const error = failure(readHelperAnswer('elevated', printed(0, line)));
      expect(error.code).toBe('secondary_token_unavailable');
      expect(error.message).not.toContain('SECRETSIG');
    }
  });
});
