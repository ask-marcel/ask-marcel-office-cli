import { describe, expect, it } from 'bun:test';
import type { Result } from '../domain/result.ts';
import type { ProcessRunResult } from '../use-cases/ports/process-runner.ts';
import type { TokenError } from '../use-cases/ports/token-source.ts';
import { readHelperAnswer } from './token-helper-answer.ts';

// The token helper adds `cause: "session_ended"` to its failure line when Entra
// refused the refresh token because the sign-in session ended. The reader
// names that cause in its message, so the hint of the read and write bins can
// name the end of the session. Like the errorCode, only a cause it knows gets
// through: nothing else of the helper's line reaches the message.

const printed = (line: unknown): ProcessRunResult => ({ exitCode: 1, stdout: `${JSON.stringify(line)}\n` });

const failure = (result: Result<unknown, TokenError>): { message: string; code?: string } => {
  if (result.ok || result.error.type !== 'auth_failed') throw new Error('expected an auth_failed error');
  return result.error;
};

describe("reading the token helper's answer for a sign-in session that ended", () => {
  it('names the cause session_ended after the errorCode, and keeps the code of the tier', () => {
    const run = printed({ errorCode: 'secondary_token_unavailable', tier: 'ic3', message: 'helper text', remedy: 'helper remedy', cause: 'session_ended' });
    const error = failure(readHelperAnswer('ic3', run));
    expect(error.code).toBe('secondary_token_unavailable');
    expect(error.message).toContain('errorCode secondary_token_unavailable, cause session_ended, ');
    expect(error.message).not.toContain('helper text');
  });

  it('drops a cause that it does not know', () => {
    const run = printed({ errorCode: 'not_authenticated', tier: 'basic', message: 'm', remedy: 'r', cause: 'something-secret' });
    const error = failure(readHelperAnswer('basic', run));
    expect(error.message).not.toContain('cause');
    expect(error.message).not.toContain('something-secret');
  });
});
