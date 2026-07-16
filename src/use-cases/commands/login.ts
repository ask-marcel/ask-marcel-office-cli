import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import type { AuthManager } from '../../infra/auth.ts';

const schema = z.object({}).strict();

/**
 * Sign in, and guarantee the browser-only tier is actually there afterwards.
 *
 * Every other token tier can be renewed headlessly from the shared refresh
 * token. The elevated (M365ChatClient) token cannot — it carries no refresh
 * token of its own and exists only via the browser dance. So a plain `login`
 * that found a valid cached basic token used to return on the cache rung and
 * report `authenticated` with elevated still missing, which left the user in a
 * loop with no exit: the command that needs elevated says "run login", login
 * says "authenticated", the command fails identically, forever.
 *
 * When elevated is confirmed missing we escalate through the SAME
 * `{ force: true }` rung the `--force` flag uses. That matters: the browser
 * adapter has its own `freshCachedToken` probe that short-circuits the dance
 * when a valid token is on disk, and only the forced path suppresses it — a
 * hand-rolled re-capture here would silently no-op (LESSONS 2026-07-13).
 *
 * Supersedes the 2026-07-13 decision that `login` is a slim confirmation and
 * `--force` the only re-capture mechanism: correct about the mechanism, but it
 * left the signpost pointing at a command that could not do the job.
 */
const execute = async (auth: AuthManager, options?: { force?: boolean }): Promise<Result<string, import('../../infra/auth.ts').AuthError>> => {
  const authenticated = await auth.getAccessToken(options);
  // Already forced: the dance just ran. If elevated is still missing (a
  // federated tenant whose capture failed), escalating again would open a
  // second browser for nothing.
  if (!authenticated.ok || options?.force === true) return authenticated;
  // `getCachedElevatedInfo` is an optional capability; when a manager does not
  // expose it we cannot know, so we leave the result alone rather than guess.
  const elevated = await auth.getCachedElevatedInfo?.();
  if (elevated?.available !== false) return authenticated;
  return auth.getAccessToken({ force: true });
};

export { execute, schema };
