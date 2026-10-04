import type { Result } from '../domain/result.ts';
import { err } from '../domain/result.ts';
import type { TenantId } from '../domain/tenant-id.ts';
import type { TokenFingerprint } from '../domain/token-fingerprint.ts';
import type { ProcessRunner, ProcessRunnerError } from '../use-cases/ports/process-runner.ts';
import type { TokenTier } from '../use-cases/ports/token-issuer.ts';
import type { TokenError } from '../use-cases/ports/token-source.ts';
import { TOKEN_VARIABLE } from './env-token-source.ts';
import type { HelperToken } from './token-helper-answer.ts';
import { readHelperAnswer, tierCode } from './token-helper-answer.ts';
import type { TokenHelperCommand, TokenHelperLocateError } from './token-helper-locator.ts';

/*
 * One run of the token helper, a separate process: `<helper> --tier <tier>
 * [--tenant <guid>] [--reject <fingerprint>]`, never through a shell. Its
 * stdin is the terminal's only when a person is there, its stdout is capped,
 * and its stderr is the caller's. A run that cannot start, is stopped or
 * answers with no token is a failure of the tier.
 */

export type TokenHelperRequest = { readonly tier: TokenTier; readonly tenant?: TenantId; readonly rejected?: TokenFingerprint };

export type TokenHelperRunDeps = {
  readonly locate: () => Promise<Result<TokenHelperCommand, TokenHelperLocateError>>;
  readonly runner: ProcessRunner;
  // A person at the terminal: the helper gets stdin and may open a browser.
  readonly interactive: boolean;
};

// Interactive: the helper's wait for a lock held by another sign-in (7 minutes,
// LOCK_WAIT_INTERACTIVE_MS in auth.ts) plus a full browser sign-in of its own
// (about 6). Otherwise its unattended lock wait (20 s) and one token request
// (60 s), with room to start.
const INTERACTIVE_DEADLINE_MS = 13 * 60_000;
const UNATTENDED_DEADLINE_MS = 90_000;
// A token line is a few kilobytes.
const MAX_STDOUT_BYTES = 64 * 1024;

const argsFor = (request: TokenHelperRequest): ReadonlyArray<string> => [
  '--tier',
  request.tier,
  ...(request.tenant === undefined ? [] : ['--tenant', request.tenant]),
  ...(request.rejected === undefined ? [] : ['--reject', request.rejected]),
];

const remedyFor = (tier: TokenTier): string =>
  tier === 'guest'
    ? 'Guest tokens come only from the token helper: install @ask-marcel/office-auth.'
    : `To fix it, install @ask-marcel/office-auth, or set ${TOKEN_VARIABLE[tier]} to a token.`;

const NOT_LOCATED: Readonly<Record<TokenHelperLocateError['type'], string>> = {
  not_found:
    'No token helper was found: ASKMARCEL_TOKEN_COMMAND is not set, ~/.ask-marcel/token-helper.json names no helper that is still installed, and ask-marcel-office-auth is not on PATH.',
  not_absolute: 'ASKMARCEL_TOKEN_COMMAND must be an absolute path to the token helper, so it was not run.',
};

const unavailable = (tier: TokenTier, why: string): Result<never, TokenError> =>
  err({ type: 'auth_failed', message: `${why} ${remedyFor(tier)}`, code: 'token_helper_unavailable' });

type Stopped = Exclude<ProcessRunnerError, { readonly type: 'spawn_failed' }>;

const stoppedBecause = (error: Stopped): string => {
  if (error.type === 'timed_out') return `did not answer within ${error.timeoutMs / 1000} s and was stopped`;
  if (error.type === 'killed') return `was stopped by ${error.signal}`;
  return `printed more than ${error.maxStdoutBytes} bytes and was stopped`;
};

const runFailed = (tier: TokenTier, command: string, error: ProcessRunnerError): Result<never, TokenError> => {
  if (error.type === 'spawn_failed') return unavailable(tier, `The token helper ${command} could not be started (${error.message}).`);
  return err({ type: 'auth_failed', message: `The token helper ${stoppedBecause(error)}, so no ${tier} token was handed out.`, code: tierCode(tier) });
};

export const runTokenHelper = async (deps: TokenHelperRunDeps, request: TokenHelperRequest): Promise<Result<HelperToken, TokenError>> => {
  const helper = await deps.locate();
  if (!helper.ok) return unavailable(request.tier, NOT_LOCATED[helper.error.type]);
  const options = {
    stdin: deps.interactive ? 'inherit' : 'ignore',
    timeoutMs: deps.interactive ? INTERACTIVE_DEADLINE_MS : UNATTENDED_DEADLINE_MS,
    maxStdoutBytes: MAX_STDOUT_BYTES,
  } as const;
  const ran = await deps.runner.run(helper.value.command, [...helper.value.args, ...argsFor(request)], options);
  return ran.ok ? readHelperAnswer(request.tier, ran.value) : runFailed(request.tier, helper.value.command, ran.error);
};
