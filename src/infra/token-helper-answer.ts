import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe, isJwtShaped } from '../domain/access-token.ts';
import { parseJson } from '../domain/json.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { ProcessRunResult } from '../use-cases/ports/process-runner.ts';
import type { TokenTier } from '../use-cases/ports/token-issuer.ts';
import type { TokenError } from '../use-cases/ports/token-source.ts';

/*
 * Reads the token helper's one line (docs/plans/2026-10-01-package-split.md,
 * Token protocol): `{ accessToken, expiresOn, region? }` with exit 0, or
 * `{ errorCode, tier, message, remedy }` with exit 1 or 2. A failure is told
 * with the exit code, the stdout length and an errorCode this caller knows,
 * and nothing else the helper printed: its stdout reaches no error, log or MCP
 * result, so a token printed where it should not be goes nowhere.
 */

// `region` comes with the two chat tiers, unchecked: it is checked where a URL
// takes it.
export type HelperToken = { readonly token: AccessToken; readonly expiresOn: number; readonly region?: string };

const KNOWN_CODES: ReadonlyArray<string> = [
  'not_authenticated',
  'secondary_token_unavailable',
  'auth_cancelled',
  'sign_in_in_progress',
  'token_cache_unwritable',
  'invalid_arguments',
];

// The helper's own rule: a failure that carries no code gets its tier's.
export const tierCode = (tier: TokenTier): string => (tier === 'basic' ? 'not_authenticated' : 'secondary_token_unavailable');

const fieldOf = (parsed: unknown, name: string): unknown =>
  typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed) ? (parsed as Record<string, unknown>)[name] : undefined;

const knownCode = (parsed: unknown): string | undefined => {
  const code = fieldOf(parsed, 'errorCode');
  return KNOWN_CODES.find((known) => known === code);
};

const byHand = (tier: TokenTier): string => `ask-marcel-office token --tier ${tier}${tier === 'guest' ? ' --tenant <guid>' : ''}`;

const failed = (tier: TokenTier, run: ProcessRunResult, parsed: unknown): Result<never, TokenError> => {
  const code = knownCode(parsed);
  const told = code === undefined ? 'no errorCode it knows' : `errorCode ${code}`;
  const bytes = new TextEncoder().encode(run.stdout).byteLength;
  return err({
    type: 'auth_failed',
    message: `The token helper gave no ${tier} token: exit code ${run.exitCode}, ${told}, ${bytes} bytes on stdout. Run it by hand (\`${byHand(tier)}\`) to read its message and remedy.`,
    code: code ?? tierCode(tier),
  });
};

const tokenOf = (parsed: unknown): HelperToken | undefined => {
  const token = fieldOf(parsed, 'accessToken');
  const expiresOn = fieldOf(parsed, 'expiresOn');
  const region = fieldOf(parsed, 'region');
  if (typeof token !== 'string' || !token.startsWith('eyJ') || !isJwtShaped(token) || typeof expiresOn !== 'number') return undefined;
  return { token: accessTokenUnsafe(token), expiresOn, ...(typeof region === 'string' ? { region } : {}) };
};

export const readHelperAnswer = (tier: TokenTier, run: ProcessRunResult): Result<HelperToken, TokenError> => {
  const parsed = parseJson(run.stdout);
  const answer = parsed.ok ? parsed.value : undefined;
  const token = run.exitCode === 0 ? tokenOf(answer) : undefined;
  return token === undefined ? failed(tier, run, answer) : ok(token);
};
