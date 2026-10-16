import type { Result } from '../domain/result.ts';
import type { TokenError, TokenSource } from '../use-cases/ports/token-source.ts';

/*
 * The second half of a token failure (docs/plans/2026-10-01-package-split.md,
 * "Token protocol"): auth owns the remedy, and the package that asked for the
 * token adds the commands of its own that need it, from its registry. Every
 * command needs the basic token, and a guest token is asked for by name, so only
 * the three secondary tokens get a list.
 */

/** The commands of the asking package that need each secondary token, derived from its registry. */
export type SecondaryTokenCommands = {
  readonly elevated: ReadonlyArray<string>;
  readonly chatsvcagg: ReadonlyArray<string>;
  readonly ic3: ReadonlyArray<string>;
};

export const NO_SECONDARY_TOKEN_COMMANDS: SecondaryTokenCommands = { elevated: [], chatsvcagg: [], ic3: [] };

// The "(Commands that need this token: …)" tail of a token-failure message,
// empty when the caller named no commands.
export const neededByNote = (names: ReadonlyArray<string>, token: string = 'this token'): string =>
  names.length === 0 ? '' : ` (Commands that need ${token}: ${names.join(', ')}.)`;

// The code of a missing secondary token, from the helper as from the in-process
// ladder. Only that failure gets the list, as in the ladder: a configuration
// failure (a helper path, a bad token in the environment) is not cured by
// knowing which commands need the token.
const SECONDARY_TOKEN_UNAVAILABLE = 'secondary_token_unavailable';

type AuthFailed = Extract<TokenError, { readonly type: 'auth_failed' }>;

const isMissingToken = (error: TokenError): error is AuthFailed => error.type === 'auth_failed' && error.code === SECONDARY_TOKEN_UNAVAILABLE;

const withNote = <T>(result: Result<T, TokenError>, names: ReadonlyArray<string>): Result<T, TokenError> =>
  !result.ok && isMissingToken(result.error) ? { ok: false, error: { ...result.error, message: `${result.error.message}${neededByNote(names)}` } } : result;

/** The source with the asking package's commands added to each secondary-token failure. */
export const withNeededBy = (source: TokenSource, neededBy: SecondaryTokenCommands): TokenSource => ({
  graphToken: async (tier, options) => {
    const result = await source.graphToken(tier, options);
    return tier === 'elevated' ? withNote(result, neededBy.elevated) : result;
  },
  guestToken: (tenant, options) => source.guestToken(tenant, options),
  substrateToken: async (tier, options) => withNote(await source.substrateToken(tier, options), neededBy[tier]),
  substrateRegion: async (tier) => withNote(await source.substrateRegion(tier), neededBy[tier]),
});
