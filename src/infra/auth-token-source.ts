import type { Result } from '../domain/result.ts';
import { err } from '../domain/result.ts';
import type { TeamsRegion } from '../domain/teams-region.ts';
import { teamsRegion } from '../domain/teams-region.ts';
import type { TokenError, TokenSource } from '../use-cases/ports/token-source.ts';
import type { AuthManager } from './auth.ts';

// The cached region is pasted into every substrate URL; one that is not a region
// name stops the request rather than steering it to another path.
const INVALID_REGION_MESSAGE =
  'The token cache names a Teams region that is not a region name, so no chat request was sent. Run `ask-marcel-office login --force` to capture the region again.';

// The in-process token source. Each call is the AuthManager call the Graph
// client made before the port existed, with the same arguments, so the ladder
// behaves as it did. The elevated getter gets no `awaitSignIn`: a background
// command fails fast rather than parking on a sign-in form.
export const createAuthManagerTokenSource = (auth: AuthManager): TokenSource => ({
  graphToken: (tier) => (tier === 'elevated' ? auth.getElevatedAccessToken() : auth.getAccessToken()),
  guestToken: (tenant) => auth.getGuestAccessToken(tenant),
  // A rejected substrate token is a dead one, so the ladder redeems past its cache.
  substrateToken: (tier, options) => {
    const ask = options?.rejected === undefined ? undefined : { ignoreCache: true };
    return tier === 'ic3' ? auth.getIc3AccessToken(ask) : auth.getChatsvcaggAccessToken(ask);
  },
  substrateRegion: async (): Promise<Result<TeamsRegion, TokenError>> => {
    const region = teamsRegion(await auth.getChatsvcaggRegion());
    return region.ok ? region : err({ type: 'auth_failed', message: INVALID_REGION_MESSAGE, code: region.error.type });
  },
});
