import type { AccessToken } from '../../domain/access-token.ts';
import type { Result } from '../../domain/result.ts';
import type { TenantId } from '../../domain/tenant-id.ts';
import type { TokenError } from './token-source.ts';

// The five kinds of bearer the `token` helper hands out.
type TokenTier = 'basic' | 'elevated' | 'chatsvcagg' | 'ic3' | 'guest';

// A guest token is issued by a partner tenant's authority, so it names one.
type TokenRequest = { readonly tier: Exclude<TokenTier, 'guest'> } | { readonly tier: 'guest'; readonly tenant: TenantId };

// Where the `token` helper gets its bearers. Today the auth ladder
// (src/infra/auth.ts) supplies it, with the browser rungs its session allows.
type TokenIssuer = {
  // A token for the tier, by the ladder's usual route.
  readonly issueToken: (request: TokenRequest) => Promise<Result<AccessToken, TokenError>>;
  // The Teams region segment as the token cache holds it, unchecked. It fetches
  // no token.
  readonly cachedRegion: () => Promise<string>;
};

export type { TokenIssuer, TokenRequest, TokenTier };
