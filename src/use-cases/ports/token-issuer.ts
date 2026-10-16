import type { AccessToken } from '@ask-marcel/office-core/domain/access-token.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import type { TenantId } from '@ask-marcel/office-core/domain/tenant-id.ts';
import type { TokenFingerprint } from '@ask-marcel/office-core/domain/token-fingerprint.ts';
import type { TokenError, TokenTier } from '@ask-marcel/office-core/use-cases/ports/token-source.ts';

// A guest token is issued by a partner tenant's authority, so it names one.
type TokenRequest = { readonly tier: Exclude<TokenTier, 'guest'> } | { readonly tier: 'guest'; readonly tenant: TenantId };

// Where the `token` helper gets its bearers. Today the auth ladder
// (src/infra/auth.ts) supplies it, with the browser rungs its session allows.
type TokenIssuer = {
  // `rejected` names a token a service just refused. The issuer then never hands
  // that one back and never opens a browser: it returns a newer cached token, or
  // redeems the refresh token, or fails.
  readonly issueToken: (request: TokenRequest, rejected?: TokenFingerprint) => Promise<Result<AccessToken, TokenError>>;
  // The Teams region segment as the token cache holds it, unchecked. It fetches
  // no token. A cache that cannot be read gives no region, never a default.
  readonly cachedRegion: () => Promise<Result<string, TokenError>>;
};

export type { TokenIssuer, TokenRequest, TokenTier };
