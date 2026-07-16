import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { AuthManager } from '../infra/auth.ts';
import { decodeJwtPayload } from '../domain/jwt-utils.ts';
import type { TenantId } from '../domain/tenant-id.ts';
import { tenantId } from '../domain/tenant-id.ts';
import { spoHostToTenantDomain } from '../domain/utilities/spo-tenant.ts';
import { BINARY_TRANSFER_TIMEOUT_MS, REQUEST_TIMEOUT_MS, networkErrorMessage, timeoutLabelFor, type HttpMethod, type TimeoutTier } from './network-error.ts';

type GraphError =
  | { type: 'api_error'; status: number; message: string; code?: string; retryAfterSeconds?: number }
  | { type: 'auth_failed'; message: string; code?: string }
  | { type: 'network_error'; message: string; code?: string }
  | { type: 'validation_error'; message: string; code?: string };

type GraphClient = {
  /**
   * `extraHeaders` lets a caller add request headers Graph requires on
   * specific endpoints — currently the only documented use is
   * `Prefer: odata.maxpagesize=N` on the calendar/mail delta endpoints,
   * which reject `$top` as a query parameter. Auth + content-type are
   * always set internally.
   */
  get: (path: string, extraHeaders?: Record<string, string>) => Promise<Result<unknown, GraphError>>;
  /**
   * Same JSON-GET shape as `get`, but signs the request with the
   * elevated Graph token (M365ChatClient). Used by commands the Teams
   * web client token cannot reach — currently `list-chats` and
   * `get-chat`, which need `Chat.ReadBasic` (only present on the
   * elevated token).
   */
  getElevated: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * JSON-GET signed with a PARTNER tenant's guest token instead of the home
   * token. Required for any path that touches a tenant the user is only a guest
   * in: home-tenant Graph cannot mint a SharePoint token for a foreign tenant,
   * so those calls die at `401 invalidAudienceUri` no matter which home tier
   * signs them.
   *
   * Get the `tenantId` from `resolve-drive-share-link` (it discovers it from the
   * sharing URL) or from the caller's `--tenant-id`.
   */
  getGuest: (path: string, tenantId: TenantId) => Promise<Result<unknown, GraphError>>;
  /**
   * Binary twin of `getGuest`: follows the Graph 302 to the partner tenant's CDN
   * and returns the bytes. The `fetchUrl` allow-list already admits any
   * `*.sharepoint.com` / `*.svc.ms` host, so a partner tenant's download URL
   * needs no special casing.
   */
  getBinaryGuest: (path: string, tenantId: TenantId) => Promise<Result<unknown, GraphError>>;
  /**
   * Resolves a SharePoint host to the Entra tenant that owns it, via the tenant's
   * public OIDC discovery document. Unauthenticated: it asks "who owns this
   * host?", not "what may I read?".
   *
   * This is what makes a bare sharing URL enough to cross tenants — the URL
   * carries the host, the host names the tenant, and the tenant is the one thing
   * `driveId` + `itemId` do not tell you.
   */
  discoverTenantId: (spoHost: string) => Promise<Result<TenantId, GraphError>>;
  /**
   * JSON-GET against the Teams chat substrate (post-2026-05:
   * `teams.microsoft.com/api/csa/<region>/api/v{N}/...` — see
   * `gotcha_chatsvcagg_substrate_moved` in memory for the migration
   * away from `chatsvcagg.teams.microsoft.com`). Signs the request
   * with the chatsvcagg-audience bearer captured at login (same Teams
   * web client identity as `get`, different audience), and injects the
   * cached substrate region between the host and `path`. Used by
   * commands that need to read chat message BODIES, which the basic
   * Graph token cannot reach (`Chat.Read*` scopes are missing).
   *
   * `path` MUST start with `/api/v{N}/...` — the host + `/api/csa/<region>`
   * prefix are added by this client.
   */
  teamsChat: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * JSON-GET against the Teams IC3 chat-message substrate at
   * `teams.microsoft.com/api/chatsvc/<region>/v1/...`. Same host as
   * `teamsChat` but a DIFFERENT path prefix AND a different bearer
   * audience (`https://ic3.teams.office.com` instead of
   * `https://chatsvcagg.teams.microsoft.com`). The path supports
   * `syncState` + `startTime` pagination — unlocking arbitrary-depth
   * chat-history reads beyond the chatsvcagg 200-message cap (see
   * `gotcha_chatsvcagg_substrate_moved` in memory). Used by
   * `list-teams-chat-history`.
   *
   * `path` MUST start with `/v1/...` (e.g. `/v1/users/ME/conversations/{id}/messages?startTime=...`)
   * — the host + `/api/chatsvc/<region>` prefix are added here.
   */
  teamsChatIc3: (path: string) => Promise<Result<unknown, GraphError>>;
  post: (path: string, body: unknown) => Promise<Result<unknown, GraphError>>;
  patch: (path: string, body: unknown) => Promise<Result<unknown, GraphError>>;
  getBinary: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * Same shape as `getBinary` but signs the request with an "elevated"
   * Graph token (issued for an app on Microsoft's ODSP
   * `logicalPermissions` allow-list — e.g., M365ChatClient). Used by
   * the historical-version commands which the Teams web client token
   * cannot fetch (403 logicalPermissionAccessDenied).
   */
  getBinaryElevated: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * Auth-less fetch of an arbitrary URL whose host MUST be on the
   * Microsoft allow-list. Used to follow `@microsoft.graph.downloadUrl`
   * 302 redirects (CDN-signed URLs) that the format-conversion
   * commands sometimes get back from Graph instead of inline bytes.
   */
  fetchUrl: (url: string) => Promise<Result<unknown, GraphError>>;
  /**
   * Upload bytes to a drive item. `basePath` is the bare driveItem
   * path (e.g. `/me/drive/root:/.ask-marcel-temp/abc.rtf`) — `put()`
   * appends `:/content` for the simple ≤4 MiB sync path or
   * `:/createUploadSession` for the chunked-session path internally
   * based on `body.byteLength`. No upper file-size limit beyond the
   * user's OneDrive quota.
   */
  put: (basePath: string, body: Uint8Array, contentType?: string) => Promise<Result<unknown, GraphError>>;
  delete: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * Decode the cached basic Teams token's JWT and return its scopes /
   * audience / expiry. Used by the `scopes-check` self-test command so the
   * LLM can predict `accessDenied` instead of discovering it on the next
   * Graph call. No network IO — operates on the cached token only.
   */
  getCachedTokenInfo: () => Promise<Result<TokenInfo, GraphError>>;
};

/**
 * Decode-only status for one non-basic token tier: availability, remaining runway,
 * the scopes granted to that token (decoded from its `scp`), and how it refreshes
 * (`automatic` = rides the shared refresh token; `interactive` = elevated, needs a login).
 */
type TokenTierInfo = {
  readonly available: boolean;
  readonly expiresInSeconds: number | undefined;
  readonly scopes: ReadonlyArray<string>;
  readonly refresh: 'automatic' | 'interactive';
  /**
   * Present ONLY when `available` is `false`: a one-line, jargon-free reason the
   * tier is absent + how to restore it. Stops the empty `scopes: []` on a missing
   * token from reading as a bug. Omitted entirely when the token is available.
   */
  readonly reason?: string;
};

type TokenInfo = {
  readonly scopes: ReadonlyArray<string>;
  readonly audience: string | undefined;
  readonly expiresAt: string | undefined;
  /**
   * Seconds remaining until the cached token's `exp` claim — derived from
   * `expiresAt - now`. Negative when the token has already expired. Absent
   * when the JWT did not carry an `exp` claim. lets
   * an LLM decide pre-emptively to run `ask-marcel-office login` (re-auth typically
   * worth doing under ~5 minutes) without parsing the ISO string itself.
   */
  readonly expiresInSeconds: number | undefined;
  /**
   * Whether the *persisted* elevated (M365ChatClient) token — the one the
   * historical-version download / convert commands need — is present and still
   * usable, plus its raw seconds-to-expiry (`undefined` when absent). `available`
   * is `false` when the auth manager cannot introspect it. Lets `deep-scan`
   * preflight elevated access in a fresh process instead of turning every
   * version download into a `403`.
   */
  readonly elevated: TokenTierInfo;
  /**
   * The two Teams-chat substrate tokens (chatsvcagg / ic3), same `TokenTierInfo`
   * shape as `elevated`. Both self-heal from the shared refresh token (refresh:
   * automatic), so they are informational rather than a preflight gate.
   */
  readonly chatsvcagg: TokenTierInfo;
  readonly ic3: TokenTierInfo;
};

// Recovery hints attached to an UNAVAILABLE tier so a bare `scopes: []` on a
// missing token reads as "not captured yet", not "this token has no scopes".
const TIER_REASON_INTERACTIVE = 'not cached (absent or expired) — run `ask-marcel-office login --force` to re-capture it; the elevated token carries no refresh token of its own';
const TIER_REASON_AUTOMATIC = 'not cached — self-heals on the next Teams-chat command from the shared refresh token, or run `ask-marcel-office login --force`';

const buildTier = (info: Omit<TokenTierInfo, 'refresh' | 'reason'>, refresh: TokenTierInfo['refresh']): TokenTierInfo => {
  if (info.available) return { ...info, refresh };
  return { ...info, refresh, reason: refresh === 'interactive' ? TIER_REASON_INTERACTIVE : TIER_REASON_AUTOMATIC };
};

const ALLOWED_FETCH_URL_HOSTS: ReadonlyArray<RegExp> = [
  /\.sharepoint\.com$/i,
  /\.onedrive\.com$/i,
  /\.live\.com$/i,
  /\.officeapps\.live\.com$/i,
  /\.1drv\.com$/i,
  /^graph\.microsoft\.com$/i,
  /\.svc\.ms$/i,
];

const isAllowedFetchUrlHost = (host: string): boolean => ALLOWED_FETCH_URL_HOSTS.some((re) => re.test(host));

type FetchFn = (url: string, init?: RequestInit) => Promise<Response>;

// Two-tier timeout constants live in src/infra/network-error.ts (shared
// with the TeamsClient adapter). The chunk constants are GraphClient-
// specific so they stay here.
const SIMPLE_PUT_THRESHOLD = 4 * 1024 * 1024; // 4 MiB
const CHUNK_SIZE = 5 * 1024 * 1024; // 5 MiB — Graph requires multiples of 320 KiB; 5 MiB is 16 × 320 KiB

const isJson = (contentType: string | null): boolean => contentType !== null && contentType.toLowerCase().includes('application/json');

const isText = (contentType: string | null): boolean => {
  if (contentType === null) return false;
  const lower = contentType.toLowerCase();
  return lower.startsWith('text/') || lower.includes('+xml') || lower.includes('application/xml') || lower.includes('application/javascript');
};

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

// Collapses the per-catch boilerplate that previously repeated across 8 sites:
// each catch had to manually format the label and pick the right timeout-tier
// constant. Putting both pieces here makes the binary-vs-json choice explicit
// at every call site without leaking the timeout-label strings outwards.
const wrapNetworkError = (e: unknown, method: HttpMethod, label: string, tier: TimeoutTier): GraphError => ({
  type: 'network_error',
  message: networkErrorMessage(e, `${method} ${label}`, timeoutLabelFor(tier)),
});

type GraphErrorBody = {
  readonly error?: {
    readonly code?: string;
    readonly message?: string;
    // Microsoft Graph uses lowercase `innererror`; SharePoint streamContent
    // (the CDN that hosts /drives/{}/items/{}/versions/{}/content) uses
    // camelCase `innerError`. Tolerate both so the inner code survives.
    readonly innererror?: { readonly code?: string };
    readonly innerError?: { readonly code?: string };
  };
};

const emptyOnJsonFailure = (): GraphErrorBody => ({});

// Graph occasionally returns `{error: {code: "UnknownError",
// message: ""}}` as a transient backend glitch. Without context the LLM sees
// "UnknownError: " (or just "UnknownError") and has nothing to act on. Detect
// the empty-message case and rewrite to a clear "retry / capture" hint.
const looksEmpty = (s: string | undefined): boolean => s === undefined || s.trim() === '';

// Graph's `Missing scope permissions` 403 inlines the
// caller's entire granted-scope list (~30 scopes, 700+ chars) into the error
// message. The trailing "Scopes on the request 'X,Y,Z,...'" is noise — the
// LLM only needs the *required* scope name(s) to know what's missing.
// Strip the granted-list suffix and replace with a pointer at scopes-check.
const SCOPE_DUMP_PATTERN = /^(.*Missing scope permissions[^.]*\.\s*API requires one of '[^']+'\.)\s*Scopes on the request '[^']*'.*$/i;

const truncateScopeDump = (message: string): string => {
  const match = SCOPE_DUMP_PATTERN.exec(message);
  if (match === null) return message;
  return `${match[1]} Run \`ask-marcel-office scopes-check\` to see granted scopes, or \`ask-marcel-office help-json | jq '.commands[] | select(.name=="<cmd>") | .scopesRequired'\` to see what a given command requires.`;
};

// HTTP/2 servers (chatsvcagg, Kestrel-fronted Teams substrates) routinely
// answer non-2xx with content-length: 0 AND an empty statusText. With no JSON
// error body to format AND no statusText to fall back to, the previous
// implementation surfaced `message: ''` — the CLI then printed bare `error: `
// with nothing after, leaving the LLM consumer no signal to act on. Synthesize
// a `HTTP <status> @ <pathname>` line so the failure is at least diagnosable.
const synthesizeEmptyBodyMessage = (status: number, url: string): string =>
  `HTTP ${status} with no error body (path: ${new URL(url).pathname}; the endpoint may have moved — see the command's "best-effort" note in --help)`;

// when an `ErrorInvalidIdMalformed`
// happens against a `/mailFolders/` URL, surface a more specific code so
// the presenter's hint table can recommend the well-known folder names
// (inbox, sentitems, drafts, …) instead of the generic "use a list-*
// command" advice. The Graph error message itself doesn't carry the URL,
// so the URL → code suffix happens here at the infra boundary where the
// URL IS still in scope (`res.url` / `fallbackUrl`). Pattern is the same
// idea as `asSubstrateError` but for path-aware error refinement.
const contextualizeCode = (code: string | undefined, url: string): string | undefined => {
  if (code !== 'ErrorInvalidIdMalformed' && code !== 'InvalidIdMalformed') return code;
  if (!url.includes('/mailFolders/') && !url.includes('mailFolders%2F')) return code;
  return `${code}_mailFolders`;
};

// RFC 9110 `Retry-After`: Graph (and its Azure front-ends) answer 429 / 503
// with a delta-seconds integer naming how long to wait before retrying. We
// surface it as `retryAfterSeconds` so a caller orchestrating tenant-scale
// crawls can honor the server's interval instead of guessing a backoff (and
// risk being re-throttled). The alternate HTTP-date form is intentionally not
// parsed — it would need a clock and Graph does not use it for throttling;
// when the header is absent or non-numeric the field is omitted and the caller
// falls back to its own backoff. `0` is a valid "retry immediately" hint, so
// the guard on the value is `!== undefined`, never truthiness.
const parseRetryAfter = (header: string | null): number | undefined => {
  if (header === null) return undefined;
  const trimmed = header.trim();
  if (!/^\d+$/.test(trimmed)) return undefined;
  return Number(trimmed);
};

const apiErrorFrom = async (res: Response, fallbackUrl: string): Promise<GraphError> => {
  const retryAfterSeconds = parseRetryAfter(res.headers.get('retry-after'));
  const retry = retryAfterSeconds === undefined ? {} : { retryAfterSeconds };
  const errBody = (await res.json().catch(emptyOnJsonFailure)) as GraphErrorBody;
  const tag = errBody.error?.innererror?.code ?? errBody.error?.innerError?.code ?? errBody.error?.code;
  const message = errBody.error?.message;
  // surface the Graph error code as a structured field
  // so LLM consumers can branch on `errorCode === "itemNotFound"` etc.
  // instead of substring-matching the human message.
  const rawCode = typeof tag === 'string' && tag !== '' ? tag : undefined;
  const effectiveUrlForCode = res.url !== '' ? res.url : fallbackUrl;
  const code = contextualizeCode(rawCode, effectiveUrlForCode);

  if (typeof tag === 'string' && tag === 'UnknownError' && looksEmpty(message)) {
    return {
      type: 'api_error',
      status: res.status,
      message:
        'UnknownError: (Graph returned an empty error body — likely a transient backend glitch; retry once. If persistent, capture the failing request URL + body and report.)',
      code: 'UnknownError',
      ...retry,
    };
  }

  // Some Graph endpoints (Planner is the canonical case) return a non-empty
  // outer error block but with `code: ""`. The previous code would format
  // that as `: <message>` — leading colon, no prefix — which the v1.0.0 audit
  // §2.7 flagged as malformed. Only prepend the tag if it's actually a
  // non-empty string.
  if (typeof tag === 'string' && tag !== '' && typeof message === 'string') {
    return { type: 'api_error', status: res.status, message: truncateScopeDump(`${tag}: ${message}`), ...(code ? { code } : {}), ...retry };
  }
  // `res.url` is empty when the Response was constructed manually (Bun's
  // fakeFetch test pattern) — fall back to the URL the caller just hit.
  const effectiveUrl = res.url !== '' ? res.url : fallbackUrl;
  const pickFallback = (): string => {
    if (typeof message === 'string' && message !== '') return message;
    if (res.statusText !== '') return res.statusText;
    return synthesizeEmptyBodyMessage(res.status, effectiveUrl);
  };
  return { type: 'api_error', status: res.status, message: truncateScopeDump(pickFallback()), ...(code ? { code } : {}), ...retry };
};

/**
 * Tag an api_error returned from a Microsoft-internal chat substrate
 * (chatsvcagg `/api/csa/<region>/...` or IC3 `/api/chatsvc/<region>/...`)
 * with a `substrateHttp{status}_{substrate}` code so the presenter's hint
 * table can match it and add the "best-effort substrate may have moved"
 * actionable hint plus `source: "substrate"` classifier. Prior shape left
 * substrate errors with whatever code (or
 * none) Graph returned, indistinguishable from regular Graph errors and
 * without the experimental-substrate context an LLM needs to decide whether
 * to retry, switch substrates, or surface the failure to the user.
 *
 * Non-api_error inputs (auth_failed, validation_error, network errors) pass
 * through unchanged — those are upstream-of-substrate failures and the
 * existing classifier handles them.
 */
const asSubstrateError = (e: GraphError, substrate: 'chatsvcagg' | 'ic3'): GraphError => {
  if (e.type !== 'api_error') return e;
  return { ...e, code: `substrateHttp${e.status}_${substrate}` };
};

const createGraphClient = (auth: AuthManager, fetchFn: FetchFn = globalThis.fetch): GraphClient => {
  // All four token tiers (basic / elevated / chatsvcagg / ic3) share the same
  // "fetch a bearer, map an auth failure to a GraphError" shape — only the
  // AuthManager getter differs. One helper, four thin bindings.
  const authHeadersFrom = async (getToken: AuthManager['getAccessToken']): Promise<Result<{ Authorization: string }, GraphError>> => {
    const tokenResult = await getToken();
    if (!tokenResult.ok) {
      const msg = tokenResult.error.type === 'auth_cancelled' ? 'Auth cancelled' : tokenResult.error.message;
      // Carry the auth layer's machine-readable code (e.g. the secondary-token
      // fail-fast) through to the envelope's `errorCode` so an agent can branch
      // on it without substring-matching the message.
      const code = tokenResult.error.type === 'auth_failed' ? tokenResult.error.code : undefined;
      return err({ type: 'auth_failed', message: msg, ...(code ? { code } : {}) });
    }
    return ok({ Authorization: `Bearer ${tokenResult.value}` });
  };
  const authHeaders = (): Promise<Result<{ Authorization: string }, GraphError>> => authHeadersFrom(auth.getAccessToken);
  const elevatedAuthHeaders = (): Promise<Result<{ Authorization: string }, GraphError>> => authHeadersFrom(auth.getElevatedAccessToken);
  const chatsvcaggAuthHeaders = (): Promise<Result<{ Authorization: string }, GraphError>> => authHeadersFrom(auth.getChatsvcaggAccessToken);
  const ic3AuthHeaders = (): Promise<Result<{ Authorization: string }, GraphError>> => authHeadersFrom(auth.getIc3AccessToken);
  // Fifth binding, and the only parameterised one: the guest tier is per-tenant,
  // so the getter closes over which tenant is being asked.
  const guestAuthHeaders = (tenantId: TenantId): Promise<Result<{ Authorization: string }, GraphError>> => authHeadersFrom(() => auth.getGuestAccessToken(tenantId));

  const request = async (method: 'GET' | 'POST' | 'PATCH', path: string, body?: unknown, extraHeaders?: Record<string, string>): Promise<Result<unknown, GraphError>> => {
    const headers = await authHeaders();
    if (!headers.ok) return headers;

    const url = `https://graph.microsoft.com/v1.0${path}`;
    try {
      const res = await fetchFn(url, {
        method,
        headers: { ...headers.value, 'content-type': 'application/json', ...(extraHeaders ?? {}) },
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
        ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      });
      if (!res.ok) return err(await apiErrorFrom(res, url));
      return ok(await res.json());
    } catch (e: unknown) {
      return err(wrapNetworkError(e, method, path, 'json'));
    }
  };

  const getElevated = async (path: string): Promise<Result<unknown, GraphError>> => {
    const headers = await elevatedAuthHeaders();
    if (!headers.ok) return headers;
    const url = `https://graph.microsoft.com/v1.0${path}`;
    try {
      const res = await fetchFn(url, {
        method: 'GET',
        headers: { ...headers.value, 'content-type': 'application/json' },
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (!res.ok) return err(await apiErrorFrom(res, url));
      return ok(await res.json());
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'GET', `${path} (elevated)`, 'json'));
    }
  };

  const getGuest = async (path: string, tenantId: TenantId): Promise<Result<unknown, GraphError>> => {
    const headers = await guestAuthHeaders(tenantId);
    if (!headers.ok) return headers;
    const url = `https://graph.microsoft.com/v1.0${path}`;
    try {
      const res = await fetchFn(url, {
        method: 'GET',
        headers: { ...headers.value, 'content-type': 'application/json' },
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (!res.ok) return err(await apiErrorFrom(res, url));
      return ok(await res.json());
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'GET', `${path} (guest ${tenantId})`, 'json'));
    }
  };

  // Teams chat substrate. Same Teams web client identity as `get`, but the
  // bearer is issued for `chatsvcagg.teams.microsoft.com` (audience claim
  // only — the actual API now lives on `teams.microsoft.com/api/csa/<region>/`
  // since the 2026-05 substrate move). We piggy-back the captured bearer to
  // read chat message bodies — Graph's `Chat.Read*`-gated endpoints can't
  // reach them with the scopes the basic Teams token carries.
  const teamsChat = async (path: string): Promise<Result<unknown, GraphError>> => {
    const headers = await chatsvcaggAuthHeaders();
    if (!headers.ok) return headers;
    const region = await auth.getChatsvcaggRegion();
    const url = `https://teams.microsoft.com/api/csa/${region}${path}`;
    try {
      const res = await fetchFn(url, {
        method: 'GET',
        headers: { ...headers.value, accept: 'application/json' },
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (!res.ok) return err(asSubstrateError(await apiErrorFrom(res, url), 'chatsvcagg'));
      return ok(await res.json());
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'GET', `${path} (chatsvcagg)`, 'json'));
    }
  };

  // IC3 substrate — same host as teamsChat (teams.microsoft.com) but a
  // different path prefix (`/api/chatsvc/<region>/` vs `/api/csa/<region>/`)
  // and a different bearer audience (`https://ic3.teams.office.com` vs
  // `https://chatsvcagg.teams.microsoft.com`). The IC3 substrate is the one
  // Teams web actually uses for chat-message scrollback — it supports
  // `syncState` + `startTime` pagination that chatsvcagg lacks. See
  // `gotcha_chatsvcagg_substrate_moved` in memory for the discovery.
  const teamsChatIc3 = async (path: string): Promise<Result<unknown, GraphError>> => {
    const headers = await ic3AuthHeaders();
    if (!headers.ok) return headers;
    const region = await auth.getChatsvcaggRegion();
    const url = `https://teams.microsoft.com/api/chatsvc/${region}${path}`;
    try {
      const res = await fetchFn(url, {
        method: 'GET',
        headers: { ...headers.value, accept: 'application/json' },
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (!res.ok) return err(asSubstrateError(await apiErrorFrom(res, url), 'ic3'));
      return ok(await res.json());
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'GET', `${path} (ic3)`, 'json'));
    }
  };

  const getBinaryWith = async (path: string, signedHeaders: { Authorization: string }): Promise<Result<unknown, GraphError>> => {
    const url = `https://graph.microsoft.com/v1.0${path}`;
    try {
      const res = await fetchFn(url, {
        method: 'GET',
        headers: signedHeaders,
        redirect: 'manual',
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (res.status >= 300 && res.status < 400) {
        const location = res.headers.get('location');
        if (location !== null) return ok({ '@microsoft.graph.downloadUrl': location });
      }
      if (!res.ok) return err(await apiErrorFrom(res, url));
      const contentType = res.headers.get('content-type');
      if (isJson(contentType)) return ok(await res.json());
      if (isText(contentType)) {
        const text = await res.text();
        // `size` is documented as the byte count of the source. JS strings are
        // UTF-16 — `.length` counts code units, NOT UTF-8 bytes — so a file
        // with multi-byte chars (any non-ASCII) reported a `size` smaller than
        // the actual byte count an `--output-path` write produced.
        // Use the encoded byte length so envelope `size` matches the disk size.
        return ok({ contentType: contentType ?? 'text/plain', size: new TextEncoder().encode(text).byteLength, text });
      }
      const buffer = await res.arrayBuffer();
      return ok({ contentType: contentType ?? 'application/octet-stream', size: buffer.byteLength, base64: toBase64(new Uint8Array(buffer)) });
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'GET', `${path} (binary)`, 'json'));
    }
  };

  const getBinary = async (path: string): Promise<Result<unknown, GraphError>> => {
    const headers = await authHeaders();
    if (!headers.ok) return headers;
    return getBinaryWith(path, headers.value);
  };

  const getBinaryElevated = async (path: string): Promise<Result<unknown, GraphError>> => {
    const headers = await elevatedAuthHeaders();
    if (!headers.ok) return headers;
    return getBinaryWith(path, headers.value);
  };

  const getBinaryGuest = async (path: string, tenantId: TenantId): Promise<Result<unknown, GraphError>> => {
    const headers = await guestAuthHeaders(tenantId);
    if (!headers.ok) return headers;
    return getBinaryWith(path, headers.value);
  };

  /**
   * Ask Entra which tenant owns a SharePoint host, using its public OIDC
   * discovery document. No credentials: the question is "who owns this host?",
   * and the answer is public.
   *
   * A host outside the `*.sharepoint.com` convention, or a domain Entra does not
   * know, is not an error to retry — it means no partner tenant applies, and the
   * caller should stay on its home token. Both surface as a clear message rather
   * than a crash, because the host->onmicrosoft mapping is a convention and a
   * tenant with a vanity arrangement may not follow it.
   */
  const discoverTenantId = async (spoHost: string): Promise<Result<TenantId, GraphError>> => {
    const domain = spoHostToTenantDomain(spoHost);
    if (domain === null) {
      return err({
        type: 'api_error',
        status: 400,
        message: `${spoHost} is not a tenant SharePoint host, so no partner tenant can be resolved from it`,
        code: 'not_a_sharepoint_host',
      });
    }
    const url = `https://login.microsoftonline.com/${domain}/v2.0/.well-known/openid-configuration`;
    try {
      const res = await fetchFn(url, { method: 'GET', headers: { accept: 'application/json' }, signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
      if (!res.ok) {
        return err({
          type: 'api_error',
          status: res.status,
          message: `could not resolve a tenant for ${spoHost} (tried ${domain}) — the host may belong to a tenant whose sign-in domain differs from its SharePoint name`,
          code: 'tenant_discovery_failed',
        });
      }
      const issuer = (await res.json())['issuer'];
      // The tenant id is the first path segment of the issuer
      // (`https://login.microsoftonline.com/{tid}/v2.0`). Brand it: it becomes an
      // authority segment on a POST that carries the refresh token.
      const segment = typeof issuer === 'string' ? (new URL(issuer).pathname.split('/')[1] ?? '') : '';
      const branded = tenantId(segment);
      if (!branded.ok) return err({ type: 'api_error', status: 502, message: `tenant discovery for ${spoHost} returned an unusable issuer`, code: 'tenant_discovery_failed' });
      return ok(branded.value);
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'GET', `tenant discovery for ${spoHost}`, 'json'));
    }
  };

  const fetchUrl = async (url: string): Promise<Result<unknown, GraphError>> => {
    let host: string;
    try {
      host = new URL(url).host;
    } catch {
      return err({ type: 'network_error', message: `fetchUrl rejected: invalid URL ${url}` });
    }
    if (!isAllowedFetchUrlHost(host)) {
      return err({ type: 'network_error', message: `fetchUrl rejected: host ${host} not in Microsoft allow-list` });
    }

    try {
      const res = await fetchFn(url, {
        method: 'GET',
        headers: { accept: 'text/html, application/xhtml+xml, application/xml;q=0.9, */*;q=0.8' },
        signal: AbortSignal.timeout(BINARY_TRANSFER_TIMEOUT_MS),
      });
      if (!res.ok) return err(await apiErrorFrom(res, url));
      const contentType = res.headers.get('content-type');
      if (isJson(contentType)) return ok(await res.json());
      if (isText(contentType)) {
        const text = await res.text();
        // `size` is documented as the byte count of the source. JS strings are
        // UTF-16 — `.length` counts code units, NOT UTF-8 bytes — so a file
        // with multi-byte chars (any non-ASCII) reported a `size` smaller than
        // the actual byte count an `--output-path` write produced.
        // Use the encoded byte length so envelope `size` matches the disk size.
        return ok({ contentType: contentType ?? 'text/plain', size: new TextEncoder().encode(text).byteLength, text });
      }
      const buffer = await res.arrayBuffer();
      return ok({ contentType: contentType ?? 'application/octet-stream', size: buffer.byteLength, base64: toBase64(new Uint8Array(buffer)) });
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'GET', `${url} (CDN follow)`, 'binary'));
    }
  };

  const simplePut = async (path: string, body: Uint8Array, contentType?: string): Promise<Result<unknown, GraphError>> => {
    const headers = await authHeaders();
    if (!headers.ok) return headers;
    const url = `https://graph.microsoft.com/v1.0${path}`;
    try {
      const res = await fetchFn(url, {
        method: 'PUT',
        headers: { ...headers.value, 'content-type': contentType ?? 'application/octet-stream' },
        body: body as unknown as BodyInit,
        signal: AbortSignal.timeout(BINARY_TRANSFER_TIMEOUT_MS),
      });
      if (!res.ok) return err(await apiErrorFrom(res, url));
      return ok(await res.json());
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'PUT', path, 'binary'));
    }
  };

  const chunkedPut = async (basePath: string, body: Uint8Array): Promise<Result<unknown, GraphError>> => {
    // 1. Create upload session via the authenticated request() helper.
    const session = await request('POST', `${basePath}:/createUploadSession`, {
      item: { '@microsoft.graph.conflictBehavior': 'replace' },
    });
    if (!session.ok) return session;
    const uploadUrl = (session.value as { uploadUrl?: string }).uploadUrl;
    if (typeof uploadUrl !== 'string') {
      return err({ type: 'api_error', status: 500, message: 'createUploadSession returned no uploadUrl' });
    }

    // Hardening #3: validate the host before any chunk PUT.
    let host: string;
    try {
      host = new URL(uploadUrl).host;
    } catch {
      return err({ type: 'network_error', message: 'createUploadSession returned an invalid uploadUrl' });
    }
    if (!isAllowedFetchUrlHost(host)) {
      return err({ type: 'network_error', message: `uploadUrl host ${host} not in Microsoft allow-list` });
    }

    // 2. PUT chunks to the pre-signed upload URL — no auth header.
    const total = body.byteLength;
    for (let start = 0; start < total; start += CHUNK_SIZE) {
      const end = Math.min(start + CHUNK_SIZE, total) - 1;
      const chunk = body.slice(start, end + 1);
      try {
        const res = await fetchFn(uploadUrl, {
          method: 'PUT',
          headers: { 'Content-Range': `bytes ${start}-${end}/${total}` },
          body: chunk as unknown as BodyInit,
          signal: AbortSignal.timeout(BINARY_TRANSFER_TIMEOUT_MS),
        });
        if (!res.ok) {
          // Best-effort session cancellation; ignore failure. DELETE keeps
          // the short-tier budget — it's a Graph-side cleanup that should
          // return promptly.
          try {
            await fetchFn(uploadUrl, { method: 'DELETE', signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
          } catch {
            /* ignore */
          }
          return err({ type: 'api_error', status: res.status, message: `chunk PUT failed at byte ${start}` });
        }
        if (res.status === 200 || res.status === 201) {
          return ok(await res.json());
        }
        // 202 Accepted — continue uploading.
      } catch (e: unknown) {
        try {
          await fetchFn(uploadUrl, { method: 'DELETE', signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
        } catch {
          /* ignore */
        }
        return err(wrapNetworkError(e, 'PUT', `chunk @ byte ${start}`, 'binary'));
      }
    }
    return err({ type: 'api_error', status: 500, message: 'chunked upload completed without final response' });
  };

  const put = async (basePath: string, body: Uint8Array, contentType?: string): Promise<Result<unknown, GraphError>> => {
    if (body.byteLength <= SIMPLE_PUT_THRESHOLD) {
      return simplePut(`${basePath}:/content`, body, contentType);
    }
    return chunkedPut(basePath, body);
  };

  const deleteResource = async (path: string): Promise<Result<unknown, GraphError>> => {
    const headers = await authHeaders();
    if (!headers.ok) return headers;
    const url = `https://graph.microsoft.com/v1.0${path}`;
    try {
      const res = await fetchFn(url, {
        method: 'DELETE',
        headers: headers.value,
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (!res.ok) return err(await apiErrorFrom(res, url));
      return ok(undefined);
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'DELETE', path, 'json'));
    }
  };

  const getCachedTokenInfo = async (): Promise<Result<TokenInfo, GraphError>> => {
    const tokenResult = await auth.getAccessToken();
    if (!tokenResult.ok) {
      const msg = tokenResult.error.type === 'auth_cancelled' ? 'Auth cancelled' : tokenResult.error.message;
      return err({ type: 'auth_failed', message: msg });
    }
    const claims = decodeJwtPayload(tokenResult.value);
    const scpRaw = claims['scp'];
    const scopes = typeof scpRaw === 'string' ? scpRaw.split(' ').filter((s) => s.length > 0) : [];
    const audRaw = claims['aud'];
    const audience = typeof audRaw === 'string' ? audRaw : undefined;
    const expRaw = claims['exp'];
    const expiresAt = typeof expRaw === 'number' ? new Date(expRaw * 1000).toISOString() : undefined;
    const expiresInSeconds = typeof expRaw === 'number' ? Math.floor(expRaw - Date.now() / 1000) : undefined;
    // Decode-only elevated preflight: the real AuthManager reads its persisted
    // elevated token; a minimal one omits the capability and we report unavailable.
    const noTier = { available: false, expiresInSeconds: undefined, scopes: [] };
    const elevatedInfo = auth.getCachedElevatedInfo ? await auth.getCachedElevatedInfo() : noTier;
    const chatsvcaggInfo = auth.getCachedChatsvcaggInfo ? await auth.getCachedChatsvcaggInfo() : noTier;
    const ic3Info = auth.getCachedIc3Info ? await auth.getCachedIc3Info() : noTier;
    // The refresh route is a fixed per-tier property: the substrate + elevated tokens
    // that self-heal from the shared RT are `automatic`; the elevated (M365) token has
    // no refresh token of its own, so it is `interactive` (a browser login re-captures it).
    const elevated = buildTier(elevatedInfo, 'interactive');
    const chatsvcagg = buildTier(chatsvcaggInfo, 'automatic');
    const ic3 = buildTier(ic3Info, 'automatic');
    return ok({ scopes, audience, expiresAt, expiresInSeconds, elevated, chatsvcagg, ic3 });
  };

  return {
    get: (path, extraHeaders) => request('GET', path, undefined, extraHeaders),
    getElevated,
    getGuest,
    discoverTenantId,
    teamsChat,
    teamsChatIc3,
    post: (path, body) => request('POST', path, body),
    patch: (path, body) => request('PATCH', path, body),
    getBinary,
    getBinaryElevated,
    getBinaryGuest,
    fetchUrl,
    put,
    delete: deleteResource,
    getCachedTokenInfo,
  };
};

export { createGraphClient };
export type { FetchFn, GraphClient, GraphError, TokenInfo };
