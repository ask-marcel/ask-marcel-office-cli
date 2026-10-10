import type { AccessToken } from '../domain/access-token.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { TenantId } from '../domain/tenant-id.ts';
import type { TokenError, TokenSource } from '../use-cases/ports/token-source.ts';
import type { HttpMethod, TimeoutTier } from './network-error.ts';
import { REQUEST_TIMEOUT_MS, networkErrorMessage, timeoutLabelFor } from './network-error.ts';

/*
 * The private request core the read graph (read-graph.ts) and the write graph
 * (write-graph.ts) are built on: error shaping, signing, the 401 replay, and
 * the GET shapes both use. It holds no method that changes the tenant, so a
 * build that imports only the read graph carries no write code (package split,
 * D9; scripts/check-read-bundle.ts checks it).
 */

type GraphError =
  | { type: 'api_error'; status: number; message: string; code?: string; retryAfterSeconds?: number }
  | { type: 'auth_failed'; message: string; code?: string }
  | { type: 'network_error'; message: string; code?: string }
  | { type: 'validation_error'; message: string; code?: string };

type FetchFn = (url: string, init?: RequestInit) => Promise<Response>;

// Who signs a Graph request: the Teams web client token, the elevated one, or
// a partner tenant's guest token.
type GraphTier = 'basic' | 'elevated' | { readonly guest: TenantId };

// A request `send` signs. Each attempt gets its own deadline, so a replay has
// the full budget.
type SignedRequest = {
  readonly method: HttpMethod;
  readonly headers?: Readonly<Record<string, string>>;
  readonly body?: BodyInit;
  readonly redirect?: RequestRedirect;
  readonly timeoutMs: number;
};

type JsonRequest = {
  readonly body?: unknown;
  readonly extraHeaders?: Readonly<Record<string, string>>;
  // How a success is read; JSON by default.
  readonly readBody?: (res: Response) => Promise<unknown>;
};

type GraphRequestCore = {
  // One request signed for the tier.
  readonly send: (tier: GraphTier, url: string, request: SignedRequest) => Promise<Result<Response, GraphError>>;
  // A JSON request to Graph v1.0: the body read from a success, or the error.
  readonly json: (tier: GraphTier, method: HttpMethod, path: string, request?: JsonRequest) => Promise<Result<unknown, GraphError>>;
};

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

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
// Strip the granted-list suffix and replace with a pointer at status.
const SCOPE_DUMP_PATTERN = /^(.*Missing scope permissions[^.]*\.\s*API requires one of '[^']+'\.)\s*Scopes on the request '[^']*'.*$/i;

const truncateScopeDump = (message: string): string => {
  const match = SCOPE_DUMP_PATTERN.exec(message);
  if (match === null) return message;
  return `${match[1]} Run \`ask-marcel-office status\` to see granted scopes, or \`ask-marcel-office help-json | jq '.commands[] | select(.name=="<cmd>") | .scopesRequired'\` to see what a given command requires.`;
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

// Carry the auth layer's machine-readable code (e.g. the secondary-token
// fail-fast) through to the envelope's `errorCode` so an agent can branch on
// it without substring-matching the message.
const asAuthFailure = (error: TokenError): GraphError => {
  const message = error.type === 'auth_cancelled' ? 'Auth cancelled' : error.message;
  const code = error.type === 'auth_failed' ? error.code : undefined;
  return { type: 'auth_failed', message, ...(code ? { code } : {}) };
};

// The network-error label names the tier a JSON request was signed for.
const tierLabel = (tier: GraphTier): string => {
  if (tier === 'basic') return '';
  return tier === 'elevated' ? ' (elevated)' : ` (guest ${tier.guest})`;
};

// Every bearer comes from the token source, which knows the tiers; the core
// only signs requests with them.
const createGraphRequestCore = (tokens: TokenSource, fetchFn: FetchFn): GraphRequestCore => {
  const tokenFor = (tier: GraphTier, rejected?: AccessToken): Promise<Result<AccessToken, TokenError>> => {
    if (tier === 'basic') return tokens.graphToken('basic', { rejected });
    if (tier === 'elevated') return tokens.graphToken('elevated');
    return tokens.guestToken(tier.guest, { rejected });
  };

  const sendWith = (token: AccessToken, url: string, request: SignedRequest): Promise<Response> =>
    fetchFn(url, {
      method: request.method,
      headers: { Authorization: `Bearer ${token}`, ...request.headers },
      signal: AbortSignal.timeout(request.timeoutMs),
      body: request.body,
      redirect: request.redirect,
    });

  const send = async (tier: GraphTier, url: string, request: SignedRequest): Promise<Result<Response, GraphError>> => {
    const token = await tokenFor(tier);
    if (!token.ok) return err(asAuthFailure(token.error));
    return ok(await sendWith(token.value, url, request));
  };

  const json = async (tier: GraphTier, method: HttpMethod, path: string, request: JsonRequest = {}): Promise<Result<unknown, GraphError>> => {
    const url = `${GRAPH_ROOT}${path}`;
    try {
      const sent = await send(tier, url, {
        method,
        headers: { 'content-type': 'application/json', ...request.extraHeaders },
        timeoutMs: REQUEST_TIMEOUT_MS,
        ...(request.body === undefined ? {} : { body: JSON.stringify(request.body) }),
      });
      if (!sent.ok) return sent;
      if (!sent.value.ok) return err(await apiErrorFrom(sent.value, url));
      return ok(await (request.readBody ?? ((res: Response): Promise<unknown> => res.json()))(sent.value));
    } catch (e: unknown) {
      return err(wrapNetworkError(e, method, `${path}${tierLabel(tier)}`, 'json'));
    }
  };

  return { send, json };
};

export { apiErrorFrom, asAuthFailure, createGraphRequestCore, isAllowedFetchUrlHost, toBase64, wrapNetworkError };
export type { FetchFn, GraphError, GraphRequestCore, GraphTier };
