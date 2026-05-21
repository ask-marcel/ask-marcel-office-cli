import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { AuthManager } from '../infra/auth.ts';
import { decodeJwtPayload } from '../domain/jwt-utils.ts';
import { BINARY_TRANSFER_TIMEOUT_MS, REQUEST_TIMEOUT_MS, networkErrorMessage, timeoutLabelFor, type HttpMethod, type TimeoutTier } from './network-error.ts';

type GraphError =
  | { type: 'api_error'; status: number; message: string; code?: string }
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
  post: (path: string, body: unknown) => Promise<Result<unknown, GraphError>>;
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

type TokenInfo = {
  readonly scopes: ReadonlyArray<string>;
  readonly audience: string | undefined;
  readonly expiresAt: string | undefined;
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

// Audit round-6 §1.2: Graph occasionally returns `{error: {code: "UnknownError",
// message: ""}}` as a transient backend glitch. Without context the LLM sees
// "UnknownError: " (or just "UnknownError") and has nothing to act on. Detect
// the empty-message case and rewrite to a clear "retry / capture" hint.
const looksEmpty = (s: string | undefined): boolean => s === undefined || s.trim() === '';

// Audit round-8 Wave F: Graph's `Missing scope permissions` 403 inlines the
// caller's entire granted-scope list (~30 scopes, 700+ chars) into the error
// message. The trailing "Scopes on the request 'X,Y,Z,...'" is noise — the
// LLM only needs the *required* scope name(s) to know what's missing.
// Strip the granted-list suffix and replace with a pointer at scopes-check.
const SCOPE_DUMP_PATTERN = /^(.*Missing scope permissions[^.]*\.\s*API requires one of '[^']+'\.)\s*Scopes on the request '[^']*'.*$/i;

const truncateScopeDump = (message: string): string => {
  const match = SCOPE_DUMP_PATTERN.exec(message);
  if (match === null) return message;
  return `${match[1]} Run \`ask-marcel scopes-check\` to see granted scopes, or \`ask-marcel help-json | jq '.commands[] | select(.name=="<cmd>") | .scopesRequired'\` to see what a given command requires.`;
};

// HTTP/2 servers (chatsvcagg, Kestrel-fronted Teams substrates) routinely
// answer non-2xx with content-length: 0 AND an empty statusText. With no JSON
// error body to format AND no statusText to fall back to, the previous
// implementation surfaced `message: ''` — the CLI then printed bare `error: `
// with nothing after, leaving the LLM consumer no signal to act on. Synthesize
// a `HTTP <status> @ <pathname>` line so the failure is at least diagnosable.
const synthesizeEmptyBodyMessage = (status: number, url: string): string =>
  `HTTP ${status} with no error body (path: ${new URL(url).pathname}; the endpoint may have moved — see the command's "best-effort" note in --help)`;

const apiErrorFrom = async (res: Response, fallbackUrl: string): Promise<GraphError> => {
  const errBody = (await res.json().catch(emptyOnJsonFailure)) as GraphErrorBody;
  const tag = errBody.error?.innererror?.code ?? errBody.error?.innerError?.code ?? errBody.error?.code;
  const message = errBody.error?.message;
  // Audit round-7 Wave G: surface the Graph error code as a structured field
  // so LLM consumers can branch on `errorCode === "itemNotFound"` etc.
  // instead of substring-matching the human message.
  const code = typeof tag === 'string' && tag !== '' ? tag : undefined;

  if (typeof tag === 'string' && tag === 'UnknownError' && looksEmpty(message)) {
    return {
      type: 'api_error',
      status: res.status,
      message:
        'UnknownError: (Graph returned an empty error body — likely a transient backend glitch; retry once. If persistent, capture the failing request URL + body and report.)',
      code: 'UnknownError',
    };
  }

  // Some Graph endpoints (Planner is the canonical case) return a non-empty
  // outer error block but with `code: ""`. The previous code would format
  // that as `: <message>` — leading colon, no prefix — which the v1.0.0 audit
  // §2.7 flagged as malformed. Only prepend the tag if it's actually a
  // non-empty string.
  if (typeof tag === 'string' && tag !== '' && typeof message === 'string') {
    return { type: 'api_error', status: res.status, message: truncateScopeDump(`${tag}: ${message}`), ...(code ? { code } : {}) };
  }
  // `res.url` is empty when the Response was constructed manually (Bun's
  // fakeFetch test pattern) — fall back to the URL the caller just hit.
  const effectiveUrl = res.url !== '' ? res.url : fallbackUrl;
  const pickFallback = (): string => {
    if (typeof message === 'string' && message !== '') return message;
    if (res.statusText !== '') return res.statusText;
    return synthesizeEmptyBodyMessage(res.status, effectiveUrl);
  };
  return { type: 'api_error', status: res.status, message: truncateScopeDump(pickFallback()), ...(code ? { code } : {}) };
};

const createGraphClient = (auth: AuthManager, fetchFn: FetchFn = globalThis.fetch): GraphClient => {
  const authHeaders = async (): Promise<Result<{ Authorization: string }, GraphError>> => {
    const tokenResult = await auth.getAccessToken();
    if (!tokenResult.ok) {
      const msg = tokenResult.error.type === 'auth_cancelled' ? 'Auth cancelled' : tokenResult.error.message;
      return err({ type: 'auth_failed', message: msg });
    }
    return ok({ Authorization: `Bearer ${tokenResult.value}` });
  };

  const request = async (method: 'GET' | 'POST', path: string, body?: unknown, extraHeaders?: Record<string, string>): Promise<Result<unknown, GraphError>> => {
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

  const elevatedAuthHeaders = async (): Promise<Result<{ Authorization: string }, GraphError>> => {
    const tokenResult = await auth.getElevatedAccessToken();
    if (!tokenResult.ok) {
      const msg = tokenResult.error.type === 'auth_cancelled' ? 'Auth cancelled' : tokenResult.error.message;
      return err({ type: 'auth_failed', message: msg });
    }
    return ok({ Authorization: `Bearer ${tokenResult.value}` });
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

  // Teams chat substrate. Same Teams web client identity as `get`, but the
  // bearer is issued for `chatsvcagg.teams.microsoft.com` (audience claim
  // only — the actual API now lives on `teams.microsoft.com/api/csa/<region>/`
  // since the 2026-05 substrate move). We piggy-back the captured bearer to
  // read chat message bodies — Graph's `Chat.Read*`-gated endpoints can't
  // reach them with the scopes the basic Teams token carries.
  const chatsvcaggAuthHeaders = async (): Promise<Result<{ Authorization: string }, GraphError>> => {
    const tokenResult = await auth.getChatsvcaggAccessToken();
    if (!tokenResult.ok) {
      const msg = tokenResult.error.type === 'auth_cancelled' ? 'Auth cancelled' : tokenResult.error.message;
      return err({ type: 'auth_failed', message: msg });
    }
    return ok({ Authorization: `Bearer ${tokenResult.value}` });
  };

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
      if (!res.ok) return err(await apiErrorFrom(res, url));
      return ok(await res.json());
    } catch (e: unknown) {
      return err(wrapNetworkError(e, 'GET', `${path} (chatsvcagg)`, 'json'));
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
        // the actual byte count an `--output-path` write produced (audit §2.1).
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
        // the actual byte count an `--output-path` write produced (audit §2.1).
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
    return ok({ scopes, audience, expiresAt });
  };

  return {
    get: (path, extraHeaders) => request('GET', path, undefined, extraHeaders),
    getElevated,
    teamsChat,
    post: (path, body) => request('POST', path, body),
    getBinary,
    getBinaryElevated,
    fetchUrl,
    put,
    delete: deleteResource,
    getCachedTokenInfo,
  };
};

export { createGraphClient };
export type { FetchFn, GraphClient, GraphError, TokenInfo };
