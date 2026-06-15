/*
 * Shared network-error formatting for HTTP adapters.
 *
 * The two-tier timeout, the human-readable labels, and the
 * `networkErrorMessage` string formatter are reused by every HTTP-port
 * adapter (graph-client, teams-client, future siblings). Per-adapter
 * `wrap...Error` helpers stay in their respective adapters because they
 * also construct the adapter-specific error variant (`GraphError`,
 * `TeamsError`, …) — only the string-level helpers belong here.
 *
 * Audit v1.0.0 — SharePoint PDF download timeout fix introduced the
 * two-tier split. Audit follow-up extracted the helpers into this
 * module so TeamsClient (chatsvcagg) and any future client adapter can
 * reuse the same formatting without depending on `graph-client.ts`.
 */

// JSON Graph / chatsvcagg calls return in seconds — 60s is the right budget
// to catch genuine hangs. Binary CDN body transfers and large chunked uploads
// scale with file size × network speed; 60s aborts mid-body for multi-MB
// files. The two-tier split gives the JSON tier fast-failure characteristics
// and the binary tier enough wall-clock to actually move bytes.
export const REQUEST_TIMEOUT_MS = 60_000;
export const BINARY_TRANSFER_TIMEOUT_MS = 5 * 60_000; // 5 min
export const REQUEST_TIMEOUT_LABEL = '60s';
export const BINARY_TRANSFER_TIMEOUT_LABEL = '5min';

// Audit v1.0.0 §2.5: bare `fetch failed` / `request timed out after 60s` had
// zero context about which URL or method failed — an LLM caller cannot
// decide whether to retry without that. Prepend the request label
// (`GET /me/messages`) so the error envelope always names the call site.
// Transient transport flakiness (single `fetch failed` on parallel
// invocations that succeed sequentially) is also called out so the LLM
// knows to retry.
export const networkErrorMessage = (e: unknown, label: string, timeoutLabel: string): string => {
  const base = (() => {
    if (e instanceof Error && e.name === 'TimeoutError') return `request timed out after ${timeoutLabel}`;
    if (e instanceof Error && e.name === 'AbortError') return 'request aborted';
    if (e instanceof Error) return e.message;
    if (typeof e === 'string') return e;
    return 'network request failed';
  })();
  return `${base} (${label}) — transient; retry once before treating as permanent`;
};

export type HttpMethod = 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';
export type TimeoutTier = 'json' | 'binary';

export const timeoutLabelFor = (tier: TimeoutTier): string => (tier === 'binary' ? BINARY_TRANSFER_TIMEOUT_LABEL : REQUEST_TIMEOUT_LABEL);
export const timeoutMsFor = (tier: TimeoutTier): number => (tier === 'binary' ? BINARY_TRANSFER_TIMEOUT_MS : REQUEST_TIMEOUT_MS);
