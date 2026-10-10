import type { HttpMethod, TimeoutTier } from './network-error.ts';
import { networkErrorMessage, timeoutLabelFor } from './network-error.ts';

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

export { isAllowedFetchUrlHost, toBase64, wrapNetworkError };
export type { FetchFn, GraphError };
