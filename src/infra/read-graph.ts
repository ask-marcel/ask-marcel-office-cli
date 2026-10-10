import type { Result } from '../domain/result.ts';
import type { TenantId } from '../domain/tenant-id.ts';
import type { GraphError } from './graph-request.ts';

/*
 * The graph a read command gets (package split, D9): every GET tier (basic,
 * elevated, guest, the chat substrate, bytes) and a POST that reaches only the
 * two endpoints that change nothing. It has no put, patch, delete or upload
 * session. The POST path is checked by its type and again at run time, so a
 * cast cannot turn a read into a write.
 */

// The two endpoints a read command POSTs to; both only answer a query.
const READ_ONLY_POST_PATHS = ['/search/query', '/me/calendar/getSchedule'] as const;

type ReadOnlyPostPath = (typeof READ_ONLY_POST_PATHS)[number];

type ReadGraph = {
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
  /**
   * A pasted chat image from Teams' media service (`*.asm.skype.com`,
   * `*.asyncgw.teams.microsoft.com`), read with the IC3 bearer, the only token it
   * accepts (probed 2026-09-27: 200 with it, 401 with every other). The URL comes
   * out of message content, so any other host, or plain http, is refused before a
   * request is made. Answers `{ contentType, size, base64 }`.
   */
  teamsChatMedia: (url: string) => Promise<Result<unknown, GraphError>>;
  /**
   * A query sent as a POST: a Microsoft Search query or a free/busy schedule.
   * No other path is accepted, by type or at run time.
   */
  post: (path: ReadOnlyPostPath, body: unknown) => Promise<Result<unknown, GraphError>>;
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
};

export { READ_ONLY_POST_PATHS };
export type { ReadGraph, ReadOnlyPostPath };
