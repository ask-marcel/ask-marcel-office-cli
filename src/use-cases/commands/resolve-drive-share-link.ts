import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { detectSiblingResolver } from './link-shape.ts';
import { buildShareToken } from './sharepoint-link-extractor.ts';

// Resolver for Microsoft Graph's `/shares/{token}` endpoint. Takes any OneDrive /
// SharePoint sharing URL, encodes it to the `u!<base64url>` share token that
// [shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get) accepts, and
// fetches `/shares/{token}/driveItem` so the caller gets the file's driveId + itemId
// in ONE call — the two ids every downstream `*-drive-item` command
// (`download-drive-item-content`, `convert-drive-item-*`, `extract-drive-item-images`,
// `get-drive-item`) needs. A raw sharing URL carries no ids, so this is the only
// entry point into the drive-item family from a "Copy link" URL.
//
// Accepted host shapes:
//   - `*.sharepoint.com`           — tenant SharePoint share URLs (`:b:/s/site/...`)
//   - `*-my.sharepoint.com`        — personal OneDrive share URLs
//   - `1drv.ms`                    — Microsoft's short-link shortener
//
// Internal helper `buildShareToken` lives in `sharepoint-link-extractor.ts`
// because three other commands (extract-sharepoint-links-in-mail,
// convert-mail-attachment-to-markdown, convert-mail-attachment-to-pdf) also
// use it. Don't inline-duplicate it here.
const schema = z.object({
  url: z.url(),
});

const ACCEPTED_HOST_PATTERNS: ReadonlyArray<RegExp> = [
  /\.sharepoint\.com$/i, // tenant + personal (`*-my.sharepoint.com` ends with `.sharepoint.com`)
  /^1drv\.ms$/i,
];

type Resolved = {
  readonly shareToken: string;
  readonly graphPath: string;
  readonly originalUrl: string;
};

const isAcceptedHost = (hostname: string): boolean => ACCEPTED_HOST_PATTERNS.some((re) => re.test(hostname));

const parse = (raw: string): Resolved | null => {
  // No try/catch — Zod's `.url()` refinement on the input schema already
  // validated the URL format before this parser runs. Also satisfies the
  // atelier rule restricting try/catch to `src/infra/**`.
  const url = new URL(raw);
  if (!isAcceptedHost(url.hostname)) return null;
  const shareToken = buildShareToken(raw);
  return {
    shareToken,
    graphPath: `/shares/${shareToken}/driveItem`,
    originalUrl: raw,
  };
};

type DriveItemHit = {
  readonly id?: unknown;
  readonly name?: unknown;
  readonly webUrl?: unknown;
  readonly size?: unknown;
  readonly lastModifiedDateTime?: unknown;
  readonly parentReference?: { readonly driveId?: unknown };
};

const str = (value: unknown): string | undefined => (typeof value === 'string' ? value : undefined);
const num = (value: unknown): number | undefined => (typeof value === 'number' ? value : undefined);

// Hoist the two ids every downstream `*-drive-item` command needs — driveId from
// `parentReference`, itemId from the driveItem's own `id` — to the top level, plus a
// few confirmation fields and the reusable share token. Reads are `unknown`-typed
// because the Graph response shape is not trusted at this boundary.
const toResolvedItem = (raw: unknown, shareToken: string): Record<string, unknown> => {
  const item = raw as DriveItemHit;
  return {
    driveId: str(item.parentReference?.driveId),
    itemId: str(item.id),
    name: str(item.name),
    webUrl: str(item.webUrl),
    size: num(item.size),
    lastModifiedDateTime: str(item.lastModifiedDateTime),
    shareToken,
  };
};

/**
 * Graph's answer when a home-tenant token is pointed at a tenant the user is only
 * a guest in: `401 invalidAudienceUri: Invalid audience Uri
 * '00000003-0000-0ff1-ce00-000000000000'` — SharePoint Online's app id. Home
 * Graph cannot mint an SPO token for a foreign tenant, so no home tier can reach
 * it (verified live 2026-07-16: basic and elevated fail identically, with and
 * without `Prefer: redeemSharingLink`).
 *
 * Narrow deliberately. A 403 `accessDenied` is a link you simply may not read,
 * and a guest token would not help — retrying it would spend a refresh-token
 * redemption and turn a clear "no access" into a confusing tenant error.
 */
const isForeignTenantAudienceError = (error: GraphError): boolean => error.type === 'api_error' && error.status === 401 && error.code === 'invalidAudienceUri';

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  // v1.4.0 re-audit Nit 1 (outlook + teams gaps): an Outlook web URL or
  // a Teams `/l/message/...` link wrongly passed to resolve-drive-share-link
  // used to fall through to the generic "not a recognised sharing URL"
  // rejection. Detect them early and emit a cross-pointer so the LLM
  // lands on the right sibling resolver.
  const sibling = detectSiblingResolver(parsed.data.url);
  if (sibling === 'mail') {
    return err({
      type: 'validation_error',
      message: '--url looks like an Outlook mail message link, not a OneDrive / SharePoint sharing URL.',
      code: 'cli_reject_mail_link_on_drive_share_resolver',
    });
  }
  if (sibling === 'calendar') {
    return err({
      type: 'validation_error',
      message: '--url looks like an Outlook calendar item link, not a OneDrive / SharePoint sharing URL.',
      code: 'cli_reject_calendar_link_on_drive_share_resolver',
    });
  }
  if (sibling === 'teams') {
    return err({
      type: 'validation_error',
      message: '--url looks like a Teams message link, not a OneDrive / SharePoint sharing URL.',
      code: 'cli_reject_teams_link_on_drive_share_resolver',
    });
  }
  const resolved = parse(parsed.data.url);
  if (resolved === null) {
    return err({
      type: 'validation_error',
      message:
        '--url: not a recognised OneDrive / SharePoint sharing URL. Accepted hosts: `*.sharepoint.com` (tenant + personal OneDrive — `*-my.sharepoint.com` is a subdomain that matches), `1drv.ms` (Microsoft short link).',
    });
  }
  // Resolve the token to the actual driveItem so the caller gets driveId + itemId in
  // one call (basic token, Files.Read.All) instead of a manual second /shares fetch.
  const item = await graph.get(resolved.graphPath);
  if (item.ok) return ok(toResolvedItem(item.value, resolved.shareToken));
  if (!isForeignTenantAudienceError(item.error)) return item;

  // The link belongs to a tenant the signed-in user is only a GUEST in.
  //
  // We do not detect that up front: the command holds no home-tenant id to compare
  // the URL against, so the audience error IS the signal. That keeps the home path
  // free (no extra call, no plumbing) and makes the foreign path self-healing.
  //
  // The URL's host is the only thing that can name the tenant — driveId and itemId
  // carry no tenant at all, which is exactly why the follow-up `*-drive-item`
  // commands need the `tenantId` this returns.
  const tenant = await graph.discoverTenantId(new URL(parsed.data.url).hostname);
  if (!tenant.ok) {
    // Every GraphError variant carries `message`, so no narrowing is needed here.
    return err({
      type: 'api_error',
      status: 502,
      message: `--url points at a tenant you are only a guest in (Graph refused your home token for its SharePoint), and that tenant could not be identified: ${tenant.error.message}`,
      code: 'guest_tenant_unresolved',
    });
  }
  const guestItem = await graph.getGuest(resolved.graphPath, tenant.value);
  if (!guestItem.ok) return guestItem;
  return ok({ ...toResolvedItem(guestItem.value, resolved.shareToken), tenantId: tenant.value });
};

const meta: CommandMeta = {
  summary:
    'Resolve a OneDrive / SharePoint sharing URL (a "Copy link" address someone sent you) to the file it points at, returning `driveId` + `itemId` ready to feed `get-drive-item`, `download-drive-item-content`, `convert-drive-item-*`, `extract-drive-item-images`, and the rest of the `*-drive-item` family. It encodes the URL to the Graph `/shares/{token}` share token (`u!<base64url>` per [shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get)) and fetches `/shares/{token}/driveItem` in ONE call (basic token, `Files.Read.All`) — a raw sharing URL carries no ids, so this is the entry point into the drive-item family from a shared link. Accepts any `*.sharepoint.com` URL (tenant + `*-my.sharepoint.com` personal OneDrive) and Microsoft\'s short-link host `1drv.ms`. **Cross-tenant links work too**: when the URL belongs to a tenant you are only a GUEST in, your home token cannot read its SharePoint (Graph answers `invalidAudienceUri` — it cannot mint a SharePoint token for a foreign tenant), so this command identifies the owning tenant from the URL host and retries with a guest token automatically, then returns that tenant as `tenantId` — see `responseShape` for how to pass it on to the follow-up calls.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/shares/(u!<base64url> of {url})/driveItem',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/shares-get',
  options: [
    {
      name: 'url',
      key: 'url',
      required: true,
      description:
        'A OneDrive / SharePoint sharing URL — the address from the "Copy link" / "Share" action in the OneDrive or SharePoint UI. Examples: `https://contoso.sharepoint.com/:b:/s/sitename/EaB1cD...`, `https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/file.pdf`, `https://1drv.ms/b/s!AbCdEfGh...`. The CLI does not follow the redirect on `1drv.ms` links — the short URL itself is encoded as the share token (Graph resolves it on the server side).',
    },
  ],
  example: "ask-marcel-office resolve-drive-share-link --url 'https://contoso.sharepoint.com/:b:/s/team/EaB1cD2eF...?e=abc'",
  responseShape:
    "`{ driveId, itemId, name, webUrl, size, lastModifiedDateTime, shareToken, tenantId? }`. `driveId` (from the item's `parentReference`) + `itemId` feed every `*-drive-item` command directly — no second call. `shareToken` is the `u!<base64url>` form, kept for reuse against other `/shares/{token}/...` endpoints. **`tenantId` is present ONLY when the link belongs to a partner tenant you are a guest in** — when it is there, pass it as `--tenant-id` to every follow-up `*-drive-item` call (`download-drive-item-content`, `convert-drive-item-*`, `extract-drive-item-images`, `get-drive-item`, …), because `driveId`/`itemId` carry no tenant and those commands would otherwise fail with `invalidAudienceUri`. When `tenantId` is absent the file is in your own tenant and no flag is needed. Any other field is absent/`undefined` when the resolved driveItem omits it (e.g. `size` on a folder).",
};

export { execute, meta, schema };
