import { bytesToBase64 } from '../../domain/utilities/base64.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';

/**
 * Shared helpers for the SharePoint-link-extraction commands
 * (`extract-sharepoint-links-in-mail`, `extract-sharepoint-links-in-documents`).
 *
 * `extractSharepointUrls(text)` finds every `https://*.sharepoint.com/...`
 * URL inside a string — an HTML mail body or the joined `Target` attributes
 * of an OOXML package's external relationships — and returns the
 * deduplicated list (both `<a href=...>` and bare-text occurrences).
 *
 * `buildShareToken(url)` encodes a URL for Graph's `/shares/{token}`
 * resolver per [shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get):
 * `u!` + base64url of the URL's UTF-8 bytes, with no padding. Encoding the
 * UTF-8 bytes (not the raw string) is what lets non-ASCII paths resolve — a
 * OneDrive URL with an accented or CJK filename otherwise mints a wrong token
 * (or throws, above U+00FF).
 *
 * `resolveSharepointUrls(graph, urls)` fans out one `/shares/{token}/driveItem`
 * resolve per URL (capped at `MAX_LINKS`, per-link errors captured in the
 * entry rather than failing the whole call) — the orchestration shared by
 * both commands.
 */

const SP_URL_PATTERN = /https:\/\/[\w-]+(?:\.[\w-]+)*\.sharepoint\.com[^\s"'<>)]*/gi;

const MAX_LINKS = 25; // Hardening #4: cap fan-out

/** Where a link points, read from its URL: someone's OneDrive, or a SharePoint site. */
type LinkLocation = { readonly kind: 'onedrive'; readonly owner: string } | { readonly kind: 'site'; readonly site: string };

type ResolvedLink = {
  readonly url: string;
  readonly driveId?: string;
  readonly itemId?: string;
  readonly name?: string;
  readonly webUrl?: string;
  readonly error?: string;
  readonly location?: LinkLocation;
  readonly hint?: string;
};

type ResolvedLinks = {
  readonly links: ReadonlyArray<ResolvedLink>;
  readonly truncated: boolean;
  readonly skippedCount: number;
};

const stripFragment = (url: string): string => {
  const hash = url.indexOf('#');
  return hash === -1 ? url : url.slice(0, hash);
};

const extractSharepointUrls = (htmlBody: string): ReadonlyArray<string> => {
  const matches = htmlBody.match(SP_URL_PATTERN) ?? [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (const raw of matches) {
    const cleaned = stripFragment(raw);
    if (seen.has(cleaned)) continue;
    seen.add(cleaned);
    out.push(cleaned);
  }
  return out;
};

const buildShareToken = (url: string): string => {
  const b64 = bytesToBase64(new TextEncoder().encode(url)).replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '');
  return `u!${b64}`;
};

// A link Graph refuses gives no owner, so the URL is read instead: a OneDrive lives
// at `-my.sharepoint.com/personal/<account>`, a site at `/sites/<name>` or
// `/teams/<name>`; a sharing link puts `/:x:/g/` or `/:w:/r/` in front, or names
// the site straight after `/:x:/s/` (a site) or `/:x:/t/` (a team).
const ONEDRIVE_URL = /-my\.sharepoint\.com\/(?::[a-z]+:\/[a-z]\/)?personal\/([^/?#]+)/i;
const SITE_URL = /\.sharepoint\.com\/(?::[a-z]+:\/[a-z]\/)?(sites|teams)\/([^/?#]+)/i;
const SHARED_SITE_URL = /\.sharepoint\.com\/:[a-z]+:\/([st])\/([^/?#]+)/i;

const locationOf = (url: string): LinkLocation | undefined => {
  const onedrive = ONEDRIVE_URL.exec(url);
  if (onedrive !== null) return { kind: 'onedrive', owner: onedrive[1] };
  const site = SITE_URL.exec(url);
  if (site !== null) return { kind: 'site', site: `${site[1].toLowerCase()}/${site[2]}` };
  const shared = SHARED_SITE_URL.exec(url);
  if (shared !== null) return { kind: 'site', site: `${shared[1].toLowerCase() === 's' ? 'sites' : 'teams'}/${shared[2]}` };
  return undefined;
};

const ASK_IN_BROWSER = 'open the link in a browser to send an access request.';

const accessHint = (location: LinkLocation | undefined): string => {
  if (location === undefined) return `Ask the file's owner for access, or ${ASK_IN_BROWSER}`;
  if (location.kind === 'onedrive')
    return `It sits in the OneDrive of ${location.owner} (a sign-in name with . and @ written as _), which has not been shared with you: ask the owner for access, or ${ASK_IN_BROWSER}`;
  return `It sits on the SharePoint site ${location.site}, which you cannot open: ask a site owner for access, or ${ASK_IN_BROWSER}`;
};

const failedLink = (url: string, error: GraphError): ResolvedLink => {
  const location = locationOf(url);
  const where = location === undefined ? {} : { location };
  if (error.type !== 'api_error') return { url, error: `${error.type}: ${error.message}`, ...where };
  return { url, error: error.message, ...where, ...(error.status === 403 ? { hint: accessHint(location) } : {}) };
};

const resolveOne = async (graph: GraphClient, url: string): Promise<ResolvedLink> => {
  const token = buildShareToken(url);
  const result = await graph.get(`/shares/${token}/driveItem`);
  if (!result.ok) return failedLink(url, result.error);
  const item = result.value as { id?: string; name?: string; webUrl?: string; parentReference?: { driveId?: string } };
  return {
    url,
    driveId: item.parentReference?.driveId,
    itemId: item.id,
    name: item.name,
    webUrl: item.webUrl,
  };
};

const resolveSharepointUrls = async (graph: GraphClient, urls: ReadonlyArray<string>): Promise<ResolvedLinks> => {
  const truncated = urls.length > MAX_LINKS;
  const skippedCount = truncated ? urls.length - MAX_LINKS : 0;
  const kept = truncated ? urls.slice(0, MAX_LINKS) : urls;
  const links = await Promise.all(kept.map((u) => resolveOne(graph, u)));
  return { links, truncated, skippedCount };
};

export { buildShareToken, extractSharepointUrls, resolveSharepointUrls };
export type { LinkLocation, ResolvedLink, ResolvedLinks };
