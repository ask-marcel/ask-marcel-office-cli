/*
 * Shared URL-shape detector used by the four `resolve-*` commands to
 * recognise when the caller passed a URL belonging to a different sibling.
 * Returns the sibling resolver name so each command can emit a precise
 * cross-pointer rejection (mirrors the existing mail↔calendar pair the
 * audit Alex-session §B baseline shipped — extended here to cover
 * teams + drive-share, closing the 5-gap matrix flagged in the
 * v1.4.0 re-audit Nit 1).
 *
 * Atelier note: callers MUST validate the URL through Zod's `z.url()`
 * BEFORE invoking these predicates. `new URL(raw)` would throw on a
 * malformed input, and try/catch is forbidden outside `src/infra/**` —
 * every resolver's schema validation already guarantees a valid URL by
 * the time the parse step runs.
 */

type SiblingResolver = 'mail' | 'calendar' | 'teams' | 'drive-share';

const OUTLOOK_HOSTS: ReadonlySet<string> = new Set(['outlook.office.com', 'outlook.office365.com', 'outlook.live.com']);

const TEAMS_MESSAGE_PREFIX = 'https://teams.microsoft.com/l/message/';

// Mirrors resolve-drive-share-link's ACCEPTED_HOST_PATTERNS. Tenant +
// personal OneDrive both end in `.sharepoint.com` (the personal variant
// is `*-my.sharepoint.com`, a subdomain); `1drv.ms` is Microsoft's
// short-link host.
const DRIVE_SHARE_HOST_PATTERNS: ReadonlyArray<RegExp> = [/\.sharepoint\.com$/i, /^1drv\.ms$/i];

const hostIsDriveShare = (hostname: string): boolean => DRIVE_SHARE_HOST_PATTERNS.some((re) => re.test(hostname));

const hostIsOutlook = (hostname: string): boolean => OUTLOOK_HOSTS.has(hostname.toLowerCase());

// Outlook URLs can be either mail or calendar. Calendar shape is either
// a `/calendar/...` pathname or an OWA URL carrying `?path=/calendar`.
// Anything else on an outlook host is treated as mail (matches the
// existing isMailLink / isCalendarLink helpers in the resolver files).
const isOutlookCalendarShape = (url: URL): boolean => {
  if (url.pathname.startsWith('/calendar/')) return true;
  const pathParam = url.searchParams.get('path');
  return pathParam !== null && pathParam.startsWith('/calendar');
};

/**
 * Classify a Zod-validated URL string by which sibling resolver it
 * belongs to. Returns `null` when the URL matches none of the known
 * resolver families (caller falls through to its own "unknown URL"
 * rejection message).
 *
 * Precondition: `raw` has been validated through `z.url()`. Calling
 * with an unvalidated string is a programmer error — `new URL(raw)`
 * will throw and the atelier rule forbids try/catch outside infra.
 */
export const detectSiblingResolver = (raw: string): SiblingResolver | null => {
  // Teams check FIRST — it's a pure prefix match and doesn't need URL parsing.
  if (raw.startsWith(TEAMS_MESSAGE_PREFIX)) return 'teams';
  const url = new URL(raw);
  if (hostIsDriveShare(url.hostname)) return 'drive-share';
  if (hostIsOutlook(url.hostname)) return isOutlookCalendarShape(url) ? 'calendar' : 'mail';
  return null;
};

export type { SiblingResolver };
