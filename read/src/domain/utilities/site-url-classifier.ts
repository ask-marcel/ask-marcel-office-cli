/**
 * Classify a SharePoint `webUrl` the Microsoft Search index returns but which is
 * NOT a navigable site collection — fetching it 404s. Three families show up:
 *   1. SharePoint Add-in (app) webs — isolated `<tenant>-<hex>.sharepoint.com`
 *      hosts. Add-ins are deprecated; the site redirects to
 *      `/_layouts/15/AddinDeprecationAnnoucement.aspx`.
 *   2. SharePoint Embedded containers — `/contentstorage/...`, backing storage
 *      for apps (Loop / Designer / Copilot), never a navigable site collection.
 *   3. System pages — any `/_layouts/...` URL is a page, not a site root.
 *
 * Detection is by URL shape alone (no Graph call), so these are dropped from the
 * site listing before the per-site archive probe even runs.
 */

const NON_NAVIGABLE_PATTERNS: ReadonlyArray<RegExp> = [
  /^https?:\/\/[^/]*-[0-9a-f]{10,}\.sharepoint\.com(?:[:/]|$)/i, // add-in / app isolated host
  /\/contentstorage\//i, // SharePoint Embedded container
  /\/_layouts\//i, // system page, not a site root
];

const isNonNavigableSiteUrl = (webUrl: string): boolean => NON_NAVIGABLE_PATTERNS.some((pattern) => pattern.test(webUrl));

export { isNonNavigableSiteUrl };
