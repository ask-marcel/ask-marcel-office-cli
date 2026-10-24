/**
 * Archive detection for Microsoft Graph `site` resources. A site collection can
 * be archived two ways the CLI must recognise:
 *   1. Graph returns the site with `siteCollection.archivalDetails.archiveStatus`
 *      set (the documented v1.0 mechanism for admin-archived team sites).
 *   2. Graph fails the metadata call itself — auto-archived OneDrives (unlicensed
 *      ≥93 days) come back as `423 resourceLocked` ("Access to this site has been
 *      blocked"), the same signal `list-accessible-drives` already treats as an
 *      unreachable, drop-silently site.
 */

const ARCHIVED_STATUSES = new Set(['recentlyArchived', 'fullyArchived', 'reactivating']);
const ARCHIVE_CODES = new Set(['resourcelocked', 'sitearchived']);
const ARCHIVE_MESSAGE = /archiv|has been blocked|sharepointerror/i;

const objField = (value: unknown, key: string): unknown => (value === null || typeof value !== 'object' ? undefined : (value as Record<string, unknown>)[key]);

const archiveStatusOf = (resource: unknown): string | undefined => {
  const status = objField(objField(objField(resource, 'siteCollection'), 'archivalDetails'), 'archiveStatus');
  return typeof status === 'string' ? status : undefined;
};

const isArchivedSite = (resource: unknown): boolean => ARCHIVED_STATUSES.has(archiveStatusOf(resource) ?? '');

const errorIndicatesArchived = (e: { readonly status?: number; readonly code?: string; readonly message?: string }): boolean =>
  e.status === 423 || ARCHIVE_CODES.has((e.code ?? '').toLowerCase()) || ARCHIVE_MESSAGE.test(e.message ?? '');

export { archiveStatusOf, errorIndicatesArchived, isArchivedSite };
