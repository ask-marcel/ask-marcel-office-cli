import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { searchDriveItems } from './drive-item-search.ts';
import { formatZodError } from './format-zod-error.ts';
import { isoDateTimeField, RELATIVE_DATE_DESCRIPTION } from './iso-datetime-schema.ts';

/**
 * Every file changed since an instant, across everything the user can open, in
 * one sweep of the Microsoft Search index: no drive walk and no stored delta
 * token (the design chose stateless). The index takes a date, not an instant,
 * and may read it in the tenant's time zone, so the query asks from the UTC day
 * before, a superset in every zone, and the exact instant is then applied to the
 * hits. Newest first; every answer says what the index cannot promise.
 */

const MAX_PAGES = 10; // 10 × 200 = 2000 files, then `truncated: true`. Narrow with --query or a later --since.
const DAY_MS = 86_400_000;

const COVERAGE_NOTE = 'Built from the Microsoft Search index: a change from the last few minutes may not be indexed yet, and a library excluded from search is not listed.';

const schema = z.object({ since: isoDateTimeField, query: z.string().min(1).optional() });

const kqlDateBefore = (since: string): string => new Date(Date.parse(since) - DAY_MS).toISOString().slice(0, 10);

const modifiedAt = (item: unknown): number => Date.parse(String((item as { readonly lastModifiedDateTime?: unknown } | null)?.lastModifiedDateTime));

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { since, query } = parsed.data;
  const bound = `LastModifiedTime>=${kqlDateBefore(since)}`;
  const sweep = await searchDriveItems(graph, query === undefined ? bound : `(${query}) AND ${bound}`, MAX_PAGES);
  if (!sweep.ok) return sweep;
  const sinceMs = Date.parse(since);
  const value = sweep.value.value.filter((item) => modifiedAt(item) >= sinceMs).toSorted((a, b) => modifiedAt(b) - modifiedAt(a));
  return ok({ since, count: value.length, value, note: COVERAGE_NOTE, ...(sweep.value.truncated ? { truncated: true } : {}) });
};

const meta: CommandMeta = {
  summary:
    'List every file changed since a date across everything the signed-in user can open — their OneDrive, files shared with them, every SharePoint and Teams library — newest first, in one sweep of the Microsoft Search index (no per-drive walk, nothing stored between runs). Each hit is a driveItem with `lastModifiedDateTime`, `lastModifiedBy`, `webUrl` and the `id` + `parentReference.driveId` pair the download commands take. The index lags behind the newest saves and skips libraries excluded from search, which the answer says in its `note`. Narrow with `--query` (KQL, e.g. `filetype:docx`).',
  category: 'drive',
  graphMethod: 'POST',
  graphPathTemplate: '/search/query',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/search-query',
  scopesRequired: ['Files.Read', 'Sites.Read.All'],
  options: [
    {
      name: 'since',
      key: 'since',
      required: true,
      description: `Only files modified at or after this instant. ${RELATIVE_DATE_DESCRIPTION}`,
    },
    {
      name: 'query',
      key: 'query',
      required: false,
      description: 'Optional KQL added to the date bound, e.g. `filetype:docx`, `filename:budget`, or a `path:"https://contoso.sharepoint.com/sites/finance"` to stay in one site.',
    },
  ],
  example: "ask-marcel-office list-changed-files --since yesterday --query 'filetype:xlsx'",
  bodyTemplate:
    "{ requests: [{ entityTypes: ['driveItem'], query: { queryString: '({query}) AND LastModifiedTime>=<the UTC day before {since}>' }, from: <page*200>, size: 200 }] } — re-issued per page until `moreResultsAvailable` is false or the 10-page ceiling is hit",
  responseShape:
    '`{ since, count, value: [<Microsoft Graph driveItem resource: { id, name, webUrl, lastModifiedDateTime, lastModifiedBy, parentReference: { driveId }, … }>], note, truncated?: true }`, newest first, only files modified at or after `since`. `note` states the index limits. `truncated: true` means the sweep stopped at 2000 files or on a failed later page — narrow with `--query` or a later `--since`.',
};

export { execute, meta, schema };
