import { z } from 'zod';
import { ok } from '../../domain/result.ts';
import { buildListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { filterOutArchivedSites } from './filter-archived-sites.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ query: z.string().min(1) });
const built = buildListCommand((p) => `/sites?search=${p.query}`, baseSchema);
const { schema } = built;

const sitesOf = (body: unknown): ReadonlyArray<unknown> => {
  const value = (body as { value?: ReadonlyArray<unknown> } | null)?.value;
  return Array.isArray(value) ? value : [];
};

// Wrap the raw `/sites?search=` GET to drop archived sites (e.g. a departed user's
// auto-archived OneDrive, which 423s on access). The cap (~25 matches) never trips
// the probe ceiling, so only `archivedExcluded` / `archiveProbeErrors` can appear.
const execute: Command['execute'] = async (graph, params) => {
  const r = await built.execute(graph, params);
  if (!r.ok) return r;
  const filtered = await filterOutArchivedSites(graph, sitesOf(r.value));
  return ok({
    value: filtered.value,
    ...(filtered.archivedExcluded > 0 ? { archivedExcluded: filtered.archivedExcluded } : {}),
    ...(filtered.probeErrors > 0 ? { archiveProbeErrors: filtered.probeErrors } : {}),
  });
};

const meta: CommandMeta = {
  summary: 'Search the tenant for SharePoint sites whose display name or description matches a free-text query (returns up to 25).',
  category: 'sharepoint',
  graphMethod: 'GET',
  graphPathTemplate: '/sites?search={query}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/site-search',
  options: [
    {
      name: 'query',
      key: 'query',
      required: true,
      description: 'Free-text query. Matches site display name and description across the tenant.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel search-sharepoint-sites-by-name --query 'marketing'",
  responseShape:
    'collection of Microsoft Graph `site` resources under `value[]` (up to 25). Archived sites are excluded: each match is probed (`GET /sites/{id}?$select=…,siteCollection`) and dropped when Graph reports it archived or fails with `423 resourceLocked` (an auto-archived OneDrive). `archivedExcluded` (omitted when 0) counts the drops; `archiveProbeErrors` (omitted when 0) counts matches kept because their probe failed for an unrelated reason.',
  pagination: true,
};

export { execute, meta, schema };
