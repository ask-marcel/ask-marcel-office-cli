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

// Wrap the raw `/sites?search=` GET to drop sites the user cannot open: archived
// ones (e.g. a departed user's auto-archived OneDrive, which 423s), non-navigable
// URL shapes (add-in app domains, `/contentstorage/` containers, `/_layouts/`
// pages), and probes that 404. The cap (~25 matches) never trips the probe ceiling.
const execute: Command['execute'] = async (graph, params) => {
  const r = await built.execute(graph, params);
  if (!r.ok) return r;
  const filtered = await filterOutArchivedSites(graph, sitesOf(r.value));
  return ok({
    value: filtered.value,
    ...(filtered.archivedExcluded > 0 ? { archivedExcluded: filtered.archivedExcluded } : {}),
    ...(filtered.nonNavigableExcluded > 0 ? { nonNavigableExcluded: filtered.nonNavigableExcluded } : {}),
    ...(filtered.notFoundExcluded > 0 ? { notFoundExcluded: filtered.notFoundExcluded } : {}),
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
    'collection of Microsoft Graph `site` resources under `value[]` (up to 25). Sites you cannot open are excluded: `nonNavigableExcluded` drops add-in app domains, `/contentstorage/` (SharePoint Embedded) containers, and `/_layouts/` system URLs by URL shape (no probe); each remaining match is probed (`GET /sites/{id}?$select=…,siteCollection`) and `archivedExcluded` drops archived / `423 resourceLocked` sites while `notFoundExcluded` drops probes that 404. All three counters are omitted when 0. `archiveProbeErrors` (omitted when 0) counts matches kept because their probe failed for an unrelated reason. Active personal OneDrives are kept.',
  pagination: true,
};

export { execute, meta, schema };
