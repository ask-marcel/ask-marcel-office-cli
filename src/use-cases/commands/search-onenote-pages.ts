import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';

// The path already pins `?$filter=contains(title,...)` for the title-substring
// match Graph requires. A user-supplied --filter would land Graph in
// double-$filter territory (BadRequest), so the OData passthrough here
// excludes --filter; the other five passthroughs (top/skip/select/orderby/
// expand) are honored.
const noFilterShape = Object.fromEntries(Object.entries(odataQuerySchema.shape).filter(([key]) => key !== 'filter')) as Omit<typeof odataQuerySchema.shape, 'filter'>;
const noFilterOptions = odataQueryOptions.filter((o) => o.name !== 'filter');

const schema = z.object({ titleSubstring: z.string().min(1) }).extend(noFilterShape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData(`/me/onenote/pages?$filter=contains(title,'${parsed.data.titleSubstring}')`, parsed.data);
  return graph.get(path);
};

const meta: CommandMeta = {
  summary:
    'Find OneNote pages whose title contains a substring (case-sensitive — page content is NOT searched). Microsoft removed full-text OneNote `?search=` from v1.0 Graph; only $filter against `title` remains, which is what this command runs. Accepts the OData passthrough flags top/skip/select/orderby/expand. The filter passthrough is intentionally omitted — the path already pins a `$filter` for the title-contains predicate, and Graph rejects two `$filter` query params.',
  category: 'notes',
  graphMethod: 'GET',
  graphPathTemplate: "/me/onenote/pages?$filter=contains(title,'{title-substring}')",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/onenote-list-pages',
  options: [
    {
      name: 'title-substring',
      key: 'titleSubstring',
      required: true,
      description:
        'Substring to look for inside OneNote page titles (case-sensitive, exact substring). ' +
        'This is title-only — full-text body search is not available on OneNote pages in v1.0 Graph. ' +
        'Use `list-onenote-section-pages` if you already know the section.',
      aliases: [{ name: 'query', key: 'query' }],
    },
    ...noFilterOptions,
  ],
  example: "ask-marcel search-onenote-pages --title-substring 'meeting notes' --top 25",
  responseShape: 'collection of Microsoft Graph `onenotePage` resources under `value[]` whose title contains the substring',
  pagination: true,
};

export { execute, meta, schema };
