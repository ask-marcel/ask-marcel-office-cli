import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ titleSubstring: z.string().min(1) });
const { execute } = buildCommand((p) => `/me/onenote/pages?$filter=contains(title,'${p.titleSubstring}')`, schema);

const meta: CommandMeta = {
  summary:
    'Find OneNote pages whose title contains a substring (case-sensitive — page content is NOT searched). Microsoft removed full-text OneNote `?search=` from v1.0 Graph; only $filter against `title` remains, which is what this command runs.',
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
  ],
  example: "ask-marcel search-onenote-pages --title-substring 'meeting notes'",
  responseShape: 'collection of Microsoft Graph `onenotePage` resources under `value[]` whose title contains the substring',
  pagination: true,
};

export { execute, meta, schema };
