import { z } from 'zod';
import { err } from '../../domain/result.ts';
import { buildListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ query: z.string().min(1) });
const inner = buildListCommand((p) => `/me/messages?$search="${p.query}"`, baseSchema);

// Audit v1.0.0 §B6: Graph rejects `$search` + `$filter` together with
// `SearchWithFilter` (not the previously documented `InvalidRestriction`).
// Reject the conflict client-side so the LLM gets a precise pointer to the
// alternative command instead of paying a 500ms round-trip for an opaque
// Graph code.
const execute: Command['execute'] = async (graph, params) => {
  if (typeof params['filter'] === 'string' && params['filter'].length > 0) {
    return err({
      type: 'validation_error',
      message:
        '--filter is incompatible with $search on /me/messages (Graph rejects the combination with `SearchWithFilter`). Use `list-mail-messages --filter ...` for OData filtering, or drop `--filter` here and rely on the KQL query string.',
    });
  }
  return inner.execute(graph, params);
};
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    "Search the signed-in user's entire Outlook mailbox using KQL or free text. Results are ranked by Graph relevance. Note: Graph does not allow `$search` and `$filter` together — the CLI rejects `--filter` client-side with a pointer to `list-mail-messages` (which supports OData filtering). For sorting, server-side `$orderby` is also not allowed with `$search`; use the relevance ranking Graph returns.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages?$search="{query}"',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-messages',
  options: [
    {
      name: 'query',
      key: 'query',
      required: true,
      description:
        'KQL or free-text query. Searches subject, body, sender, and recipients. ' + 'Examples: `Q3 budget`, `from:alice@contoso.com`, `subject:invoice received>=2026-01-01`.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel search-mail-messages --query 'from:alice subject:Q3'",
  responseShape: 'collection of Microsoft Graph `message` resources under `value[]`, ranked by relevance',
  pagination: true,
};

export { execute, meta, schema };
