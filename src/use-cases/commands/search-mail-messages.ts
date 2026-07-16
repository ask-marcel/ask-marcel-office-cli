import { z } from 'zod';
import { err } from '../../domain/result.ts';
import { buildListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { MAIL_MESSAGE_DEFAULT_SELECT } from './mail-message-select.ts';
import { odataQueryOptions } from './odata-query.ts';
import { kqlSearchClause } from './search-escape.ts';

const baseSchema = z.object({ query: z.string().min(1) });
// Slim default projection shared with list-mail-messages / get-mail-message
// (see mail-message-select.ts); a user-supplied `--select` always wins.
const inner = buildListCommand((p) => `/me/messages?${kqlSearchClause(p.query)}`, baseSchema, { defaultSelect: MAIL_MESSAGE_DEFAULT_SELECT });

// Graph rejects `$search` + `$filter` together with
// `SearchWithFilter` (not the previously documented `InvalidRestriction`).
// Reject the conflict client-side so the LLM gets a precise pointer to the
// alternative command instead of paying a 500ms round-trip for an opaque
// Graph code.
const execute: Command['execute'] = async (graph, params) => {
  if (typeof params['filter'] === 'string' && params['filter'].length > 0) {
    // short `error` headline, with the
    // actionable remedy carried by the matching `hint` rule in
    // src/presenter/error-hints.ts (matched by `code`). Prior shape packed
    // both diagnosis AND remedy into `message`, leaving the generic
    // validation hint as boilerplate noise.
    return err({
      type: 'validation_error',
      message: '--filter is incompatible with $search on /me/messages — Graph rejects the combination with `SearchWithFilter`.',
      code: 'cli_reject_search_with_filter',
    });
  }
  return inner.execute(graph, params);
};
const { schema } = inner;

const meta: CommandMeta = {
  summary:
    "Search the signed-in user's entire Outlook mailbox using KQL or free text. Results are ranked by Graph relevance. The CLI ships a slim default `--select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview,conversationId` (same as `list-mail-messages`; `conversationId` is included so you can group hits into a thread or feed one to `list-conversation-messages`) so a 3-result page stays ~3 KB instead of ~30 KB. Pass `--select id,subject,body` to widen, or override entirely. Note: Graph does not allow `$search` and `$filter` together — the CLI rejects `--filter` client-side with a pointer to `list-mail-messages` (which supports OData filtering). For sorting, server-side `$orderby` is also not allowed with `$search`; use the relevance ranking Graph returns. **Exact-phrase search works**: `--query '\"budget allocation\"'` and embedded field phrases like `--query 'subject:\"Contoso A2 & B7 timeline\"'` are supported — the CLI escapes your double quotes into KQL phrase quotes, wraps the whole expression in the `\"…\"` Graph requires, and percent-encodes the value so `&`, `#`, and `+` are wire-safe. Pass raw KQL otherwise, e.g. `--query 'subject:invoice from:alice'`.",
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
  example: "ask-marcel-office search-mail-messages --query 'from:alice subject:Q3'",
  responseShape:
    'collection of Microsoft Graph `message` resources under `value[]`, ranked by relevance, each projected to the default `--select` set (or the requested fields when overridden). The default omits `body`, `internetMessageHeaders`, and `uniqueBody`.',
  pagination: true,
};

export { execute, meta, schema };
