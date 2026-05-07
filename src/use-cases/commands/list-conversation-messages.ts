import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';

// Hardcoded `$filter=conversationId eq '...'` in the path means a
// user-supplied `--filter` would cause Graph to receive two `$filter` query
// params. Graph also rejects this filter combined with `$orderby`
// (`InefficientFilter` — `conversationId` is not a sortable index), so the
// orderby passthrough is also omitted. Expose top/skip/select/expand only.
const allowedShape = Object.fromEntries(Object.entries(odataQuerySchema.shape).filter(([key]) => key !== 'filter' && key !== 'orderby')) as Omit<
  typeof odataQuerySchema.shape,
  'filter' | 'orderby'
>;
const allowedOptions = odataQueryOptions.filter((o) => o.name !== 'filter' && o.name !== 'orderby');

const schema = z.object({ conversationId: z.string().min(1) }).extend(allowedShape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const escaped = parsed.data.conversationId.replace(/'/g, "''");
  const path = appendOData(`/me/messages?$filter=conversationId eq '${escaped}'`, parsed.data);
  return graph.get(path);
};

const meta: CommandMeta = {
  summary:
    "List every message in a single Outlook conversation (thread) using `$filter=conversationId eq '...'`. Reconstructs a complete thread regardless of which subject lines or folders the replies landed in. Accepts the OData passthrough flags top/skip/select/expand — the filter and orderby passthroughs are intentionally omitted (the path already pins a `$filter`, and Graph rejects this filter combined with `$orderby` as `InefficientFilter` since `conversationId` is not a sortable index). The caller can sort by `receivedDateTime` client-side. KQL `$search` does not index `conversationId`, so `$filter` is the only documented Graph idiom for whole-thread retrieval.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: "/me/messages?$filter=conversationId eq '{conversation-id}'",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-messages',
  options: [
    {
      name: 'conversation-id',
      key: 'conversationId',
      required: true,
      description: 'Outlook `conversationId` of any message in the thread (returned by every mail-listing command and by `get-mail-message`).',
    },
    ...allowedOptions,
  ],
  example: "ask-marcel list-conversation-messages --conversation-id 'AAQkAD...=' --top 5 --select id,subject,receivedDateTime",
  responseShape: 'collection of Microsoft Graph `message` resources under `value[]` (unordered)',
  pagination: true,
};

export { execute, meta, schema };
