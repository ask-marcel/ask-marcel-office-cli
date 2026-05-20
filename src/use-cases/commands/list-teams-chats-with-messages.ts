import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

// Microsoft Teams chat-aggregator endpoint. Returns the signed-in user's
// chats AND inlines the last N message bodies per chat in a single
// round-trip. This is the path Teams web/desktop uses to populate the
// chat sidebar — Graph's `Chat.Read*`-gated endpoints can't reach the
// message bodies with the scopes the CLI's basic Teams token carries,
// but the chatsvcagg-audience bearer captured at login DOES have access.
//
// Page-size cap empirically ~20 (Teams web uses 20). Larger values are
// silently capped server-side. Pagination via `_skipToken` in the
// response body — pass it back as `--skip-token` to fetch the next page.
const schema = z.object({
  pageSize: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
  skipToken: z.string().min(1).optional(),
});

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const pageSize = parsed.data.pageSize ?? '20';
  const skipToken = parsed.data.skipToken;
  const qs = new URLSearchParams({ isPaginationEnabled: 'true', pageSize });
  if (skipToken !== undefined) qs.set('skipToken', skipToken);
  return graph.teamsChat(`/api/v2/users/me/chats?${qs.toString()}`);
};

const meta: CommandMeta = {
  summary:
    "List the signed-in user's Microsoft Teams chats with the last few message bodies inlined per chat. Uses the chatsvcagg-audience bearer captured at login (same identity as the basic Teams token, different audience) — this is the path Teams web/desktop uses to populate the chat sidebar. **Best-effort, may break on Microsoft client updates**: the chatsvcagg surface is not part of the public Microsoft Graph API; Microsoft can change route shapes without notice. If the response shape looks different from what you expect, run `ask-marcel logout && ask-marcel login` to refresh the captured tokens. Pagination via `_skipToken` field in the response — pass it back as `--skip-token` on the next call. Caller scopes do NOT matter here; the chatsvcagg server gates access on the appid + identity, not on Graph scopes.",
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: 'https://chatsvcagg.teams.microsoft.com/api/v2/users/me/chats',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-list',
  options: [
    {
      name: 'page-size',
      key: 'pageSize',
      required: false,
      description:
        'Number of chats per page (positive integer; Teams web uses 20 by default). Server may silently cap larger values. Combine with `--skip-token` from a previous response to paginate.',
    },
    {
      name: 'skip-token',
      key: 'skipToken',
      required: false,
      description: "Opaque pagination cursor returned in the prior response's `_skipToken` field. Pass it back to fetch the next page; omit on the first call.",
    },
  ],
  example: 'ask-marcel list-teams-chats-with-messages --page-size 10',
  responseShape:
    'chatsvcagg aggregator envelope: `{ chats: [...], _skipToken?: string }`. Each chat carries `id`, `topic`, `chatType`, `lastUpdatedDateTime`, AND an inlined `messages: [...]` with the last few bodies — this is the high-value field the existing `list-chats` cannot return. **Schema is Microsoft-internal — fields may change without notice; treat the response as semi-structured.**',
};

export { execute, meta, schema };
