import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const PREFIX = 'https://graph.microsoft.com/v1.0';

// Chat-scoped cursor URLs are re-signed to match the token their originating
// command uses, so pagination never switches identity mid-walk. Chat METADATA
// (`/me/chats` → list-chats, `/chats/{id}` → get-chat) needs the elevated
// M365ChatClient identity (`Chat.ReadBasic`). But `/chats/{id}/members`
// (list-chat-members) reads on the BASIC token via `ChatMember.Read` — see that
// command — so its page-2 cursor must NOT be elevated, else it would 401 once
// the login-only elevated token goes stale even though page 1 succeeded.
const requiresElevated = (path: string): boolean => {
  if (path.startsWith('/me/chats')) return true;
  if (path.startsWith('/chats/')) return !path.includes('/members');
  return false;
};

const schema = z.object({
  url: z
    .string()
    .min(1)
    .refine((v) => v.startsWith(`${PREFIX}/`), { message: `must be a Microsoft Graph v1.0 URL starting with ${PREFIX}/` }),
});

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = parsed.data.url.slice(PREFIX.length);
  return requiresElevated(path) ? graph.getElevated(path) : graph.get(path);
};

const meta: CommandMeta = {
  summary:
    'Fetch the next page of a paginated Graph response. Pass the cursor the previous command emitted — in text mode the `---` footer prints the whole ready-to-run command (`next: ask-marcel-office next-page --url \'<url>\'`), so copy the line as-is (the URL is single-quoted because it contains `$`); in JSON mode use the top-level `nextLink` field. Never reach into `data["@odata.nextLink"]`; the CLI strips that and surfaces it as a first-class envelope/footer field. Automatically signs `/me/chats` and `/chats/...` cursors with the M365ChatClient elevated token to match the chat-metadata commands.',
  category: 'meta',
  graphMethod: 'GET',
  graphPathTemplate: '{url}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/paging',
  options: [
    {
      name: 'url',
      key: 'url',
      required: true,
      description:
        "Full Graph v1.0 URL — copy the top-level `nextLink` field from the previous response (the CLI hoists Graph's `@odata.nextLink` out of `data` to envelope level). " +
        'Example: `https://graph.microsoft.com/v1.0/me/messages?$skiptoken=AKDsfg...`. ' +
        'Loop: keep calling until the response no longer contains `nextLink`. ' +
        'Also handles `deltaLink` (also hoisted) if you want to resume a delta query.',
      argumentHint: { kind: 'graphSubpath' },
    },
  ],
  example: "ask-marcel-office next-page --url 'https://graph.microsoft.com/v1.0/me/messages?$skip=10'",
  responseShape: 'same shape as the originating endpoint — `{ ok: true, data: { value: [...] }, nextLink: "..." }` with the cursor at envelope level.',
};

export { execute, meta, schema };
