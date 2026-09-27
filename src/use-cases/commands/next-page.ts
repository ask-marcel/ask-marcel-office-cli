import { z } from 'zod';
import { err, map } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { topOnlyShape } from './odata-query.ts';
import { brandTenantId, tenantIdShape } from './tenant-option.ts';
import { withTaskLinks } from './todo-web-url.ts';

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
  ...tenantIdShape,
  ...topOnlyShape,
});

// A To Do page gets each task's web link, as the To Do listings give page one.
const TODO_TASKS = /^\/me\/todo\/lists\/[^/?]+\/tasks(?:\/delta)?(?:\?|$)/;

const TOP_REFUSED = '--top applies to a cursor read on the basic token (mail, files, To Do and the like); a chat or partner-tenant cursor keeps the page size it was given.';

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = parsed.data.url.slice(PREFIX.length);
  // A partner-tenant cursor is unreadable on any home token, so `--tenant-id`
  // wins over the elevated/basic split (a chat cursor never carries one). The
  // cursor URL itself has no tenant, which is why the flag must supply it;
  // without it page 2 would 401 with `invalidAudienceUri` even after page 1 ok.
  const { top, tenantId } = parsed.data;
  if (top !== undefined && (tenantId !== undefined || requiresElevated(path))) return err({ type: 'validation_error', message: TOP_REFUSED });
  if (tenantId !== undefined) {
    const branded = brandTenantId(tenantId);
    if (!branded.ok) return branded;
    return graph.getGuest(path, branded.value);
  }
  if (requiresElevated(path)) return graph.getElevated(path);
  // Graph honours a page-size preference only on the request that carries it: a
  // delta walk that asked for 100 gets 10 a page on every continuation without it.
  const page = await graph.get(path, top === undefined ? undefined : { Prefer: `odata.maxpagesize=${top}` });
  return TODO_TASKS.test(path) ? map(page, withTaskLinks) : page;
};

const meta: CommandMeta = {
  summary:
    'Fetch the next page of a paginated Graph response. Pass the cursor the previous command emitted — in text mode the `---` footer prints the whole ready-to-run command (`next: ask-marcel-office next-page --url \'<url>\'`), so copy the line as-is (the URL is single-quoted because it contains `$`); in JSON mode use the top-level `nextLink` field. Never reach into `data["@odata.nextLink"]`; the CLI strips that and surfaces it as a first-class envelope/footer field. Automatically signs `/me/chats` and `/chats/...` cursors with the M365ChatClient elevated token to match the chat-metadata commands. When the cursor came from a partner-tenant (guest) drive listing, pass the same `--tenant-id` you used on the originating command, since the cursor carries no tenant and without it page 2 fails with `invalidAudienceUri`. `--top` keeps a page size on the continuation (Graph drops it after page one otherwise), and a To Do task page gets the `webUrl` of each task, as the To Do listings give page one.',
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
    {
      name: 'tenant-id',
      key: 'tenantId',
      required: false,
      description:
        'Tenant GUID of a PARTNER tenant you are a guest in. Pass it only to continue a partner-tenant drive listing (a folder / `*-drive-item` cursor whose file lives in another tenant), using the same `--tenant-id` you gave the originating command (ultimately from `resolve-drive-share-link`). ' +
        'The page is then signed with a guest token for that tenant; without it a partner-tenant cursor 401s with `invalidAudienceUri` on page 2 even though page 1 succeeded. ' +
        'Omit it for your own tenant and for every `/me/...` and chat cursor (the normal case).',
    },
    {
      name: 'top',
      key: 'top',
      required: false,
      description:
        'Page size for this continuation (positive integer, ≤ 1000), sent as `Prefer: odata.maxpagesize=N`: Graph honours a page size only on the request that carries it, so a mail delta walk that asked for `--top 100` gets 10 a page on every continuation without it. Basic-token cursors only: a chat or partner-tenant cursor is refused with it.',
    },
  ],
  example: "ask-marcel-office next-page --url 'https://graph.microsoft.com/v1.0/me/messages?$skip=10'",
  responseShape: 'same shape as the originating endpoint — `{ ok: true, data: { value: [...] }, nextLink: "..." }` with the cursor at envelope level.',
};

export { execute, meta, schema };
