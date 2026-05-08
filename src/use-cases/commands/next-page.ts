import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const PREFIX = 'https://graph.microsoft.com/v1.0';

const ELEVATED_PATH_PREFIXES: ReadonlyArray<string> = ['/me/chats', '/chats/'];

const requiresElevated = (path: string): boolean => ELEVATED_PATH_PREFIXES.some((prefix) => path.startsWith(prefix));

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
    'Fetch the next page of a paginated Graph response. Pass the top-level `nextLink` value (NOT `data["@odata.nextLink"]` — the CLI promotes that field to envelope-level and strips it from data) returned by any list / search / delta command to walk pagination yourself. Automatically uses the elevated M365ChatClient token when the nextLink path starts with `/me/chats` or `/chats/...` so chat pagination follows the original auth context instead of 403ing on the basic Teams token.',
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
    },
  ],
  example: "ask-marcel next-page --url 'https://graph.microsoft.com/v1.0/me/messages?$skip=10'",
  responseShape: 'same shape as the originating endpoint — `{ ok: true, data: { value: [...] }, nextLink: "..." }` with the cursor at envelope level.',
};

export { execute, meta, schema };
