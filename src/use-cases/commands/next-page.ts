import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const PREFIX = 'https://graph.microsoft.com/v1.0';

const schema = z.object({
  url: z
    .string()
    .min(1)
    .refine((v) => v.startsWith(`${PREFIX}/`), { message: `must be a Microsoft Graph v1.0 URL starting with ${PREFIX}/` }),
});

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  return graph.get(parsed.data.url.slice(PREFIX.length));
};

const meta: CommandMeta = {
  summary:
    'Fetch the next page of a paginated Graph response. Pass the cursor the previous command emitted — in text mode that is the `next: <url>` value in the `---` footer; in JSON mode it is the top-level `nextLink` field. Never reach into `data["@odata.nextLink"]`; the CLI strips that and surfaces it as a first-class envelope/footer field.',
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
  example: "ask-marcel next-page --url 'https://graph.microsoft.com/v1.0/me/messages?$skip=10'",
  responseShape: 'same shape as the originating endpoint — `{ ok: true, data: { value: [...] }, nextLink: "..." }` with the cursor at envelope level.',
};

export { execute, meta, schema };
