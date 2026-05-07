import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, odataQueryOptions, odataQuerySchema } from './odata-query.ts';

const schema = z.object({}).extend(odataQuerySchema.shape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData('/me/events/delta()', parsed.data);
  const result = await graph.get(path);
  if (result.ok) return result;
  // Graph's calendar delta endpoint returns `UnknownError: ` (empty message) when
  // the call is made without a Prefer: odata.maxpagesize header AND without a
  // resumption deltaLink. The CLI doesn't pass custom headers today, so the
  // best we can do is augment the empty-message case with actionable guidance.
  if (result.error.type === 'api_error' && result.error.message.trim() === 'UnknownError:') {
    return err({
      type: 'api_error',
      status: result.error.status,
      message:
        "UnknownError: (Graph returned an empty error message — this endpoint typically requires a Prefer: odata.maxpagesize=N header on the first call, OR pass --top to cap the page size. The CLI's --top is forwarded as $top, which Graph accepts as an alternative.)",
    });
  }
  return result;
};

const meta: CommandMeta = {
  summary:
    'Get the incremental change set (added / modified / deleted events) for the signed-in user’s default calendar. Use the `@odata.deltaLink` from a previous response to resume. Pass `--top` on the FIRST call (Graph rejects the call with empty `UnknownError:` if neither `--top` nor a Prefer header caps the page size).',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/events/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-delta',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-calendar-events-delta --top 50',
  responseShape: 'collection of changed Microsoft Graph `event` resources under `value[]` plus an `@odata.deltaLink` token',
  pagination: true,
};

export { execute, meta, schema };
