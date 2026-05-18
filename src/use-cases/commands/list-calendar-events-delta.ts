import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { topOnlyOptions, topOnlyShape } from './odata-query.ts';

const schema = z.object({}).extend(topOnlyShape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  // Graph's calendar delta endpoint REJECTS `$top` as a query parameter
  // (`ErrorInvalidUrlQuery`) but ACCEPTS the equivalent `Prefer: odata.maxpagesize=N`
  // header. The CLI keeps `--top` as the user-facing flag and translates it
  // to the header at the IO boundary so the cross-command contract stays
  // uniform.
  const headers: Record<string, string> = {};
  if (parsed.data.top !== undefined) headers['Prefer'] = `odata.maxpagesize=${parsed.data.top}`;
  const result = await graph.get('/me/events/delta()', headers);
  if (result.ok) return result;
  // Graph's calendar delta endpoint returns `UnknownError: ` (empty message)
  // when called without a Prefer: odata.maxpagesize header AND without a
  // resumption deltaLink. Augment the empty-message case with actionable
  // guidance.
  if (result.error.type === 'api_error' && result.error.message.trim() === 'UnknownError:') {
    return err({
      type: 'api_error',
      status: result.error.status,
      message:
        'UnknownError: (Graph returned an empty error message — this endpoint requires either `--top <N>` to cap the page size, or a resumption `@odata.deltaLink` from a previous call. The CLI translates `--top` into the `Prefer: odata.maxpagesize=N` header that Graph actually accepts; `$top` as a URL query parameter is rejected.)',
    });
  }
  return result;
};

const meta: CommandMeta = {
  summary:
    "Get the incremental change set (added / modified / deleted events) for the signed-in user's default calendar. Use the `@odata.deltaLink` from a previous response to resume. The CLI translates `--top` into the `Prefer: odata.maxpagesize=N` header internally; `$top` as a URL query is rejected by Graph (`ErrorInvalidUrlQuery`). Other OData passthroughs (`$select`, `$filter`, `$orderby`, `$skip`) are silently ignored by Graph on this delta endpoint, so the CLI does NOT expose them — slice / sort / project client-side. Most tenants accept the call without `--top` and return a sane page (~200 events); pass `--top` only when you want a smaller bound. If Graph returns an empty `UnknownError:` (rare), the CLI rewrites it to a hint pointing at the `--top` workaround.",
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/events/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-delta',
  options: [...topOnlyOptions],
  example: 'ask-marcel list-calendar-events-delta --top 50',
  responseShape:
    'collection of changed Microsoft Graph `event` resources under `data.value[]`. Cursor tokens are hoisted to envelope level: top-level `nextLink` while paging, then top-level `deltaLink` on the final page.',
  pagination: true,
};

export { execute, meta, schema };
