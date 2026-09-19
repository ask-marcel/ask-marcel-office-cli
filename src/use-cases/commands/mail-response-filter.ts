import type { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandOptionMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

/**
 * Graph will neither `$filter` on `isof(...)` nor `$select`
 * `meetingMessageType`, but every listed message carries `@odata.type`, and
 * an accepted, declined or tentative reply to an invite is
 * `#microsoft.graph.eventMessageResponse`. Dropping those client-side is what
 * `--exclude-meeting-responses true` does; the number dropped comes back as
 * `excludedMeetingResponses` so a shorter page is not mistaken for the end.
 */

const MEETING_RESPONSE_TYPE = '#microsoft.graph.eventMessageResponse';

const EXCLUDE_MEETING_RESPONSES_OPTION: CommandOptionMeta = {
  name: 'exclude-meeting-responses',
  key: 'excludeMeetingResponses',
  required: false,
  description:
    'Pass `--exclude-meeting-responses true` to drop the Accepted / Declined / Tentative replies to meeting invites (`@odata.type` `eventMessageResponse`) from the page, client-side; the count dropped comes back as `excludedMeetingResponses`. Invites themselves stay.',
  argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
};

const isMeetingResponse = (item: unknown): boolean => item !== null && typeof item === 'object' && (item as Record<string, unknown>)['@odata.type'] === MEETING_RESPONSE_TYPE;

/** Drops the meeting responses from a `value[]` page and records how many went. */
const excludeMeetingResponses = (result: Result<unknown, GraphError>): Result<unknown, GraphError> => {
  if (!result.ok) return result;
  const body = result.value as { readonly value?: unknown } & Record<string, unknown>;
  if (!Array.isArray(body.value)) return result;
  const kept = body.value.filter((item) => !isMeetingResponse(item));
  const excluded = body.value.length - kept.length;
  return ok({ ...body, value: kept, ...(excluded === 0 ? {} : { excludedMeetingResponses: excluded }) });
};

/** Wraps a listing's execute: the merged schema validates the flag, then the flag is kept away from the OData schema behind it. */
const withMeetingResponseFilter = (schema: z.ZodType, inner: Command['execute']): Command['execute'] => {
  const filtered: Command['execute'] = async (graph, params) => {
    const parsed = schema.safeParse(params);
    if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
    const { excludeMeetingResponses: flag, ...rest } = params;
    const result = await inner(graph, rest);
    return flag === 'true' ? excludeMeetingResponses(result) : result;
  };
  return filtered;
};

export { EXCLUDE_MEETING_RESPONSES_OPTION, excludeMeetingResponses, withMeetingResponseFilter };
