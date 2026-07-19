import { ok, type Result } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';

/**
 * Graph answers createReply / createForward / PATCH with the ENTIRE message,
 * quoted history included — a `--comment` edit on a long thread echoed a 174 KB
 * body straight back at the caller who had just written it. For an LLM consumer
 * that is pure context cost: what it needs is confirmation the write landed and
 * the id to act on next, not its own text read back.
 *
 * No `--select` escape hatch by design: the full body is one `get-mail-message`
 * away, and a flag that re-exposes already-reachable data is a flag that has to
 * be documented, tested, and explained forever.
 *
 * `bodyPreview` stays — it is Graph's own ~255-char summary, enough to tell
 * WHICH draft answered without dragging the thread along.
 */
const DRAFT_RESPONSE_FIELDS = ['id', 'subject', 'toRecipients', 'ccRecipients', 'bccRecipients', 'importance', 'bodyPreview', 'isDraft', 'webLink', 'conversationId'] as const;

const slimDraftResponse = (value: unknown): unknown => {
  // A 204 surfaces as `{ ok: true }` and carries no id. Project only what is
  // recognisably a message; anything else passes through, because blanking it
  // would destroy whatever Graph was actually saying.
  if (typeof value !== 'object' || value === null || !Object.hasOwn(value, 'id')) return value;
  const source = value as Record<string, unknown>;
  const slim: Record<string, unknown> = {};
  for (const field of DRAFT_RESPONSE_FIELDS) {
    if (Object.hasOwn(source, field)) slim[field] = source[field];
  }
  return slim;
};

const slimDraftResult = (result: Result<unknown, GraphError>): Result<unknown, GraphError> => (result.ok ? ok(slimDraftResponse(result.value)) : result);

export { slimDraftResult };
