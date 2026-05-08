import type { Logger } from '../use-cases/ports/logger.ts';

type SuccessEnvelope = {
  readonly ok: true;
  readonly data: unknown;
  readonly nextLink?: string;
  readonly deltaLink?: string;
  readonly count?: number;
};

const isPlainRecord = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);

// Pagination / cursor tokens that the presenter lifts to the top of the
// envelope so an LLM consumer can write `if (resp.nextLink) ...` instead of
// reaching into `data["@odata.nextLink"]`. `@odata.deltaLink` is included
// (audit v1.0.0 §4) so resumption tokens land in the same place — both
// nextLink and deltaLink are pagination cursors and should sit at the same
// level. `@odata.count` is also lifted as a sibling.
const HOIST_KEYS: ReadonlySet<string> = new Set(['@odata.nextLink', '@odata.deltaLink', '@odata.count']);

const wrap = (data: unknown): SuccessEnvelope => {
  if (!isPlainRecord(data)) return { ok: true, data };
  const nextLink = data['@odata.nextLink'];
  const deltaLink = data['@odata.deltaLink'];
  const count = data['@odata.count'];
  const cleaned: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(data)) {
    if (HOIST_KEYS.has(key)) continue;
    cleaned[key] = value;
  }
  return {
    ok: true,
    data: cleaned,
    ...(typeof nextLink === 'string' ? { nextLink } : {}),
    ...(typeof deltaLink === 'string' ? { deltaLink } : {}),
    ...(typeof count === 'number' ? { count } : {}),
  };
};

const render = (data: unknown, logger: Logger): void => {
  logger.info('output_rendered', {});
  process.stdout.write(`${JSON.stringify(wrap(data))}\n`);
};

const renderError = (message: string): void => {
  process.stdout.write(`${JSON.stringify({ ok: false, error: message })}\n`);
};

export { render, renderError };
