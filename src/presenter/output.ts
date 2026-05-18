import type { Logger } from '../use-cases/ports/logger.ts';
import { renderTextOutput } from './output-text.ts';

type OutputFormat = 'text' | 'json';

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

const renderJson = (data: unknown): void => {
  process.stdout.write(`${JSON.stringify(wrap(data))}\n`);
};

const render = (data: unknown, logger: Logger, format: OutputFormat): void => {
  logger.info('output_rendered', {});
  if (format === 'json') renderJson(data);
  else process.stdout.write(renderTextOutput(data));
};

const renderError = (message: string, format: OutputFormat): void => {
  if (format === 'json') process.stdout.write(`${JSON.stringify({ ok: false, error: message })}\n`);
  else process.stdout.write(`error: ${message}\n`);
};

export { render, renderError };
export type { OutputFormat };
