import type { Logger } from '../use-cases/ports/logger.ts';
import { findErrorHint } from './error-hints.ts';
import { renderTextOutput } from './output-text.ts';

type OutputFormat = 'text' | 'json';

type SuccessEnvelope = {
  readonly ok: true;
  readonly data: unknown;
  readonly nextLink?: string;
  readonly deltaLink?: string;
  readonly count?: number;
  /**
   * Surfaced when the rendered payload exceeds `SIZE_HINT_THRESHOLD_BYTES`.
   * Tells the LLM consumer how to shrink the next call — universal advice,
   * since the presenter doesn't know which command produced the data
   * (slim-default `--select`, lower `--top`, or `--output-path` for binary).
   * Audit Alex-session §3 fix.
   */
  readonly sizeHint?: string;
};

// 50 KB — roughly 12 500 tokens at the typical 4-chars/token ratio. The hint
// only fires when the cost of a re-fetch is genuinely worth flagging; below
// this an LLM consumer that doesn't trim is fine.
const SIZE_HINT_THRESHOLD_BYTES = 50_000;

const buildSizeHint = (bytes: number): string =>
  `Response is ${Math.round(bytes / 1024)} KB (> 50 KB threshold). Trim with \`--select id,...\` if the command supports OData projection, lower \`--top\`, or use the global \`--output-path <file>\` flag to write bytes to disk instead of inlining them.`;

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
  const envelope = wrap(data);
  const initial = JSON.stringify(envelope);
  // Adding the hint AFTER the size check means a payload that's borderline
  // doesn't get flagged just because the hint itself pushes it over 50 KB.
  if (initial.length <= SIZE_HINT_THRESHOLD_BYTES) {
    process.stdout.write(`${initial}\n`);
    return;
  }
  const withHint = JSON.stringify({ ...envelope, sizeHint: buildSizeHint(initial.length) });
  process.stdout.write(`${withHint}\n`);
};

const renderText = (data: unknown): void => {
  const body = renderTextOutput(data);
  if (body.length <= SIZE_HINT_THRESHOLD_BYTES) {
    process.stdout.write(body);
    return;
  }
  // Prepend (not append) so the LLM sees the hint before scrolling through
  // a 50 KB body — matches the `error:` / `hint:` / `source:` placement
  // convention for the error path.
  process.stdout.write(`sizeHint: ${buildSizeHint(body.length)}\n${body}`);
};

const render = (data: unknown, logger: Logger, format: OutputFormat): void => {
  logger.info('output_rendered', {});
  if (format === 'json') renderJson(data);
  else renderText(data);
};

const renderError = (message: string, format: OutputFormat, errorCode?: string): void => {
  // Audit round-7 Wave G: `errorCode` is an additive field — old consumers
  // keying on `error: string` continue to work; new consumers can branch on
  // the structured code (`itemNotFound`, `InvalidIdMalformed`, `MissingScope`,
  // CLI-rewrite codes like `cli_rewrite_orderby_title`, etc.) without
  // substring-matching the human message.
  //
  // Audit Alex-session §2: ALSO additive — `hint` and `source` come from
  // the central table in `error-hints.ts`. Surfaced in both formats so an
  // LLM in text mode no longer has to guess what `ErrorInvalidIdMalformed`
  // means.
  const hint = findErrorHint(message, errorCode);
  if (format === 'json') {
    const payload = {
      ok: false,
      error: message,
      ...(errorCode ? { errorCode } : {}),
      ...(hint ? { hint: hint.hint, source: hint.source } : {}),
    };
    process.stdout.write(`${JSON.stringify(payload)}\n`);
    return;
  }
  const lines = [`error: ${message}`];
  if (hint) {
    lines.push(`hint: ${hint.hint}`);
    lines.push(`source: ${hint.source}`);
  }
  process.stdout.write(`${lines.join('\n')}\n`);
};

export { render, renderError };
export type { OutputFormat };
