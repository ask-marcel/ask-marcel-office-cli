/*
 * Pure rendering core — every formatter here RETURNS a string and writes
 * nothing.
 *
 * `output.ts` is the stdout shim over this module (the CLI's channel).
 * `composition/mcp.ts` is the second consumer: an MCP stdio server owns
 * stdout as its JSON-RPC frame channel, so it MUST format without writing.
 * That is why the envelope logic lives here rather than inline in `output.ts`
 * as it did through v2.2.
 *
 * The strings returned here are byte-identical to what `output.ts` used to
 * write directly, trailing newline included — `output.test.ts` captures
 * stdout and pins that contract.
 */
import type { ErrorSource } from './error-hints.ts';
import { findErrorHint } from './error-hints.ts';
import { canonicalizeGraphCursor } from './graph-cursor.ts';
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
   * fix.
   */
  readonly sizeHint?: string;
  /**
   * Surfaced when `data.value[]` carries N entries but each entry is empty
   * after stripping `@odata.etag` — the telltale shape of a `--select` with
   * unknown field names (Graph silently drops bogus fields). Distinct from
   * `sizeHint` (which fires on byte count).
   */
  readonly selectHint?: string;
};

// 50 KB — roughly 12 500 tokens at the typical 4-chars/token ratio. The hint
// only fires when the cost of a re-fetch is genuinely worth flagging; below
// this an LLM consumer that doesn't trim is fine.
const SIZE_HINT_THRESHOLD_BYTES = 50_000;

/**
 * Which front end the caller is on. It changes what a remedy can even BE: a
 * terminal caller can redirect stdout to a file, an MCP client cannot (it has
 * no shell), and the `--output-path` flag reaches MCP as an `outputPath` tool
 * param rather than a flag.
 */
type RenderSurface = 'cli' | 'mcp';

/**
 * What the presenter needs in order to name a remedy the caller can actually
 * use. Built in composition (which owns the command manifest); omitted by the
 * renders that have no command behind them at all — the 548 KB `list-commands`
 * manifest, a login summary, a logout status.
 *
 * 2026-07-23: the banner used to call `--output-path` a "universal remedy
 * (works on every command)". It is not. Plain-JSON commands REFUSE the flag,
 * and the two commands that trip the banner most often (`search-all-files`,
 * `microsoft-search-query`) advertise only `--query` — so every remedy the
 * banner named was a dead end there, and an agentic caller following it burned
 * a call to find out.
 */
type SizeHintContext = {
  readonly commandName: string;
  /** `meta.producesBytes` — the manifest flag that decides whether --output-path is accepted. */
  readonly producesBytes: boolean;
  readonly supportsSelect: boolean;
  readonly supportsTop: boolean;
  readonly surface: RenderSurface;
};

const bytesRemedy = (context: SizeHintContext): string =>
  context.surface === 'mcp'
    ? 'Set the `outputPath` param to write the body to a file; the envelope then carries `savedTo` instead of the bytes.'
    : 'Write the body to a file with `--output-path <file>`; the envelope then carries `savedTo` instead of the bytes.';

const jsonRemedy = (context: SizeHintContext): string =>
  context.surface === 'mcp'
    ? `${context.commandName} returns plain JSON, so \`outputPath\` is refused here: there is no body to write.`
    : `${context.commandName} returns plain JSON, so \`--output-path\` is refused; redirect to a file instead: \`ask-marcel-office ${context.commandName} ... > out.json\`.`;

const slimmingRemedy = (context: SizeHintContext): string => {
  const levers = [...(context.supportsSelect ? ['`--select id,subject,...` to slim each item'] : []), ...(context.supportsTop ? ['`--top N` to cut the item count'] : [])];
  if (levers.length === 0) return `${context.commandName} advertises no slimming flags, so narrow the query text itself and re-run.`;
  return `Then re-run with ${levers.join(' and ')}.`;
};

// No command behind the render (manifest, login summary, logout status): claim
// nothing, since every flag-level remedy depends on which command ran.
const NEUTRAL_REMEDY = 'Narrow the request and re-run; slimming flags apply only where the command advertises them (check `--help`).';

const buildSizeHint = (bytes: number, context?: SizeHintContext): string => {
  const head = `Response is ${Math.round(bytes / 1024)} KB (> 50 KB threshold).`;
  if (context === undefined) return `${head} ${NEUTRAL_REMEDY}`;
  return `${head} ${context.producesBytes ? bytesRemedy(context) : jsonRemedy(context)} ${slimmingRemedy(context)}`;
};

const SELECT_HINT =
  "`value[]` contains entries but each is empty (only `@odata.etag`) — likely caused by `--select` field names Graph did not recognise. Graph silently drops unknown `$select` fields; check spelling against the command's `responseShape` in `ask-marcel-office docs <command>`.";

const isPlainRecord = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);

// "Telltale bogus-select" detector. Two shapes both indicate the same trap
// (user passed `--select` with field names Graph silently dropped):
//
// Collection: `{ value: [{@odata.etag}, {@odata.etag}, ...] }`
// — N>0 entries, each empty after stripping `@odata.etag`.
//
// Single resource: `{ "@odata.context": "..." }`
// — top-level object with no keys other than `@odata.*` metadata.
// Example: `get-current-user --select aaaaaaaa,bbbbbbbb`.
//
// Negatives: legitimately empty collection (`value: []`, N=0); single
// resource carrying real fields alongside `@odata.context` (the normal
// happy path); non-object data (string / number / null) — none trigger.
const isMeaningfulKey = (key: string): boolean => !key.startsWith('@odata.');

const looksLikeBogusSelectResponse = (data: unknown): boolean => {
  if (!isPlainRecord(data)) return false;
  const value = data['value'];
  if (value === undefined) {
    // Single-resource shape. Bogus when only `@odata.*` keys remain.
    return Object.keys(data).every((k) => !isMeaningfulKey(k));
  }
  // Collection shape. Bogus when N>0 entries are all empty-after-etag.
  if (!Array.isArray(value) || value.length === 0) return false;
  return value.every((entry) => {
    if (!isPlainRecord(entry)) return false;
    const meaningfulKeys = Object.keys(entry).filter((k) => k !== '@odata.etag');
    return meaningfulKeys.length === 0;
  });
};

// Pagination / cursor tokens that the presenter lifts to the top of the
// envelope so an LLM consumer can write `if (resp.nextLink) ...` instead of
// reaching into `data["@odata.nextLink"]`. `@odata.deltaLink` is included
// so resumption tokens land in the same place — both
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
    ...(typeof nextLink === 'string' ? { nextLink: canonicalizeGraphCursor(nextLink) } : {}),
    ...(typeof deltaLink === 'string' ? { deltaLink: canonicalizeGraphCursor(deltaLink) } : {}),
    ...(typeof count === 'number' ? { count } : {}),
  };
};

const renderJsonToString = (data: unknown, context?: SizeHintContext): string => {
  const envelope = wrap(data);
  const selectHintField = looksLikeBogusSelectResponse(data) ? { selectHint: SELECT_HINT } : {};
  const initial = JSON.stringify({ ...envelope, ...selectHintField });
  // Adding the sizeHint AFTER the size check means a payload that's
  // borderline doesn't get flagged just because the hint itself pushes it
  // over 50 KB. (The selectHint is small enough not to need the same
  // guard.)
  if (initial.length <= SIZE_HINT_THRESHOLD_BYTES) return `${initial}\n`;
  return `${JSON.stringify({ ...envelope, ...selectHintField, sizeHint: buildSizeHint(initial.length, context) })}\n`;
};

const renderTextToString = (data: unknown, context?: SizeHintContext): string => {
  const body = renderTextOutput(data);
  const selectHintLine = looksLikeBogusSelectResponse(data) ? `selectHint: ${SELECT_HINT}\n` : '';
  if (body.length <= SIZE_HINT_THRESHOLD_BYTES) return `${selectHintLine}${body}`;
  // Prepend (not append) so the LLM sees the hint before scrolling through
  // a 50 KB body — matches the `error:` / `hint:` / `source:` placement
  // convention for the error path. selectHint comes BEFORE sizeHint so the
  // more-actionable signal (you have a typo) is at the very top.
  return `${selectHintLine}sizeHint: ${buildSizeHint(body.length, context)}\n${body}`;
};

/**
 * Render a use-case success value to its final string, newline included.
 * `output.ts` writes this to stdout; `mcp.ts` returns it as tool content.
 *
 * `context` is what makes the oversized-response banner honest — pass it
 * whenever a registry command produced the data. Omit it for the renders that
 * have no command behind them; the banner then claims no flag-level remedy.
 */
const renderToString = (data: unknown, format: OutputFormat, context?: SizeHintContext): string =>
  format === 'json' ? renderJsonToString(data, context) : renderTextToString(data, context);

/**
 * Render an error to its final string, newline included. The `hint` / `source`
 * lookup is shared with the CLI so an MCP consumer sees the same curated
 * remedies (`error-hints.ts`) the terminal does.
 */
const renderErrorToString = (message: string, format: OutputFormat, errorCode?: string, explicitSource?: ErrorSource, retryAfterSeconds?: number): string => {
  // `errorCode` is an additive field — old consumers
  // keying on `error: string` continue to work; new consumers can branch on
  // the structured code (`itemNotFound`, `InvalidIdMalformed`, `MissingScope`,
  // CLI-rewrite codes like `cli_rewrite_orderby_title`, etc.) without
  // substring-matching the human message.
  //
  // ALSO additive — `hint` and `source` come from
  // the central table in `error-hints.ts`. Surfaced in both formats so an
  // LLM in text mode no longer has to guess what `ErrorInvalidIdMalformed`
  // means.
  //
  // asymmetry between hint-matched and
  // bare-Graph errors was a real problem — an LLM that branched on
  // `response.source` saw `{ok, error, errorCode}` for half its Graph
  // failures and `{ok, error, errorCode, hint, source}` for the other half.
  // The fix: `explicitSource` lets the call site (cli.ts action handler)
  // stamp `source: 'graph'` from the `result.error.type` discriminator
  // regardless of whether a hint rule matched. The hint's source (when
  // matched) still wins because it's curated; the explicit source is the
  // typed fallback. End-result envelope: `{ok, error, errorCode?, hint?,
  // source?}` where only `hint` is conditional — `source` is always present
  // when the caller knew the failure category.
  const hint = findErrorHint(message, errorCode);
  const source = hint?.source ?? explicitSource;
  if (format === 'json') {
    const payload = {
      ok: false,
      error: message,
      ...(errorCode ? { errorCode } : {}),
      ...(hint ? { hint: hint.hint } : {}),
      ...(source ? { source } : {}),
      ...(retryAfterSeconds === undefined ? {} : { retryAfterSeconds }),
    };
    return `${JSON.stringify(payload)}\n`;
  }
  const lines = [`error: ${message}`];
  if (hint) lines.push(`hint: ${hint.hint}`);
  if (source) lines.push(`source: ${source}`);
  if (retryAfterSeconds !== undefined) lines.push(`retryAfter: ${retryAfterSeconds}s`);
  return `${lines.join('\n')}\n`;
};

export { renderErrorToString, renderToString };
export type { OutputFormat, RenderSurface, SizeHintContext };
