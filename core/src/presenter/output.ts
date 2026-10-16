/*
 * The CLI's stdout channel — a thin shim over the pure formatters in
 * `render-to-string.ts`.
 *
 * This module is one of the three sanctioned `process.stdout` writers, beside
 * the CLI builder and the token helper (a lint allowlist bans it everywhere
 * else). Everything renderable lives in
 * `render-to-string.ts` so a consumer that does NOT own stdout can reuse it:
 * `composition/mcp-shell.ts` serves the MCP stdio protocol, where stdout carries
 * JSON-RPC frames and a stray write corrupts the stream.
 */
import type { NameSource } from '../use-cases/commands/bin-names.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import type { ErrorSource } from './error-hints.ts';
import type { OutputFormat, RenderContext } from './render-to-string.ts';
import { renderErrorToString, renderToString } from './render-to-string.ts';

// `names`: the package whose bins fill the CLI's own text (see render-to-string.ts).
const render = (data: unknown, logger: Logger, format: OutputFormat, context?: RenderContext, names?: NameSource): void => {
  logger.info('output_rendered', {});
  process.stdout.write(renderToString(data, format, context, names));
};

const renderError = (message: string, format: OutputFormat, errorCode?: string, explicitSource?: ErrorSource, retryAfterSeconds?: number, names?: NameSource): void => {
  process.stdout.write(renderErrorToString(message, format, errorCode, explicitSource, retryAfterSeconds, names));
};

export { render, renderError };
export type { OutputFormat };
