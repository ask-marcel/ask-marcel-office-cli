/*
 * The CLI's stdout channel — a thin shim over the pure formatters in
 * `render-to-string.ts`.
 *
 * This module is the ONLY sanctioned `process.stdout` writer in the codebase
 * (enforced at the lint level). Everything renderable lives in
 * `render-to-string.ts` so a consumer that does NOT own stdout can reuse it:
 * `composition/mcp.ts` serves the MCP stdio protocol, where stdout carries
 * JSON-RPC frames and a stray write corrupts the stream.
 */
import type { Logger } from '../use-cases/ports/logger.ts';
import type { ErrorSource } from './error-hints.ts';
import type { OutputFormat, SizeHintContext } from './render-to-string.ts';
import { renderErrorToString, renderToString } from './render-to-string.ts';

const render = (data: unknown, logger: Logger, format: OutputFormat, context?: SizeHintContext): void => {
  logger.info('output_rendered', {});
  process.stdout.write(renderToString(data, format, context));
};

const renderError = (message: string, format: OutputFormat, errorCode?: string, explicitSource?: ErrorSource, retryAfterSeconds?: number): void => {
  process.stdout.write(renderErrorToString(message, format, errorCode, explicitSource, retryAfterSeconds));
};

export { render, renderError };
export type { OutputFormat };
