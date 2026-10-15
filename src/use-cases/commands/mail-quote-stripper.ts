import { z } from 'zod';
import type { CommandOptionMeta } from './command-types.ts';
import { findPlainTextQuoteBoundary, findQuoteBoundary } from './quote-boundary.ts';

/**
 * Strips quoted reply chains / forwarded-message blocks from an Outlook or
 * Gmail HTML email body so long threads don't duplicate quoted content into the
 * model's context. Conservative: truncates the body at the EARLIEST well-known
 * reply/forward boundary marker (quote-boundary.ts lists them) and replaces the
 * tail with a single visible placeholder — nothing is removed silently, and
 * `--keep-quoted true` on `convert-mail-to-markdown` restores the full body.
 * Pure string transform.
 */

const STRIP_MARKER = '<p><em>[Quoted reply chain removed — pass --keep-quoted true to include it]</em></p>';

const PLAINTEXT_MARKER = '[Quoted reply chain removed — pass --keep-quoted true to include it]';

const stripQuotedReplies = (html: string): { readonly html: string; readonly stripped: boolean } => {
  const cut = findQuoteBoundary(html);
  if (cut === -1) return { html, stripped: false };
  return { html: `${html.slice(0, cut)}${STRIP_MARKER}`, stripped: true };
};

const stripQuotedPlainText = (text: string): { readonly text: string; readonly stripped: boolean } => {
  const cut = findPlainTextQuoteBoundary(text);
  if (cut === -1) return { text, stripped: false };
  return { text: `${text.slice(0, cut)}${PLAINTEXT_MARKER}`, stripped: true };
};

// Shared `--keep-quoted` schema field + option meta for every command that can
// hand an Outlook `.msg` to the markdown dispatch (the six converters below and
// `convert-mail-to-markdown`'s own Graph-body path). Declared once so the six
// descriptions cannot drift apart, and so the strip marker's remedy is real on
// every command that can emit it.
const keepQuotedSchemaField = z.enum(['true', 'false']).optional();

const keepQuotedOption: CommandOptionMeta = {
  name: 'keep-quoted',
  key: 'keepQuoted',
  required: false,
  description:
    'Applies to email input only: an Outlook `.msg` or a raw `.eml`. The quoted reply chain / forwarded-message block is stripped by default (it duplicates history and inflates the context budget) and replaced with a single visible marker naming this flag, so nothing is removed silently. Pass `--keep-quoted true` to render the full body. The recognized markers are the same set `convert-mail-to-markdown` uses — see `{bin} docs convert-mail-to-markdown`.',
  argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
};

export { keepQuotedOption, keepQuotedSchemaField, stripQuotedPlainText, stripQuotedReplies };
// Transitional: the finders live in quote-boundary.ts (core). Existing tests still import them
// from here; the re-export goes when those imports move (package-split step 16, rule-24 batch).
export { findPlainTextQuoteBoundary, findQuoteBoundary } from './quote-boundary.ts';
