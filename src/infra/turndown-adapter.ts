import TurndownService from 'turndown';
import { gfm } from 'turndown-plugin-gfm';
import type { Result } from '../domain/result.ts';
import { ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

/**
 * Turn arbitrary HTML — mammoth-converted docx, Graph-converted Office
 * docs, OneNote pages, Outlook email bodies, SharePoint pages — into
 * clean markdown.
 *
 * Wrapped in a function-only adapter so the rest of the codebase stays
 * class-free per atelier rule 1. Each call constructs a fresh service
 * (turndown is cheap to instantiate; no mutable shared state).
 *
 * GFM plugin enabled by default so HTML tables render as pipe-delimited
 * markdown tables instead of flat paragraphs (needed for docx → markdown
 * to preserve table structure that mammoth emits as <table><tr><td>).
 *
 * Turndown's default HTML parser drops comment nodes during parsing,
 * so we don't need a custom comment-stripping rule. <script> and
 * <style> aren't dropped by default — those need explicit removal so
 * their text content doesn't leak into the markdown output.
 *
 * Lives in `src/infra/` because turndown can throw on malformed DOM. The
 * dominant culprit is the GFM table plugin's `cellInternals` walker,
 * which assumes every `<td>` has an ancestor `<tr>` reachable via
 * `parentNode` — Outlook MSO HTML routinely violates this with floating
 * `<td>` cells, nested misnested tables, or `<table>` wrappers with no
 * `<tbody>`. The walker then dereferences `undefined.parentNode` and
 * throws `Cannot read properties of undefined (reading 'parentNode')`.
 *
 * Three-tier graceful degradation when that happens:
 *
 *   1. turndown WITH GFM (clean markdown tables)
 *   2. turndown WITHOUT GFM (core markdown; tables become flat paragraphs
 *      but the rest of the document converts cleanly)
 *   3. stripped-text fallback (state-machine HTML walker; preserves text,
 *      basic entities, and newlines on block boundaries)
 *
 * Tier 2 covers the vast majority of Outlook MSO bodies; tier 3 only
 * fires when even the core walker chokes. Each downgrade prepends a
 * markdown blockquote note naming the underlying error so observability
 * survives the degradation.
 */
const buildService = (options: { readonly gfm: boolean }): TurndownService => {
  const td = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced',
    bulletListMarker: '-',
    emDelimiter: '_',
    strongDelimiter: '**',
  });
  if (options.gfm) td.use(gfm);
  td.remove(['script', 'style']);
  return td;
};

const errorMessageOf = (e: unknown): string => (e instanceof Error ? e.message : String(e));

const decodeBasicEntities = (s: string): string =>
  s
    .replaceAll(/&nbsp;/gi, ' ')
    .replaceAll(/&amp;/gi, '&')
    .replaceAll(/&lt;/gi, '<')
    .replaceAll(/&gt;/gi, '>')
    .replaceAll(/&quot;/gi, '"')
    .replaceAll(/&#39;/gi, "'")
    .replaceAll(/&apos;/gi, "'");

const BLOCK_END_TAGS = new Set(['/p', '/div', '/h1', '/h2', '/h3', '/h4', '/h5', '/h6', '/li', '/tr']);

const tagInsertsNewline = (rawTag: string): boolean => {
  const lower = rawTag.toLowerCase();
  if (lower.startsWith('br')) {
    if (lower === 'br' || lower === 'br/' || lower.startsWith('br ') || lower.startsWith('br/')) return true;
  }
  return BLOCK_END_TAGS.has(lower);
};

/**
 * Walk the HTML byte-by-byte and emit either the character (when outside a
 * tag) or a newline (when closing a block tag / hitting <br>). Avoids the
 * regex `<[^>]*>` pattern that sonarjs flags as super-linear-backtracking
 * vulnerable, and skips <script>/<style> bodies in the same pass.
 */
const stripHtmlToText = (html: string): string => {
  let out = '';
  let cursor = 0;
  while (cursor < html.length) {
    const lt = html.indexOf('<', cursor);
    if (lt === -1) {
      out += html.slice(cursor);
      break;
    }
    out += html.slice(cursor, lt);
    const gt = html.indexOf('>', lt + 1);
    if (gt === -1) break;
    const rawTag = html.slice(lt + 1, gt);
    const lowerTag = rawTag.toLowerCase();
    if (rawTag.startsWith('!--')) {
      const end = html.indexOf('-->', lt + 4);
      cursor = end === -1 ? html.length : end + 3;
      continue;
    }
    if (lowerTag.startsWith('script') || lowerTag.startsWith('style')) {
      const tagName = lowerTag.startsWith('script') ? 'script' : 'style';
      const closer = html.toLowerCase().indexOf(`</${tagName}>`, gt + 1);
      cursor = closer === -1 ? html.length : closer + tagName.length + 3;
      continue;
    }
    if (tagInsertsNewline(rawTag)) out += '\n';
    cursor = gt + 1;
  }
  return decodeBasicEntities(out)
    .replaceAll(/[ \t]+/g, ' ')
    .replaceAll(/\n{3,}/g, '\n\n')
    .trim();
};

const htmlToMarkdown = (html: string): Result<string, GraphError> => {
  try {
    return ok(buildService({ gfm: true }).turndown(html));
  } catch (gfmError: unknown) {
    try {
      const md = buildService({ gfm: false }).turndown(html);
      const note = `> _GFM table conversion failed: ${errorMessageOf(gfmError)}; tables flattened to paragraphs_`;
      return ok(`${note}\n\n${md}`);
    } catch (coreError: unknown) {
      const fallback = stripHtmlToText(html);
      const note = `> _markdown conversion failed: ${errorMessageOf(coreError)}; showing stripped HTML body_`;
      return ok(fallback.length > 0 ? `${note}\n\n${fallback}` : note);
    }
  }
};

export { htmlToMarkdown };
