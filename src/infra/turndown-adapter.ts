import TurndownService from 'turndown';
import { gfm } from 'turndown-plugin-gfm';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
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
 * GFM plugin enabled so HTML tables render as pipe-delimited markdown
 * tables instead of flat paragraphs (needed for docx → markdown to
 * preserve table structure that mammoth emits as <table><tr><td>).
 *
 * Turndown's default HTML parser drops comment nodes during parsing,
 * so we don't need a custom comment-stripping rule. <script> and
 * <style> aren't dropped by default — those need explicit removal so
 * their text content doesn't leak into the markdown output.
 *
 * Lives in `src/infra/` because turndown can throw on malformed DOM
 * (e.g. Outlook MSO HTML that produces a `<td>` whose parent is undefined
 * during traversal — surfaced as `Cannot read properties of undefined
 * (reading 'parentNode')`). The adapter catches that and returns a
 * `markdown_conversion_failed` GraphError so the JSON envelope contract
 * holds even when the underlying library blows up.
 */
const buildService = (): TurndownService => {
  const td = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced',
    bulletListMarker: '-',
    emDelimiter: '_',
    strongDelimiter: '**',
  });
  td.use(gfm);
  td.remove(['script', 'style']);
  return td;
};

const htmlToMarkdown = (html: string): Result<string, GraphError> => {
  try {
    return ok(buildService().turndown(html));
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : String(e);
    return err({ type: 'api_error', status: 500, message: `markdown conversion failed: ${message}` });
  }
};

export { htmlToMarkdown };
