import { findPlainTextQuoteBoundary, findQuoteBoundary } from './mail-quote-stripper.ts';

/**
 * Places an author's reply text into a draft body Graph already minted, without
 * disturbing the quoted history below it. Pure string transforms.
 *
 * Graph's reply/forward drafts are full HTML documents (`<html><head><style>…`)
 * whose head styles are load-bearing for the quoted tail, so nothing here ever
 * rebuilds the document: the comment is spliced in and every other byte Graph
 * wrote survives verbatim. That is the whole point. PATCHing `body` with freshly
 * built HTML is what destroyed the quoted thread once before (fixed 2026-07-13);
 * the rule is not "never PATCH body", it is "never PATCH body without the quote
 * still in it".
 *
 * Two operations, deliberately different:
 *   - `insertCommentAboveQuote` INSERTS at the quote boundary. Used when Graph
 *     has just minted the draft, so everything above the quote is Graph's own
 *     empty-comment scaffolding and must be kept.
 *   - `replaceCommentAboveQuote` REPLACES everything from the body tag to the
 *     quote boundary. Used when revising an existing draft: inserting there
 *     would stack a second reply above the first on every edit.
 */

const BODY_OPEN_TAG = /<body[^>]*>/i;

type SpliceResult = { readonly html: string; readonly boundaryFound: boolean };
type PlainTextSpliceResult = { readonly text: string; readonly boundaryFound: boolean };

/**
 * Renders an author's plain text as HTML: markup they typed shows as characters
 * rather than taking effect, and their newlines survive as breaks. No wrapper
 * element, so the caller decides the block context.
 */
const escapeTextAsHtml = (text: string): string =>
  text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
    .replaceAll(/\r\n|\n/g, '<br>');

/** Index just past the `<body>` open tag, or 0 for a fragment that has none. */
const findBodyInsertStart = (html: string): number => {
  const match = BODY_OPEN_TAG.exec(html);
  return match === null ? 0 : match.index + match[0].length;
};

const insertCommentAboveQuote = (html: string, commentHtml: string): SpliceResult => {
  const cut = findQuoteBoundary(html);
  const at = cut === -1 ? findBodyInsertStart(html) : cut;
  return { html: `${html.slice(0, at)}${commentHtml}${html.slice(at)}`, boundaryFound: cut !== -1 };
};

const replaceCommentAboveQuote = (html: string, commentHtml: string): SpliceResult => {
  const cut = findQuoteBoundary(html);
  if (cut === -1) return { html, boundaryFound: false };
  return { html: `${html.slice(0, findBodyInsertStart(html))}${commentHtml}${html.slice(cut)}`, boundaryFound: true };
};

const replacePlainTextCommentAboveQuote = (text: string, comment: string): PlainTextSpliceResult => {
  const cut = findPlainTextQuoteBoundary(text);
  if (cut === -1) return { text, boundaryFound: false };
  return { text: `${comment}\n\n${text.slice(cut)}`, boundaryFound: true };
};

export { escapeTextAsHtml, findBodyInsertStart, insertCommentAboveQuote, replaceCommentAboveQuote, replacePlainTextCommentAboveQuote };
