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
 *   - `insertCommentAboveQuote` INSERTS at the top of the body. Used when Graph
 *     has just minted the draft: everything below is Graph's own scaffolding
 *     (empty comment div, `<hr>` separator, quoted history), kept verbatim, and
 *     the author's reply leads the body — above Graph's separator, not under it.
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

/**
 * True when the author's own markup carries a quote boundary marker, e.g. they
 * pasted a reply chain into it. Such a comment must be refused: the splice would
 * keep the marker verbatim, and the NEXT revision would cut the draft AT it,
 * silently dropping the real quoted history below.
 */
const commentCarriesQuoteBoundary = (commentHtml: string): boolean => findQuoteBoundary(commentHtml) !== -1;

/** The refusal copy for `commentCarriesQuoteBoundary`, named for the flag that carried it. */
const boundaryMarkerRefusal = (flagName: string): string =>
  `${flagName} carries a quoted-reply boundary marker (a pasted gmail_quote container, an Outlook divRplyFwdMsg / appendonsend / border-top separator, or a bold From: + Sent: header pair). It would be kept verbatim, and a later \`update-mail-draft --comment\` edit would cut the draft at that marker and lose the real quoted history below it. Remove the pasted quote from your text, or pass --body-content-type Text to have it escaped into literal characters.`;

const insertCommentAboveQuote = (html: string, commentHtml: string): SpliceResult => {
  // Insert at the top of the body, ABOVE Graph's reply scaffolding (its empty
  // comment div, its `<hr>` separator, and the quoted history). Graph does not
  // always emit `appendonsend`: on tenants where it doesn't, the earliest quote
  // boundary is `divRplyFwdMsg` sitting BELOW the plain `<hr>`, so inserting at
  // the boundary parked the comment under the separator (reported 2026-07-19).
  // Body-start is above every scaffolding variant, so the author's text always
  // leads the body. The revise path (`replaceCommentAboveQuote`) already writes
  // from body-start, so create and revise agree on where the comment lives.
  const at = findBodyInsertStart(html);
  return { html: `${html.slice(0, at)}${commentHtml}${html.slice(at)}`, boundaryFound: findQuoteBoundary(html) !== -1 };
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

export {
  boundaryMarkerRefusal,
  commentCarriesQuoteBoundary,
  escapeTextAsHtml,
  findBodyInsertStart,
  insertCommentAboveQuote,
  replaceCommentAboveQuote,
  replacePlainTextCommentAboveQuote,
};
