import { findQuoteBoundary } from './mail-quote-stripper.ts';

/**
 * Lifts the sender's own signature out of a message they sent, so a draft Graph
 * minted (which carries no signature at all) can be given one.
 *
 * OWA and new Outlook wrap the signature in `<div id="Signature">`. Outlook
 * desktop does not, which is why the caller must be able to say "no signature
 * here" rather than fall back to a guess: any heuristic (last N lines, a `--`
 * rule) would sometimes return the last paragraph of a real message.
 *
 * The search stops at the quote boundary. A reply quotes the other person's
 * signature too, and theirs sits BELOW the boundary; without the limit, a
 * message whose author had no signature would return their colleague's.
 *
 * String-index based, like the quote stripper it borrows the boundary from: the
 * repo has no DOM parser and this does not justify one.
 */

const SIGNATURE_DIV = /<div[^>]*\bid="Signature"/i;
const DIV_TAG = /<div\b|<\/div>/gi;

/** Index just past the `</div>` matching the `<div` at `start`, or -1 if unbalanced. */
const findMatchingDivClose = (html: string, start: number): number => {
  DIV_TAG.lastIndex = start;
  let depth = 0;
  for (let tag = DIV_TAG.exec(html); tag !== null; tag = DIV_TAG.exec(html)) {
    depth += tag[0].startsWith('</') ? -1 : 1;
    if (depth === 0) return tag.index + tag[0].length;
  }
  return -1;
};

const extractSignatureBlock = (html: string): string | undefined => {
  const boundary = findQuoteBoundary(html);
  const limit = boundary === -1 ? html.length : boundary;
  const marker = SIGNATURE_DIV.exec(html);
  if (marker === null || marker.index >= limit) return undefined;
  const end = findMatchingDivClose(html, marker.index);
  if (end === -1) return undefined;
  return html.slice(marker.index, end);
};

export { extractSignatureBlock };
