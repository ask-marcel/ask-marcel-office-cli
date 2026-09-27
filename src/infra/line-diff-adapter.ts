import { FILE_HEADERS_ONLY, formatPatch, structuredPatch } from 'diff';

/**
 * A unified line diff of two texts, through jsdiff (the `diff` package): the
 * `--- a` / `+++ b` headers, hunks with three lines of context, and the count
 * of added and removed lines.
 *
 * Past `MAX_EDITS` added plus removed lines (a changed line counts once each
 * way) a diff would be about as long as both texts: jsdiff stops early
 * (`maxEditLength`) and the answer is `undefined`, so a pair of unrelated
 * 10,000-line decks costs neither the time nor the memory of a full diff.
 */

const MAX_EDITS = 1_000;
const CONTEXT_LINES = 3;

type LineDiff = { readonly patch: string; readonly added: number; readonly removed: number };

// A render rarely ends with a newline; without one jsdiff marks the last line
// "No newline at end of file" on both sides, which is noise here. An empty render
// stays empty: a lone newline would diff as one removed blank line.
const withFinalNewline = (text: string): string => (text === '' || text.endsWith('\n') ? text : `${text}\n`);

const lineDiff = (fromName: string, fromText: string, toName: string, toText: string): LineDiff | undefined => {
  const patch = structuredPatch(fromName, toName, withFinalNewline(fromText), withFinalNewline(toText), undefined, undefined, {
    context: CONTEXT_LINES,
    maxEditLength: MAX_EDITS,
  });
  if (patch === undefined) return undefined;
  const lines = patch.hunks.flatMap((hunk) => hunk.lines);
  return {
    patch: patch.hunks.length === 0 ? '' : formatPatch(patch, FILE_HEADERS_ONLY),
    added: lines.filter((line) => line.startsWith('+')).length,
    removed: lines.filter((line) => line.startsWith('-')).length,
  };
};

export { lineDiff, MAX_EDITS };
