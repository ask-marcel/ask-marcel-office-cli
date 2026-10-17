/**
 * Compact a multi-sentence summary to its first sentence — everything up to the
 * first `. ` (period-space) or the first newline, whichever comes first. Returns
 * the full text untouched when no sensible cut-point exists, so the result is
 * never less readable than the original. Keeps the trailing period when cutting
 * on `. ` so the line ends naturally.
 *
 * Shared by the top-level `--help` subcommand listing (composition) and the
 * `help-json --terse` discovery projection (use-cases) so both compact the same
 * way — the terse manifest in particular must stay token-cheap per category.
 */
export const firstSentence = (full: string): string => {
  const newlineIdx = full.indexOf('\n');
  const periodIdx = full.search(/\. /);
  const candidates = [newlineIdx, periodIdx].filter((i) => i > 0);
  if (candidates.length === 0) return full;
  const cut = Math.min(...candidates);
  return full.charAt(cut) === '.' ? full.slice(0, cut + 1) : full.slice(0, cut);
};
