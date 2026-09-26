/**
 * The names closest to one a caller got wrong, for a "did you mean" line: an agent
 * that guessed `list-chat-messages` or `--folder-id` learns `list-teams-chat-messages`
 * or `--mail-folder-id` from the error instead of a round-trip through the docs.
 *
 * Two kinds of near miss. A word match: every word of the guess appears in the
 * candidate, in order, where a candidate word may extend a guessed one (`chat` in
 * `chats`); fewer extra words rank first. A typo: within three single-character
 * edits of the whole name. Word matches rank before typos; ties go to the shorter
 * name, then to the order the candidates came in.
 */

const MAX_TYPO_EDITS = 3;
const MAX_SUGGESTIONS = 3;

// Candidates are registry names and flags, kebab-case by construction; only the
// caller's guess needs its case folded (see closestNames).
const wordsOf = (name: string): ReadonlyArray<string> =>
  name
    .replace(/^--/, '')
    .split('-')
    .filter((word) => word !== '');

/** Every guessed word appears in the candidate, in order; a candidate word may extend a guessed one. */
const holdsWordsInOrder = (guessed: ReadonlyArray<string>, candidate: ReadonlyArray<string>): boolean => {
  let at = 0;
  for (const word of guessed) {
    while (at < candidate.length && !candidate[at].startsWith(word)) at += 1;
    if (at === candidate.length) return false;
    at += 1;
  }
  return true;
};

const editDistance = (a: string, b: string): number => {
  let previous = Array.from({ length: b.length + 1 }, (_unused, j) => j);
  for (let i = 1; i <= a.length; i += 1) {
    const current = [i];
    for (let j = 1; j <= b.length; j += 1) current.push(Math.min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)));
    previous = current;
  }
  return previous[b.length];
};

type Ranked = { readonly name: string; readonly kind: number; readonly score: number };

const rank = (guessed: ReadonlyArray<string>, name: string): Ranked | undefined => {
  const candidate = wordsOf(name);
  // Fewer words in a match means fewer words beyond the guessed ones.
  if (holdsWordsInOrder(guessed, candidate)) return { name, kind: 0, score: candidate.length };
  const edits = editDistance(guessed.join('-'), candidate.join('-'));
  return edits <= MAX_TYPO_EDITS ? { name, kind: 1, score: edits } : undefined;
};

const closestNames = (wanted: string, candidates: ReadonlyArray<string>): ReadonlyArray<string> => {
  const guessed = wordsOf(wanted.toLowerCase());
  if (guessed.length === 0) return [];
  return candidates
    .map((name) => rank(guessed, name))
    .filter((r): r is Ranked => r !== undefined)
    .toSorted((a, b) => a.kind - b.kind || a.score - b.score || a.name.length - b.name.length)
    .slice(0, MAX_SUGGESTIONS)
    .map((r) => r.name);
};

/** ` Did you mean `a`, `b` or `c`?` for the closest names (led by a space, to follow a sentence), or '' when none is close. */
const didYouMean = (wanted: string, candidates: ReadonlyArray<string>, lead = 'Did you mean'): string => {
  const names = closestNames(wanted, candidates).map((name) => `\`${name}\``);
  if (names.length === 0) return '';
  const listed = names.length === 1 ? names[0] : `${names.slice(0, -1).join(', ')} or ${names.at(-1)}`;
  return ` ${lead} ${listed}?`;
};

export { closestNames, didYouMean };
