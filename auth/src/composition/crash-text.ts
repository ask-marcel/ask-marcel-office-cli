import { formatError } from '@ask-marcel/office-core/domain/utilities/format-error.ts';

/*
 * What the one crash catch of the auth bin (src/main.ts) reads from its argv
 * and its error. Kept out of main.ts so a test pins it: main.ts is an entry
 * point, and coverage skips it.
 */

export const isTokenInvocation = (argv: ReadonlyArray<string>): boolean => argv[2] === 'token';

// The token helper path tells only the error's kind, as dist/token.js does: a
// message can quote what was being parsed, a token included.
export const crashText = (argv: ReadonlyArray<string>, e: unknown): string => {
  if (!isTokenInvocation(argv)) return formatError(e);
  return `the token helper stopped on an unexpected ${e instanceof Error ? e.name : 'throw'}; its message is withheld.`;
};

// Commander parser errors are already rendered as a JSON envelope on stdout by
// exitOverride in the CLI builder: the catch writes no second line for them.
export const isCommanderError = (e: unknown): boolean =>
  e !== null && typeof e === 'object' && 'code' in e && typeof e.code === 'string' && (e as { code: string }).code.startsWith('commander.');
