import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';

/*
 * The key of a reference (a link or a file) of a Planner task. The key is the
 * URL of the reference, but OData does not permit `.`, `:`, `%`, `@` or `#` in
 * a property name, so Graph asks for each of them encoded: `%` first (to
 * `%25`, so a `%` of the URL is not read as a code), then `.` to `%2E`, `:`
 * to `%3A`, `@` to `%40` and `#` to `%23`. The other characters stay as they
 * are, as in the example of the Graph docs
 * (`https%3A//contoso%2Esharepoint%2Ecom/...`). Only an http or https URL is
 * a reference.
 */

const CODES: ReadonlyArray<readonly [string, string]> = [
  ['%', '%25'],
  ['.', '%2E'],
  [':', '%3A'],
  ['@', '%40'],
  ['#', '%23'],
];

/** The longest URL that this CLI gives as a reference. */
export const REFERENCE_URL_MAX = 2048;

/** The key of the reference to `url`. */
export const referenceKeyOf = (url: string): string => CODES.reduce((key, [character, code]) => key.replaceAll(character, code), url);

/** The URL of a reference key: the codes back to their characters, `%25` last. */
export const urlOfReferenceKey = (key: string): string => CODES.toReversed().reduce((url, [character, code]) => url.replaceAll(code, character), key);

const refused = (message: string): Result<never, GraphError> => err({ type: 'validation_error', code: 'invalid_reference_url', message });

/** Whether `url` is an http or https URL. */
export const isWebUrl = (url: string): boolean => {
  const parsed = URL.parse(url);
  return parsed !== null && (parsed.protocol === 'https:' || parsed.protocol === 'http:');
};

/** Checks the URL of a link: an http or https URL of at most 2048 characters. The refusal does not quote the URL. */
export const parseReferenceUrl = (raw: string, what: string): Result<string, GraphError> => {
  if (raw.length > REFERENCE_URL_MAX) return refused(`${what}: the URL has ${raw.length} characters, above the limit of ${REFERENCE_URL_MAX}.`);
  return isWebUrl(raw) ? ok(raw) : refused(`${what}: give a full http or https URL, for example https://example.com/page.`);
};
