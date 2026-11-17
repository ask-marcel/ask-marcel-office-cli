import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';

/*
 * The params of an MCP call, in the form that a command reads.
 *
 * A command reads camelCase keys and string values (`{ messageId: 'x', top: '5' }`),
 * which is what commander gives the CLI. An MCP client has only the docs and the
 * error messages, and they name a parameter as a flag (`--message-id`). So the
 * gateway accepts the flag name, with or without its `--`, and the camelCase
 * key, and it accepts a number or a boolean as a value (QA 2026-11-16, P2-1).
 *
 * This is one name in more than one spelling, not an alias: the 2026-07-24
 * decision (one name for each flag) stays. Two spellings of one key in one call
 * are refused, because one of the two values would be lost with no message.
 */

type RawParamValue = string | number | boolean;

type DuplicateParam = { readonly type: 'duplicate_parameter'; readonly first: string; readonly second: string };

const canonicalKey = (key: string): string => key.replace(/^--/, '').replaceAll(/-(.)/g, (_match, letter: string) => letter.toUpperCase());

const canonicalParams = (raw: Readonly<Record<string, RawParamValue>>): Result<Record<string, string>, DuplicateParam> => {
  const spellings = new Map<string, string>();
  const entries: Array<[string, string]> = [];
  for (const [key, value] of Object.entries(raw)) {
    const canonical = canonicalKey(key);
    const first = spellings.get(canonical);
    if (first !== undefined) return err({ type: 'duplicate_parameter', first, second: key });
    spellings.set(canonical, key);
    entries.push([canonical, String(value)]);
  }
  return ok(Object.fromEntries(entries));
};

export { canonicalParams };
export type { RawParamValue };
