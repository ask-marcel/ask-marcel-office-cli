import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';

/*
 * The id of a mail message, a mail folder or a conversation, as it goes into a
 * Graph path or a `$filter` (hard rule 12). Graph gives these ids in base64url
 * with `=` padding, so only letters, digits, `_`, `-` and `=` pass: no `/`, `?`,
 * `#`, `%`, quote or space can change the path or the filter.
 */

export type GraphId = string & { readonly __brand: 'GraphId' };

const GRAPH_ID = /^[\w=-]{1,512}$/;

/** Checks an id from a param, a file or the run journal. `what` names it in the refusal, for example `--message-id`. */
export const parseGraphId = (raw: string, what: string): Result<GraphId, GraphError> =>
  GRAPH_ID.test(raw)
    ? ok(raw as GraphId)
    : err({
        type: 'validation_error',
        code: 'invalid_id',
        message: `${what}: "${raw.slice(0, 40)}" is not a Graph id. An id has only letters, digits, "_", "-" and "=", as list-mail-folder-messages and get-mail-message return it.`,
      });

/** The id check as a zod schema, for the ids of the run journal (a file that a person can edit). */
export const graphIdSchema = z
  .string()
  .regex(GRAPH_ID)
  .transform((value) => value as GraphId);

/** Asserts an id that a Graph answer gave: Graph is the source, so a bad id is a defect, not an input. */
export const graphId = (value: string): GraphId => {
  const parsed = parseGraphId(value, 'Graph answer');
  if (!parsed.ok) throw new Error(parsed.error.message);
  return parsed.value;
};
