import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';

/*
 * The name of a new mail folder, checked before any read. A slash is refused
 * because --parent reads it as the step from a folder to its child folder, so
 * a folder with a slash in its name cannot be a --parent later. A backslash
 * and a control character are refused because Outlook does not show them in a
 * folder name.
 */

/** The longest folder name that this CLI sends. */
export const MAX_FOLDER_NAME = 255;

const BAD_CHARACTER = /[/\\\p{Cc}]/u;

const refused = (message: string): Result<string, GraphError> => err({ type: 'validation_error', code: 'invalid_folder_name', message });

/** The name with no space around it, or the refusal that says what is wrong. */
export const parseMailFolderName = (raw: string): Result<string, GraphError> => {
  const name = raw.trim();
  if (name === '') return refused('--name: give the name of the new folder.');
  if (name.length > MAX_FOLDER_NAME) return refused(`--name: a folder name has at most ${MAX_FOLDER_NAME} characters. This name has ${name.length}.`);
  if (BAD_CHARACTER.test(name))
    return refused('--name: a folder name cannot have "/", "\\" or a control character. --parent reads "/" as the step from a folder to its child folder.');
  return ok(name);
};
