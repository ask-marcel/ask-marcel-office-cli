/*
 * Loads the package that `gen-docs.ts` and `check-doc-numbers.ts` document.
 *
 * Since phase 7 of the package split, each package folder documents its own
 * package, so a run reads one package. `--package` names a module that exports
 * `commandPackage`, or the export that `--export` names (auth: `authPackage`).
 */

import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { CommandPackage } from '@ask-marcel/office-core/use-cases/commands/command-package.ts';
import type { Command } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';

/** A package as the docs scripts read it: only the meta of each command is used. */
type DocumentedPackage = CommandPackage<Command<never>>;

const isCommandPackage = (value: unknown): value is DocumentedPackage =>
  typeof value === 'object' && value !== null && 'registry' in value && 'lifecycle' in value && 'packageName' in value && 'binName' in value && 'bins' in value;

// A module that cannot be imported (a wrong path, a syntax error) is an error of
// the arguments, so the scripts exit 2 and not with the exit 1 of a crash.
const importModule = async (path: string): Promise<Result<Readonly<Record<string, unknown>>, string>> => {
  try {
    return ok((await import(pathToFileURL(resolve(path)).href)) as Readonly<Record<string, unknown>>);
  } catch (error) {
    return err(`${path} cannot be imported (${error instanceof Error ? error.message : String(error)}).`);
  }
};

/** Imports the module at `path` (relative to the working directory) and takes its export `exportName` (by default `commandPackage`). */
export const loadPackage = async (path: string, exportName = 'commandPackage'): Promise<Result<DocumentedPackage, string>> => {
  const imported = await importModule(path);
  if (!imported.ok) return imported;
  const loaded: unknown = imported.value[exportName];
  return isCommandPackage(loaded) ? ok(loaded) : err(`${path} exports no \`${exportName}\`.`);
};
