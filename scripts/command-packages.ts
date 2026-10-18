/*
 * Loads the packages that `gen-docs.ts` and `check-doc-numbers.ts` document, and joins
 * them into one surface.
 *
 * Until the docs split (phases 6 and 7) the root documents the read commands and the
 * write package together, so the product docs keep one surface of 211 commands while
 * the root registry holds the read commands only (package split, phase 5, open
 * question 4). Each `--package` names a module that exports `commandPackage`. The
 * registries are joined. The lifecycle entries, the description, the names and the
 * npm name come from the first package, so every entry is filled with its bins. A
 * command name in two packages must have the same meta (compared before the names
 * are filled), else the join fails.
 *
 * Transitional: the move of write into `write/` changes the path of the write package.
 * When `src/composition/write-package.ts` does not exist, the loader takes
 * `write/src/composition/write-package.ts`, so the move needs no edit of
 * `package.json`. Only this path has the fallback, so a typo in another path still
 * fails. Phase 5 removes it when the root names the new path.
 */

import { existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { CommandPackage } from '@ask-marcel/office-core/use-cases/commands/command-package.ts';
import type { Command } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';

/** A package as the docs scripts read it: only the meta of each command is used. */
type DocumentedPackage = CommandPackage<Command<never>>;

const WRITE_PACKAGE_BEFORE_MOVE = 'src/composition/write-package.ts';
const WRITE_PACKAGE_AFTER_MOVE = 'write/src/composition/write-package.ts';

const packagePath = (path: string): string => (path === WRITE_PACKAGE_BEFORE_MOVE && !existsSync(path) ? WRITE_PACKAGE_AFTER_MOVE : path);

const isCommandPackage = (value: unknown): value is DocumentedPackage =>
  typeof value === 'object' && value !== null && 'registry' in value && 'lifecycle' in value && 'packageName' in value && 'binName' in value && 'bins' in value;

/** Imports each path (relative to the working directory) and takes its `commandPackage`. */
export const loadPackages = async (paths: ReadonlyArray<string>): Promise<Result<ReadonlyArray<DocumentedPackage>, string>> => {
  const packages: DocumentedPackage[] = [];
  for (const path of paths) {
    const loaded: unknown = (await import(pathToFileURL(resolve(packagePath(path))).href)).commandPackage;
    if (!isCommandPackage(loaded)) return err(`${path} exports no \`commandPackage\`.`);
    packages.push(loaded);
  }
  return ok(packages);
};

/** Joins the registries; everything else comes from the first package. */
export const mergePackages = (packages: ReadonlyArray<DocumentedPackage>): Result<DocumentedPackage, string> => {
  const [first, ...others] = packages;
  if (first === undefined) return err('no package to document.');
  const registry: Record<string, Command<never>> = { ...first.registry };
  for (const other of others) {
    for (const [name, command] of Object.entries(other.registry)) {
      const held = registry[name];
      if (held !== undefined && !Bun.deepEquals(held.meta, command.meta)) return err(`\`${name}\` is in two packages with different entries (${first.packageName}, ${other.packageName}).`);
      registry[name] = command;
    }
  }
  return ok({ ...first, registry });
};

/** Loads and joins the packages of the `--package` arguments. */
export const loadDocumentedPackage = async (paths: ReadonlyArray<string>): Promise<Result<DocumentedPackage, string>> => {
  const loaded = await loadPackages(paths);
  return loaded.ok ? mergePackages(loaded.value) : loaded;
};
