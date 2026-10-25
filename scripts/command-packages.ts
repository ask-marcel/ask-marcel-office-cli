/*
 * Loads the packages that `gen-docs.ts` and `check-doc-numbers.ts` document, and joins
 * them into one surface.
 *
 * Until the docs split (phase 7) the root docs describe the read and the write
 * packages together, so the product docs keep one surface of 211 commands while each
 * registry holds its own commands (package split, phase 5, open question 4). Each
 * `--package` names a module that exports `commandPackage`, or the export that
 * `--export` names (auth: `authPackage`). The registries are
 * joined. The lifecycle entries, the description, the names and the npm name come
 * from the first package, so every entry is filled with its bins. A command name in
 * two packages must have the same meta (compared before the names are filled), else
 * the join fails.
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

/** Imports each path (relative to the working directory) and takes its export `exportName` (by default `commandPackage`). */
export const loadPackages = async (paths: ReadonlyArray<string>, exportName = 'commandPackage'): Promise<Result<ReadonlyArray<DocumentedPackage>, string>> => {
  const packages: DocumentedPackage[] = [];
  for (const path of paths) {
    const loaded: unknown = (await import(pathToFileURL(resolve(path)).href))[exportName];
    if (!isCommandPackage(loaded)) return err(`${path} exports no \`${exportName}\`.`);
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
export const loadDocumentedPackage = async (paths: ReadonlyArray<string>, exportName?: string): Promise<Result<DocumentedPackage, string>> => {
  const loaded = await loadPackages(paths, exportName);
  return loaded.ok ? mergePackages(loaded.value) : loaded;
};
