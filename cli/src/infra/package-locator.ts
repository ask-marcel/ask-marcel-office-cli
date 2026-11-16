import { createRequire } from 'node:module';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { PackageLocator } from '../use-cases/ports/package-locator.ts';

/*
 * Asks the runtime where a package is, from the file that holds this code: in
 * the built facade that file is a bin of the facade, so the lookup starts in
 * the folder where a package manager put the facade and its dependencies. Each
 * package exports its `./package.json`. `require.resolve` does the lookup on
 * every Node 20 (`import.meta.resolve` needs a flag before 20.6), and it
 * throws for a package that it cannot find.
 */
export const createPackageLocator = (): PackageLocator => {
  const lookup = createRequire(import.meta.url);
  return {
    manifestPath: async (packageName) => {
      try {
        return ok(lookup.resolve(`${packageName}/package.json`));
      } catch {
        return err({ type: 'not_installed' });
      }
    },
  };
};
