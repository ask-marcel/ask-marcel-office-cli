import { pathToFileURL } from 'node:url';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { formatError } from '@ask-marcel/office-core/domain/utilities/format-error.ts';
import type { ModuleLoader } from '../use-cases/ports/module-loader.ts';

/*
 * Runs a module file with a dynamic import. The path goes in as a file URL: on
 * Windows an absolute path is not a valid import specifier. The import throws
 * for a file that is absent or that fails while it loads.
 */
export const createModuleLoader = (): ModuleLoader => ({
  load: async (file) => {
    try {
      await import(pathToFileURL(file).href);
      return ok(undefined);
    } catch (e) {
      return err({ type: 'load_failed', message: formatError(e) });
    }
  },
});
