import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { ModuleLoader } from '../use-cases/ports/module-loader.ts';

export type ModuleLoaderFake = ModuleLoader & {
  // The files that the loader was asked to run, in order.
  readonly loaded: () => ReadonlyArray<string>;
};

// A loader that runs nothing and records each file. With a `failure`, each
// load fails with that message.
export const createModuleLoaderFake = (failure?: string): ModuleLoaderFake => {
  const files: string[] = [];
  return {
    load: async (file) => {
      files.push(file);
      return failure === undefined ? ok(undefined) : err({ type: 'load_failed', message: failure });
    },
    loaded: () => files,
  };
};
