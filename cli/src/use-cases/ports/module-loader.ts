import type { Result } from '@ask-marcel/office-core/domain/result.ts';

export type ModuleLoadError = { readonly type: 'load_failed'; readonly message: string };

// Runs the module at a file path in this process. A bin is such a module: it
// reads process.argv and does its work when it loads.
export type ModuleLoader = {
  readonly load: (file: string) => Promise<Result<void, ModuleLoadError>>;
};
