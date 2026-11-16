import type { Result } from '@ask-marcel/office-core/domain/result.ts';

export type PackageLocateError = { readonly type: 'not_installed' };

// Where an installed package is: the path of its package.json, looked up from
// the folder of the facade, as a package manager laid the packages out.
export type PackageLocator = {
  readonly manifestPath: (packageName: string) => Promise<Result<string, PackageLocateError>>;
};
