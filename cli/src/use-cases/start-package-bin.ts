import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { FileSystem } from '@ask-marcel/office-core/use-cases/ports/filesystem.ts';
import { binFileOf } from '../domain/package-bin.ts';
import type { StartError } from '../domain/start-error.ts';
import type { ModuleLoader } from './ports/module-loader.ts';
import type { PackageLocator } from './ports/package-locator.ts';

export type Started = { readonly binFile: string };

export type StartPackageBin = (packageName: string) => Promise<Result<Started, StartError>>;

export type StartPackageBinDeps = {
  readonly locator: PackageLocator;
  readonly fs: Pick<FileSystem, 'readJson'>;
  readonly loader: ModuleLoader;
};

/*
 * Starts the bin of an installed package in this process (ADR 0005): finds the
 * package.json of the package, reads its one `bin` entry and runs that file.
 * The bin reads process.argv itself, so the arguments of the facade bin reach
 * it with no change.
 */
export const createStartPackageBin =
  (deps: StartPackageBinDeps): StartPackageBin =>
  async (packageName) => {
    const located = await deps.locator.manifestPath(packageName);
    if (!located.ok) return err({ type: 'package_not_installed', packageName });
    const manifest = await deps.fs.readJson<unknown>(located.value);
    if (!manifest.ok) return err({ type: 'manifest_unreadable', packageName });
    const binFile = binFileOf(located.value, manifest.value);
    if (!binFile.ok) return err({ type: binFile.error.type, packageName });
    const loaded = await deps.loader.load(binFile.value);
    if (!loaded.ok) return err({ type: 'bin_failed', packageName, message: loaded.error.message });
    return ok({ binFile: binFile.value });
  };
