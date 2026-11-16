import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { PackageLocator } from '../use-cases/ports/package-locator.ts';

// The installed packages, as a map from the package name to the path of its
// package.json. A name that the map does not hold is not installed.
export const createPackageLocatorFake = (manifests: Readonly<Record<string, string>>): PackageLocator => ({
  manifestPath: async (packageName) => {
    const path = Object.hasOwn(manifests, packageName) ? manifests[packageName] : undefined;
    return path === undefined ? err({ type: 'not_installed' }) : ok(path);
  },
});
