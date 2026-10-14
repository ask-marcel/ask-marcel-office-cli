/*
 * The docs builders with the single package's lifecycle entries built in: the
 * call shape `buildManifest(registry, packageName, version)` that the
 * real-registry suites have always used. The builders are core and take a
 * `CommandPackage` (`docs-builders.ts`); the shells and the docs scripts call
 * them directly. This module goes with the read package at the split.
 */
import type { Result } from '../../domain/result.ts';
import type { Command } from './command-types.ts';
import type { DocsError, TerseManifest } from './docs-builders.ts';
import { buildManifest as buildPackageManifest, buildTerseManifest as buildPackageTerseManifest, renderSingleCommand as renderPackageCommand } from './docs-builders.ts';
import type { CommandManifest } from './docs-render.ts';
import { LIFECYCLE_ENTRIES } from './lifecycle-entries.ts';

type Registry = Readonly<Record<string, Command>>;

export const buildManifest = (registry: Registry, packageName: string, version: string, now?: () => Date): CommandManifest =>
  buildPackageManifest({ registry, lifecycle: LIFECYCLE_ENTRIES, packageName }, version, now);

export const buildTerseManifest = (registry: Registry, packageName: string, version: string, now?: () => Date): TerseManifest =>
  buildPackageTerseManifest({ registry, lifecycle: LIFECYCLE_ENTRIES, packageName }, version, now);

export const renderSingleCommand = (registry: Registry, name: string): Result<string, DocsError> => renderPackageCommand({ registry, lifecycle: LIFECYCLE_ENTRIES }, name);

export { filterManifestByCategory } from './docs-builders.ts';
export type { DocsError, ManifestFilterError, TerseManifest, TerseManifestEntry } from './docs-builders.ts';
