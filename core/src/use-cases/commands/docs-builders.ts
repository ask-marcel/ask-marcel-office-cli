/*
 * The docs builders: the `help-json` manifest (full and terse), the category
 * filter and the `docs <command>` page. They are core (package split, step 13):
 * each package hands them its registry, its lifecycle entries and its npm name
 * as a `CommandPackage`, so they import no registry and no lifecycle entry.
 */
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { CommandPackage } from './command-package.ts';
import type { CommandCategory, Command, CommandMeta } from './command-types.ts';
import type { CommandManifest, CommandManifestEntry } from './docs-render.ts';
import { CATEGORY_LABELS, renderCommandMarkdown } from './docs-render.ts';
import { firstSentence } from './first-sentence.ts';

export type DocsError = { type: 'unknown_command'; readonly name: string; readonly available: ReadonlyArray<string> };

/**
 * Terse manifest entry — only the fields an LLM needs to *discover* a command
 * (i.e. "does this CLI do X?"). Drops `options`, `example`, `graphPathTemplate`,
 * `graphDocsUrl`, `responseShape`, `bodyTemplate`, `paginationStrategy`,
 * `scopesRequired` — everything the LLM only needs once it's already decided
 * to invoke. `stability` is kept (it's a discovery-time concern: LLMs prefer
 * stable siblings when they exist, so they need to see the tag at discovery
 * time, not after a second full-manifest fetch). /§6.
 */
export type TerseManifestEntry = {
  readonly name: string;
  readonly summary: string;
  readonly category: CommandCategory;
  readonly stability?: CommandMeta['stability'];
};

export type TerseManifest = {
  readonly package: string;
  readonly version: string;
  readonly generatedAt: string;
  readonly commands: ReadonlyArray<TerseManifestEntry>;
};

export type ManifestFilterError = { readonly type: 'unknown_category'; readonly category: string; readonly available: ReadonlyArray<string> };

const toEntry = (name: string, cmd: Command): CommandManifestEntry => {
  // Default every `pagination: true` command to `nextLink` strategy when the
  // command file doesn't specify one — that's the standard $top + nextLink
  // shape, true of ~80% of paginated commands. Audit-makes
  // the manifest field always populated on paginated commands so LLM
  // consumers don't need to read prose to learn the cursor mechanism.
  const paginationStrategy = cmd.meta.paginationStrategy ?? (cmd.meta.pagination ? 'nextLink' : undefined);
  // Each command declares its own scopes (the central graph-scopes.ts map was
  // dissolved into the command files at package-split step 11).
  const scopesRequired = cmd.meta.scopesRequired;
  return {
    name,
    summary: cmd.meta.summary,
    category: cmd.meta.category,
    graphMethod: cmd.meta.graphMethod,
    graphPathTemplate: cmd.meta.graphPathTemplate,
    graphDocsUrl: cmd.meta.graphDocsUrl,
    options: cmd.meta.options,
    example: cmd.meta.example,
    ...(cmd.meta.positionalArguments ? { positionalArguments: cmd.meta.positionalArguments } : {}),
    ...(cmd.meta.responseShape ? { responseShape: cmd.meta.responseShape } : {}),
    ...(cmd.meta.bodyTemplate ? { bodyTemplate: cmd.meta.bodyTemplate } : {}),
    ...(cmd.meta.pagination ? { pagination: cmd.meta.pagination } : {}),
    ...(paginationStrategy ? { paginationStrategy } : {}),
    ...(scopesRequired && scopesRequired.length > 0 ? { scopesRequired } : {}),
    ...(cmd.meta.needsElevatedToken ? { needsElevatedToken: cmd.meta.needsElevatedToken } : {}),
    ...(cmd.meta.needsSubstrateToken ? { needsSubstrateToken: cmd.meta.needsSubstrateToken } : {}),
    ...(cmd.meta.producesBytes ? { producesBytes: cmd.meta.producesBytes } : {}),
    ...(cmd.meta.producesMedia ? { producesMedia: cmd.meta.producesMedia } : {}),
    effect: cmd.meta.effect,
    ...(cmd.meta.stability ? { stability: cmd.meta.stability } : {}),
  };
};

/** What a manifest is built from: the package's registry, its lifecycle entries and its npm name. */
type ManifestSource = Pick<CommandPackage, 'registry' | 'lifecycle' | 'packageName'>;

/** What one docs page is looked up in. */
type DocsSource = Pick<CommandPackage, 'registry' | 'lifecycle'>;

const buildEntries = ({ registry, lifecycle }: DocsSource): ReadonlyArray<CommandManifestEntry> =>
  [...Object.entries(registry).map(([name, cmd]) => toEntry(name, cmd)), ...lifecycle].toSorted((a, b) => a.name.localeCompare(b.name));

export const buildManifest = (source: ManifestSource, version: string, now: () => Date = () => new Date()): CommandManifest => ({
  package: source.packageName,
  version,
  generatedAt: now().toISOString(),
  commands: buildEntries(source),
});

/**
 * Terse manifest — `{ name, summary, category }` per command. Roughly 95%
 * smaller than the full manifest (no options/example/Graph endpoint per entry).
 * Use `help-json --terse` to surface this to an LLM as the discovery view.
 */
export const buildTerseManifest = (source: ManifestSource, version: string, now: () => Date = () => new Date()): TerseManifest => ({
  package: source.packageName,
  version,
  generatedAt: now().toISOString(),
  commands: buildEntries(source).map((e) => ({
    name: e.name,
    // Terse is the per-category discovery view — keep it token-cheap by
    // projecting each multi-sentence summary to its first sentence (the full
    // summary stays in the unflagged manifest). Keeps drive/mail under budget.
    summary: firstSentence(e.summary),
    category: e.category,
    ...(e.stability ? { stability: e.stability } : {}),
  })),
});

/**
 * Filter a `CommandManifest` (or terse variant) down to a single category.
 * Returns `err({ type: 'unknown_category', ... })` if the requested category
 * isn't a known one — the CLI surfaces this through the standard error
 * envelope rather than silently returning an empty list.
 */
export const filterManifestByCategory = <M extends { readonly commands: ReadonlyArray<{ readonly category: CommandCategory }> }>(
  manifest: M,
  category: string
): Result<M, ManifestFilterError> => {
  const knownCategories = Object.keys(CATEGORY_LABELS).toSorted((a, b) => a.localeCompare(b));
  if (!knownCategories.includes(category)) {
    return err({ type: 'unknown_category', category, available: knownCategories });
  }
  const filtered = manifest.commands.filter((c) => c.category === category);
  return ok({ ...manifest, commands: filtered });
};

export const renderSingleCommand = ({ registry, lifecycle }: DocsSource, name: string): Result<string, DocsError> => {
  const cmd = registry[name];
  if (cmd) return ok(renderCommandMarkdown(toEntry(name, cmd)));
  const entry = lifecycle.find((candidate) => candidate.name === name);
  if (entry) return ok(renderCommandMarkdown(entry));
  return err({ type: 'unknown_command', name, available: [...Object.keys(registry), ...lifecycle.map((e) => e.name)].toSorted((a, b) => a.localeCompare(b)) });
};
