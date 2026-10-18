import type { PackageRole } from './command-directory.ts';
import type { Command } from './command-types.ts';
import type { CommandManifestEntry } from './docs-render.ts';

/**
 * What a package hands the registry-agnostic shells (the CLI builder, the MCP
 * gateway, run-registry-command, the docs builders): its commands and its names.
 * The shells never import a registry, so the same code serves auth, read and
 * write after the package split (docs/plans/2026-10-01-package-split.md, step 13).
 */
export type CommandPackage<C extends Command<never> = Command> = {
  /** The commands the package runs, by name. `C` names the graph they run on. */
  readonly registry: Readonly<Record<string, C>>;
  /**
   * The manifest entries of the commands that are not in the registry (`login`,
   * `docs`, `help-json`, `mcp`…), so `help-json` and `docs` cover them too.
   */
  readonly lifecycle: ReadonlyArray<CommandManifestEntry>;
  /**
   * What the package is, in its own words: the lead of the top-level `--help`
   * description. The CLI adds what the registry gives (the writes and what each
   * leaves behind, the read counts), so this text names no count and no command list.
   */
  readonly description: string;
  /** The command that the docs help gives as its example, e.g. `list-mail-messages`. */
  readonly exampleCommand: string;
  /** The bin a user types, e.g. `ask-marcel-office`. */
  readonly binName: string;
  /** The npm name, e.g. `ask-marcel-office-cli`; the manifest carries it. */
  readonly packageName: string;
  /**
   * The bin of each package of the split, as this package names it in a command
   * reference (`{auth-bin} login`, see bin-names.ts). The single package holds
   * all three, so all three are its own bin.
   */
  readonly bins: Readonly<Record<PackageRole, string>>;
};
