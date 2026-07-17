import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { Command } from './command-types.ts';

/**
 * Name -> command lookup that also honours `meta.commandAliases`.
 *
 * The CLI gets alias resolution for free from commander (`commandDef.alias(...)`
 * in `cli.ts`), so nothing needed this until a second front end appeared. The
 * MCP gateway has no commander, and its three command-taking tools
 * (`get-command-docs`, `run-command`, `run-write-command`) each need the same
 * lookup — Rule of Three, so it lives here rather than inline.
 *
 * Returns the CANONICAL name alongside the command so every downstream concern
 * (docs rendering, the `mutates` read/write gate, error messages) keys off one
 * spelling regardless of which one the caller used.
 */
export type ResolveCommandError = { readonly type: 'unknown_command'; readonly name: string; readonly available: ReadonlyArray<string> };

export type ResolvedCommand = { readonly name: string; readonly command: Command };

const findByAlias = (registry: Readonly<Record<string, Command>>, name: string): ResolvedCommand | undefined => {
  for (const [canonical, command] of Object.entries(registry)) {
    if (command.meta.commandAliases?.includes(name)) return { name: canonical, command };
  }
  return undefined;
};

const availableNames = (registry: Readonly<Record<string, Command>>): ReadonlyArray<string> =>
  Object.entries(registry)
    .flatMap(([name, command]) => [name, ...(command.meta.commandAliases ?? [])])
    .toSorted((a, b) => a.localeCompare(b));

export const resolveCommand = (registry: Readonly<Record<string, Command>>, name: string): Result<ResolvedCommand, ResolveCommandError> => {
  // `Object.hasOwn` rather than a truthy `registry[name]`: a name like
  // `constructor` or `toString` would otherwise resolve to a prototype member
  // and be handed downstream as if it were a command.
  if (Object.hasOwn(registry, name)) return ok({ name, command: registry[name] });
  // Canonical wins over any alias that shadows it — a rename must never be
  // able to hijack a live command name.
  const aliased = findByAlias(registry, name);
  if (aliased) return ok(aliased);
  return err({ type: 'unknown_command', name, available: availableNames(registry) });
};
