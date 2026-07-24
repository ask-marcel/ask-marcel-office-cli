import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { Command } from './command-types.ts';

/**
 * Name -> command lookup shared by every non-commander front end.
 *
 * The MCP gateway's three command-taking tools (`get-command-docs`,
 * `run-command`, `run-write-command`) each need the same lookup — Rule of
 * Three, so it lives here rather than inline.
 *
 * 2026-07-24: one name per command. Deprecated-name resolution
 * (`meta.commandAliases`) was removed with the alias system; an old name gets
 * the unknown-command rejection with the full list, same as any typo.
 */
export type ResolveCommandError = { readonly type: 'unknown_command'; readonly name: string; readonly available: ReadonlyArray<string> };

export type ResolvedCommand = { readonly name: string; readonly command: Command };

const availableNames = (registry: Readonly<Record<string, Command>>): ReadonlyArray<string> => Object.keys(registry).toSorted((a, b) => a.localeCompare(b));

export const resolveCommand = (registry: Readonly<Record<string, Command>>, name: string): Result<ResolvedCommand, ResolveCommandError> => {
  // `Object.hasOwn` rather than a truthy `registry[name]`: a name like
  // `constructor` or `toString` would otherwise resolve to a prototype member
  // and be handed downstream as if it were a command.
  if (Object.hasOwn(registry, name)) return ok({ name, command: registry[name] });
  return err({ type: 'unknown_command', name, available: availableNames(registry) });
};
