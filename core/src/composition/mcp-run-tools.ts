import { describeWrites, WRITE_EFFECTS } from '../use-cases/commands/command-effect.ts';
import type { Command, CommandEffect } from '../use-cases/commands/command-types.ts';

/*
 * The run tools of the MCP gateway, one for each mark that a host reads: the
 * reads (`readOnlyHint`), the writes that are not destructive, and the writes
 * that delete (`destructiveHint`). A command runs only on the tool of its
 * effect, so a host that approves run-write-command never approves a delete.
 */

export type RunTool = 'run-command' | 'run-write-command' | 'run-destructive-command';

/** The tool that runs a command of this effect. */
export const runToolOf = (effect: CommandEffect): RunTool => {
  if (effect === 'read') return 'run-command';
  return WRITE_EFFECTS[effect].destructive ? 'run-destructive-command' : 'run-write-command';
};

/** The command names of each run tool, and of all the writes, sorted. */
export type RunToolNames = Readonly<Record<RunTool | 'writes', ReadonlyArray<string>>>;

const sorted = (names: ReadonlyArray<string>): ReadonlyArray<string> => names.toSorted((a, b) => a.localeCompare(b));

export const runToolNames = (registry: Readonly<Record<string, Command<never>>>): RunToolNames => {
  const entries = Object.entries(registry);
  const on = (tool: RunTool): ReadonlyArray<string> => entries.filter(([, command]) => runToolOf(command.meta.effect) === tool).map(([name]) => name);
  return {
    'run-command': on('run-command'),
    'run-write-command': sorted(on('run-write-command')),
    'run-destructive-command': sorted(on('run-destructive-command')),
    writes: sorted(entries.filter(([, command]) => command.meta.effect !== 'read').map(([name]) => name)),
  };
};

/** Why the tool `called` does not run the command `name`, whose tool is `own`; undefined when it does. */
export const wrongToolMessage = (name: string, own: RunTool, called: RunTool, names: RunToolNames): string | undefined => {
  if (own === called) return undefined;
  // Gate BEFORE executing: a write routed through run-command would silently
  // mutate under a `readOnlyHint: true` tool the client may have auto-approved.
  if (called === 'run-command')
    return `"${name}" writes to Microsoft 365, so it is not available through run-command (which is declared read-only). Use ${own} instead. Write commands: ${names.writes.join(', ')}.`;
  if (own === 'run-command') return `"${name}" is a read command — use run-command. ${called} only accepts: ${names[called].join(', ')}.`;
  const deletes = names['run-destructive-command'].join(', ');
  // The same gate for the destructive mark: a delete never runs under a tool that is not marked destructive.
  if (own === 'run-destructive-command')
    return `"${name}" deletes, so it is not available through run-write-command (which is not declared destructive). Use run-destructive-command instead. run-destructive-command accepts: ${deletes}.`;
  return `"${name}" does not delete — use run-write-command. run-destructive-command only accepts: ${deletes}.`;
};

type Registry = Readonly<Record<string, Command<never>>>;

/** The run tools that a package with these commands offers, as a description names them: `run-command or run-write-command`. */
export const runToolsPhrase = (names: RunToolNames): string => {
  const offered: ReadonlyArray<RunTool> = ['run-command', 'run-write-command', 'run-destructive-command'];
  return offered.filter((tool) => names[tool].length > 0).join(' or ');
};

/** Where run-command sends a write: `elsewhere` names the MCP server of the write package, for a package with no write. */
export const writesElsewhere = (names: RunToolNames, elsewhere: string): string => {
  if (names.writes.length === 0) return `the write commands are on ${elsewhere}`;
  const tools = names['run-destructive-command'].length > 0 ? 'run-write-command and run-destructive-command' : 'run-write-command';
  return `the ${names.writes.length} write commands are rejected here and live on ${tools}`;
};

// The commands of `names`, in the order of the registry: a description names the write classes in that order.
const only = (registry: Registry, names: ReadonlyArray<string>): Registry => Object.fromEntries(Object.entries(registry).filter(([name]) => names.includes(name)));

/** The description of run-write-command; `readsElsewhere` says where the reads are. */
export const writeToolDescription = (registry: Registry, names: RunToolNames, readsElsewhere: string): string => {
  const plain = names['run-write-command'];
  const deletes = names['run-destructive-command'];
  const surface = deletes.length > 0 ? `The commands that delete are on run-destructive-command: ${deletes.join(', ')}.` : 'This CLI has no other write surface.';
  return `Run one of the ${plain.length} commands that WRITE to Microsoft 365. ${describeWrites(only(registry, plain))}. ${surface} ${readsElsewhere}`;
};

/** The description of run-destructive-command; `readsTool` names the tool or the server of the reads. */
export const deleteToolDescription = (registry: Registry, names: RunToolNames, readsTool: string): string => {
  const deletes = names['run-destructive-command'];
  return `Run one of the ${deletes.length} commands that DELETE in Microsoft 365. ${describeWrites(only(registry, deletes))}. Each one takes dryRun, and each run has an undo unless the words above say that it has none. The other writes are on run-write-command, and the reads on ${readsTool}.`;
};
