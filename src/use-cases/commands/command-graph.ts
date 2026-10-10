import { readGraphOf } from '../../infra/read-graph.ts';
import { writeGraphOf } from '../../infra/write-graph.ts';
import type { Command, ReadCommand, RegisteredCommand } from './command-types.ts';

/*
 * Each command runs on the graph its effect allows (package split, D9). The
 * type of `execute` already says so; this wrap makes it true at run time too,
 * whatever graph a caller passes (the single package's full client holds
 * both): a read command gets the read view, a write command the write view.
 * Applied at registry assembly, the one point the CLI, the MCP gateway and a
 * library caller all pass through.
 */

const isReadCommand = (command: RegisteredCommand): command is ReadCommand => command.meta.effect === 'read';

const withOwnGraph = (command: RegisteredCommand): Command => {
  if (isReadCommand(command)) return { ...command, execute: async (graph, params) => command.execute(readGraphOf(graph), params) };
  return { ...command, execute: async (graph, params) => command.execute(writeGraphOf(graph), params) };
};

export { withOwnGraph };
