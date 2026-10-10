import type { ReadGraph } from '../../infra/read-graph.ts';
import type { WriteGraph } from '../../infra/write-graph.ts';
import type { Command, ReadCommand, RegisteredCommand } from './command-types.ts';
import { postIfReadOnly } from './read-only-post.ts';

/*
 * Each command runs on the graph its effect allows (package split, D9). The
 * type of `execute` already says so; this wrap makes it true at run time too,
 * whatever graph a caller passes (the single package's full client holds
 * both): a read command gets the read view, a write command the write view.
 * Applied at registry assembly, the one point the CLI, the MCP gateway and a
 * library caller all pass through.
 */

/**
 * The read graph inside any graph, the single package's full client included:
 * its read members only, with the POST check in front. The command registry
 * gives every read command this view, so no read command holds a member that
 * writes, whichever graph its caller passed. Each member calls the graph's own
 * member on the graph, so a caller's graph whose methods read `this` still works.
 */
const readGraphOf = (graph: ReadGraph): ReadGraph => ({
  get: (path, extraHeaders) => graph.get(path, extraHeaders),
  getElevated: (path) => graph.getElevated(path),
  getGuest: (path, tenant) => graph.getGuest(path, tenant),
  getBinaryGuest: (path, tenant) => graph.getBinaryGuest(path, tenant),
  discoverTenantId: (spoHost) => graph.discoverTenantId(spoHost),
  teamsChat: (path) => graph.teamsChat(path),
  teamsChatIc3: (path) => graph.teamsChatIc3(path),
  teamsChatMedia: (url) => graph.teamsChatMedia(url),
  post: (path, body) => postIfReadOnly((readOnly, query) => graph.post(readOnly, query), path, body),
  getBinary: (path) => graph.getBinary(path),
  getBinaryElevated: (path) => graph.getBinaryElevated(path),
  fetchUrl: (url) => graph.fetchUrl(url),
});

/**
 * The write graph inside any graph, the single package's full client included:
 * its basic-tier members only. The command registry gives every write command
 * this view, so no write command holds an elevated, guest or chat reader. Each
 * member calls the graph's own member on the graph, so a caller's graph whose
 * methods read `this` still works.
 */
const writeGraphOf = (graph: WriteGraph): WriteGraph => ({
  get: (path, extraHeaders) => graph.get(path, extraHeaders),
  getBinary: (path) => graph.getBinary(path),
  fetchUrl: (url) => graph.fetchUrl(url),
  post: (path, body) => graph.post(path, body),
  patch: (path, body) => graph.patch(path, body),
  put: (basePath, body, contentType) => graph.put(basePath, body, contentType),
  delete: (path) => graph.delete(path),
});

const isReadCommand = (command: RegisteredCommand): command is ReadCommand => command.meta.effect === 'read';

const withOwnGraph = (command: RegisteredCommand): Command => {
  if (isReadCommand(command)) return { ...command, execute: async (graph, params) => command.execute(readGraphOf(graph), params) };
  return { ...command, execute: async (graph, params) => command.execute(writeGraphOf(graph), params) };
};

export { readGraphOf, withOwnGraph, writeGraphOf };
