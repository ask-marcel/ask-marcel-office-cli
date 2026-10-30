import type { ReadGraph, WriteGraph } from '../ports/graph.ts';
import type { Command, ReadCommand, RegisteredCommand, WriteCommand } from './command-types.ts';
import { postIfReadOnly } from '../../domain/read-only-post.ts';

/*
 * Each command runs on the graph its effect allows (package split, D9). The
 * type of `execute` already says so; this wrap makes it true at run time too,
 * whatever graph a caller passes (the single package's full client holds
 * both): a read command gets the read view, a write command the write view.
 * Applied at registry assembly, the one point the CLI, the MCP gateway and a
 * library caller all pass through. A package that holds one kind of command
 * wraps each one with `readCommandOf` or `writeCommandOf`, so it runs on the
 * graph of that kind and needs no full client (package split, phase 5).
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

/** A read command that runs on the read view of any read graph it is given. */
const readCommandOf = (command: ReadCommand): Command<ReadGraph> => ({ ...command, execute: async (graph, params) => command.execute(readGraphOf(graph), params) });

/** A write command that runs on the write view of any write graph it is given. */
const writeCommandOf = (command: WriteCommand): Command<WriteGraph> => ({ ...command, execute: async (graph, params) => command.execute(writeGraphOf(graph), params) });

const isReadCommand = (command: RegisteredCommand): command is ReadCommand => command.meta.effect === 'read';

const withOwnGraph = (command: RegisteredCommand): Command => (isReadCommand(command) ? readCommandOf(command) : writeCommandOf(command));

export { readCommandOf, readGraphOf, withOwnGraph, writeCommandOf, writeGraphOf };
