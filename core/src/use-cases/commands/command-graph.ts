import type { ReadGraph, WriteGraph } from '../ports/graph.ts';
import type { Command, ReadCommand, RegisteredCommand, WriteCommand } from './command-types.ts';
import { postIfReadOnly } from '../../domain/read-only-post.ts';
import { unlessSendsMail } from '../../domain/mail-send-guard.ts';
import { patchOnAllowedRoot } from '../../domain/beta-patch.ts';
import type { GraphError } from '../../domain/graph-error.ts';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { LocalFileReader } from './read-local-file.ts';
import type { BoundPlan, GraphReads, PlanWrite } from './write-plan.ts';

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
 * its basic-tier members only, with the no-send check (mail-send-guard.ts) in
 * front of each change: a change to a path that sends mail is refused before
 * any request, and so is a beta PATCH of a path that is not on the beta
 * allow-list (beta-patch.ts). The command registry gives every write command this view, so no
 * write command holds an elevated, guest or chat reader, and none sends mail.
 * Each member calls the graph's own member on the graph, so a caller's graph
 * whose methods read `this` still works.
 */
const writeGraphOf = (graph: WriteGraph): WriteGraph => ({
  get: (path, extraHeaders) => graph.get(path, extraHeaders),
  getBinary: (path) => graph.getBinary(path),
  fetchUrl: (url) => graph.fetchUrl(url),
  post: (path, body) => unlessSendsMail('POST', path, () => graph.post(path, body)),
  patch: (path, body, options) => unlessSendsMail('PATCH', path, () => patchOnAllowedRoot(path, options?.api, () => graph.patch(path, body, options))),
  put: (basePath, body, contentType) => graph.put(basePath, body, contentType),
  delete: (path, options) => unlessSendsMail('DELETE', path, () => graph.delete(path, options)),
  uploadAttachment: (target, file) => graph.uploadAttachment(target, file),
});

/** A read command that runs on the read view of any read graph it is given. */
const readCommandOf = (command: ReadCommand): Command<ReadGraph> => ({ ...command, execute: async (graph, params) => command.execute(readGraphOf(graph), params) });

/**
 * The read members of a write graph, in a new object: a plan (write-plan.ts)
 * gets this, so no write member is in its reach, whatever it casts.
 */
const readsOf = (graph: WriteGraph): GraphReads => ({
  get: (path, extraHeaders) => graph.get(path, extraHeaders),
  getBinary: (path) => graph.getBinary(path),
  fetchUrl: (url) => graph.fetchUrl(url),
});

// The plan reads through the read members only; its writes run on the graph it was planned on.
const bindPlan = async (plan: PlanWrite, graph: WriteGraph, params: Record<string, string>, files?: LocalFileReader): Promise<Result<BoundPlan, GraphError>> => {
  const planned = await plan(readsOf(graph), params, files);
  if (!planned.ok) return planned;
  const { changes, skipped } = planned.value;
  return ok({ changes, ...(skipped === undefined ? {} : { skipped }), apply: () => planned.value.apply(graph) });
};

/** A write command that runs on the write view of any write graph it is given, its plan too. */
const writeCommandOf = (command: WriteCommand): Command<WriteGraph> => {
  const { plan, ...rest } = command;
  return {
    ...rest,
    execute: async (graph, params, files) => command.execute(writeGraphOf(graph), params, files),
    ...(plan === undefined
      ? {}
      : { plan: async (graph: WriteGraph, params: Record<string, string>, files?: LocalFileReader) => bindPlan(plan, writeGraphOf(graph), params, files) }),
  };
};

const isReadCommand = (command: RegisteredCommand): command is ReadCommand => command.meta.effect === 'read';

const withOwnGraph = (command: RegisteredCommand): Command => (isReadCommand(command) ? readCommandOf(command) : writeCommandOf(command));

export { readCommandOf, readGraphOf, readsOf, withOwnGraph, writeCommandOf, writeGraphOf };
