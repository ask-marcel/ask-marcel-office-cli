import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { Command } from './command-types.ts';

/*
 * One rule for unknown parameters, applied once, for every surface.
 *
 * Commander already rejected unknown `--flags` on the CLI, but Zod's default is
 * to STRIP unknown keys, so the two surfaces that hand params straight to a
 * schema — the MCP gateway's `run-command` and a library caller invoking
 * `commands[x].execute(...)` — silently swallowed them and returned data that
 * looked like it had obeyed. The 2026-07-24 delta audit found this on 5 of 7
 * delta commands.
 *
 * The allowed set is derived from the command's own Zod shape, so it cannot
 * drift from what the command actually reads. `.strict()` per schema would give
 * the same rejection with 184 touch points and a worse message; this wraps the
 * registry once instead.
 *
 * The rejection names the flag in CLI spelling (`--filter`), because that is
 * how every caller reads the docs, and carries `unknown_parameter` so an agent
 * can branch on a code rather than a substring.
 */

const camelToKebab = (value: string): string => value.replaceAll(/[A-Z]/g, (c) => `-${c.toLowerCase()}`);

const allowedKeys = (schema: Command['schema']): ReadonlyArray<string> => {
  // Every registry command is built on a z.object; anything else declares no
  // keys, which would reject every param, so treat a non-object shape as
  // "cannot check" and allow it through to the command's own validation.
  if (!(schema instanceof z.ZodObject)) return [];
  return Object.keys(schema.shape);
};

const unknownParamError = (unknown: ReadonlyArray<string>, allowed: ReadonlyArray<string>): GraphError => {
  const named = unknown.map((k) => `--${camelToKebab(k)}`).join(', ');
  const supported = allowed
    .map((k) => `--${camelToKebab(k)}`)
    .toSorted((a, b) => a.localeCompare(b))
    .join(', ');
  return {
    type: 'validation_error',
    code: 'unknown_parameter',
    message: `${named} ${unknown.length === 1 ? 'is not a parameter' : 'are not parameters'} of this command, so it would have been ignored rather than applied. Supported: ${supported || '(none)'}.`,
  };
};

/**
 * Returns the rejection for any param key the command does not declare, or
 * `undefined` when every key is known. Exported for the registry wrapper.
 */
const rejectUnknownParams = (schema: Command['schema'], params: Record<string, string>): Result<never, GraphError> | undefined => {
  const allowed = allowedKeys(schema);
  if (allowed.length === 0) return undefined;
  const unknown = Object.keys(params).filter((key) => !allowed.includes(key));
  if (unknown.length === 0) return undefined;
  return err(unknownParamError(unknown, allowed));
};

/**
 * Wraps a command so `execute` / `executeLocal` refuse unknown params before
 * running. Applied at registry assembly, which is the single choke point every
 * surface passes through — including a library caller that never touches
 * composition.
 */
const withUnknownParamRejection = (command: Command): Command => {
  const { executeLocal } = command;
  return {
    ...command,
    execute: async (graph, params) => rejectUnknownParams(command.schema, params) ?? (await command.execute(graph, params)),
    ...(executeLocal === undefined ? {} : { executeLocal: async (fs, params) => rejectUnknownParams(command.schema, params) ?? (await executeLocal(fs, params)) }),
  };
};

export { rejectUnknownParams, withUnknownParamRejection };
