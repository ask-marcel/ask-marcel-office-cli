import type { GraphError } from './graph-error.ts';
import type { Result } from './result.ts';
import { err } from './result.ts';

/*
 * The beta check of the write graph (decision D7): a PATCH may go to Graph
 * beta only for a path on this list. The HTML `notes` of a Planner task exist
 * on beta only (v1.0 `plannerTaskDetails` has the plain `description`), so the
 * details PATCH is the one entry. The write graph and the write view of any
 * graph both send a beta PATCH through `patchOnAllowedRoot`.
 */

// A Planner id is base64url (letters, digits, `-`, `_`), so `..` or an escape
// cannot stand for one and move the PATCH to another beta path.
const BETA_PATCH_PATHS: ReadonlyArray<{ readonly template: string; readonly pattern: RegExp }> = [
  { template: '/planner/tasks/{task-id}/details', pattern: /^\/planner\/tasks\/[\w-]+\/details$/ },
];

const mayUseBeta = (path: string): boolean => BETA_PATCH_PATHS.some((entry) => entry.pattern.test(path));

const betaRefused = (path: string): Result<never, GraphError> =>
  err({
    type: 'validation_error',
    code: 'beta_refused',
    message: `PATCH ${path} was not sent to Graph beta: only ${BETA_PATCH_PATHS.map((entry) => entry.template).join(' and ')} may use beta. Send it without the beta option, on v1.0.`,
  });

/** Runs `patch` unless it asks for beta on a path that is not on the list; refuses it then, before any request. */
const patchOnAllowedRoot = async (path: string, api: 'beta' | undefined, patch: () => Promise<Result<unknown, GraphError>>): Promise<Result<unknown, GraphError>> =>
  api === 'beta' && !mayUseBeta(path) ? betaRefused(path) : patch();

export { patchOnAllowedRoot };
