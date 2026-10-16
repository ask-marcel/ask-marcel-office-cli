import type { Result } from '../../domain/result.ts';
import { err } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-request.ts';

/*
 * The POST check of the read graph (package split, D9): a read command may
 * POST only to the two endpoints that change nothing. The path type refuses
 * any other path at compile time; `postIfReadOnly` refuses it at run time, so
 * a cast cannot turn a read into a write. The read graph and the read view of
 * a full graph both send their POST through it.
 */

// The two endpoints a read command POSTs to; both only answer a query.
const READ_ONLY_POST_PATHS = ['/search/query', '/me/calendar/getSchedule'] as const;

type ReadOnlyPostPath = (typeof READ_ONLY_POST_PATHS)[number];

type ReadOnlyPost = (path: ReadOnlyPostPath, body: unknown) => Promise<Result<unknown, GraphError>>;

const isReadOnlyPost = (path: string): path is ReadOnlyPostPath => READ_ONLY_POST_PATHS.some((readOnly) => readOnly === path);

const writeRefused = (path: string): Result<never, GraphError> =>
  err({
    type: 'validation_error',
    code: 'write_refused',
    message: `POST ${path} was not sent: this graph only reads, and a read command may POST only to ${READ_ONLY_POST_PATHS.join(' and ')}.`,
  });

// The run-time half of the POST check: a path outside the two is never sent.
const postIfReadOnly = async (post: ReadOnlyPost, path: string, body: unknown): Promise<Result<unknown, GraphError>> =>
  isReadOnlyPost(path) ? post(path, body) : writeRefused(path);

export { postIfReadOnly, READ_ONLY_POST_PATHS };
export type { ReadOnlyPost, ReadOnlyPostPath };
