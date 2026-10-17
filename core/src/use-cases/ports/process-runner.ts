import type { Result } from '../../domain/result.ts';

// A child that could not start, outlived its deadline, was killed by a signal,
// or printed more than the caller allows. Each is a failure: a child that
// exited on its own returns its exit code and stdout, non-zero included, so the
// caller can read an error answer the child printed.
export type ProcessRunnerError =
  | { readonly type: 'spawn_failed'; readonly message: string }
  | { readonly type: 'timed_out'; readonly timeoutMs: number }
  | { readonly type: 'killed'; readonly signal: string }
  | { readonly type: 'output_too_large'; readonly maxStdoutBytes: number };

export type ProcessRunResult = { readonly exitCode: number; readonly stdout: string };

// stdout is captured up to maxStdoutBytes and stderr is inherited, so a child
// can show progress in the user's terminal. stdin is inherited only when the
// caller asks: in an MCP server it is the JSON-RPC channel.
export type ProcessRunOptions = {
  readonly stdin: 'inherit' | 'ignore';
  readonly timeoutMs: number;
  readonly maxStdoutBytes: number;
};

export type ProcessRunner = {
  readonly run: (command: string, args: ReadonlyArray<string>, options: ProcessRunOptions) => Promise<Result<ProcessRunResult, ProcessRunnerError>>;
};
