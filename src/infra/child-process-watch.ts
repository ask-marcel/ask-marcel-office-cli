/*
 * Watches one child process for both ProcessRunner adapters, so Bun and Node
 * map an outcome the same way. The deadline and the stdout cap each stop the
 * child, and the first reason to stop is the one reported; a child killed by
 * any other signal is a failure too, never an exit code.
 */
import { err, ok } from '../domain/result.ts';
import type { Result } from '../domain/result.ts';
import type { ProcessRunnerError, ProcessRunOptions, ProcessRunResult } from '../use-cases/ports/process-runner.ts';

type Outcome = Result<ProcessRunResult, ProcessRunnerError>;

export type ChildWatch = {
  readonly capture: (chunk: Uint8Array) => void;
  readonly settle: (exitCode: number | null, signal: string | null) => Outcome;
  readonly fail: (message: string) => Outcome;
};

export const createChildWatch = (options: ProcessRunOptions, kill: () => void): ChildWatch => {
  let stop: 'timed_out' | 'output_too_large' | undefined;
  let bytes = 0;
  let stdout = '';
  const decoder = new TextDecoder();
  const halt = (reason: 'timed_out' | 'output_too_large'): void => {
    stop ??= reason;
    kill();
  };
  // Unref'd so a timer outliving its child never keeps the process alive.
  const timer = setTimeout(() => halt('timed_out'), options.timeoutMs);
  timer.unref();
  const capture = (chunk: Uint8Array): void => {
    if (stop !== undefined) return;
    bytes += chunk.byteLength;
    if (bytes > options.maxStdoutBytes) return halt('output_too_large');
    stdout += decoder.decode(chunk, { stream: true });
  };
  const settle = (exitCode: number | null, signal: string | null): Outcome => {
    clearTimeout(timer);
    if (stop === 'timed_out') return err({ type: 'timed_out', timeoutMs: options.timeoutMs });
    if (stop === 'output_too_large') return err({ type: 'output_too_large', maxStdoutBytes: options.maxStdoutBytes });
    if (signal !== null || exitCode === null) return err({ type: 'killed', signal: signal ?? 'unknown' });
    return ok({ exitCode, stdout: stdout + decoder.decode() });
  };
  const fail = (message: string): Outcome => {
    clearTimeout(timer);
    return err({ type: 'spawn_failed', message });
  };
  return { capture, settle, fail };
};
