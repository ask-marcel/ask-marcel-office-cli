import type { Result } from '../domain/result.ts';
import type { ProcessRunner, ProcessRunnerError, ProcessRunOptions, ProcessRunResult } from '../use-cases/ports/process-runner.ts';

export type ProcessRunnerCall = { readonly command: string; readonly args: ReadonlyArray<string>; readonly options: ProcessRunOptions };

type Answer = Result<ProcessRunResult, ProcessRunnerError>;

export type ProcessRunnerFake = ProcessRunner & { readonly calls: ReadonlyArray<ProcessRunnerCall> };

// A ProcessRunner that starts nothing: it records every call and hands back
// what `answer` gives for it. An answer that is a pending promise holds the
// "child" running, so a test can pile up callers while it is in flight.
export const createProcessRunnerFake = (answer: (call: ProcessRunnerCall) => Answer | Promise<Answer>): ProcessRunnerFake => {
  const calls: ProcessRunnerCall[] = [];
  return {
    calls,
    run: async (command, args, options) => {
      const call = { command, args, options };
      calls.push(call);
      return answer(call);
    },
  };
};
