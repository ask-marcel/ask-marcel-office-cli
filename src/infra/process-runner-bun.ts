import { err } from '../domain/result.ts';
import { formatError } from '../domain/utilities/format-error.ts';
import type { ProcessRunner } from '../use-cases/ports/process-runner.ts';
import { createChildWatch } from './child-process-watch.ts';

const drain = async (stream: ReadableStream<Uint8Array>, onChunk: (chunk: Uint8Array) => void): Promise<void> => {
  const reader = stream.getReader();
  for (let next = await reader.read(); !next.done; next = await reader.read()) onChunk(next.value);
};

export const createBunProcessRunner = (): ProcessRunner => ({
  run: async (command, args, options) => {
    try {
      const proc = Bun.spawn([command, ...args], { stdin: options.stdin, stdout: 'pipe', stderr: 'inherit' });
      const watch = createChildWatch(options, () => proc.kill('SIGKILL'));
      await drain(proc.stdout, watch.capture);
      await proc.exited;
      return watch.settle(proc.exitCode, proc.signalCode);
    } catch (e) {
      return err({ type: 'spawn_failed', message: formatError(e) });
    }
  },
});
