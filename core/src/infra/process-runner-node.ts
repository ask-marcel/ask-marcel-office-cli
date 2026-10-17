/*
 * Node.js process runner — atelier rule-20 quarantine companion.
 *
 * This is the ONLY file under `src/**` that may import `node:child_process`.
 * It exists so the published `dist/cli.js` can spawn child processes when
 * running under plain Node, where `Bun.spawn` is not available.
 *
 * The composition root selects between this and `process-runner-bun.ts` at
 * runtime via `typeof globalThis.Bun`. All other production code consumes
 * the `ProcessRunner` port (`src/use-cases/ports/process-runner.ts`).
 */

import { spawn } from 'node:child_process';
import { err } from '../domain/result.ts';
import { formatError } from '../domain/utilities/format-error.ts';
import type { ProcessRunner } from '../use-cases/ports/process-runner.ts';
import { createChildWatch } from './child-process-watch.ts';

// `close`, not `exit`: it fires once stdout is drained, so nothing printed is
// lost. A spawn that fails (no such command) emits `error`; the first resolve
// wins, so a `close` that follows it changes nothing.
export const createNodeProcessRunner = (): ProcessRunner => ({
  run: async (command, args, options) =>
    new Promise((resolve) => {
      try {
        const child = spawn(command, [...args], { stdio: [options.stdin, 'pipe', 'inherit'] });
        const watch = createChildWatch(options, () => child.kill('SIGKILL'));
        child.stdout?.on('data', (chunk: Uint8Array) => watch.capture(chunk));
        child.on('error', (e) => resolve(watch.fail(formatError(e))));
        child.on('close', (exitCode, signal) => resolve(watch.settle(exitCode, signal)));
      } catch (e) {
        resolve(err({ type: 'spawn_failed', message: formatError(e) }));
      }
    }),
});
