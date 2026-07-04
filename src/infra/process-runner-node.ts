/*
 * Node.js process runner — atelier rule-20 quarantine companion.
 *
 * This is the ONLY file under `src/**` that may import `node:child_process`.
 * It exists so the published `dist/cli.js` can spawn child processes (used
 * by the `ask-marcel-office update` command to invoke `npm` or `bun`) when running
 * under plain Node, where `Bun.spawn` is not available.
 *
 * The composition root selects between this and `process-runner-bun.ts` at
 * runtime via `typeof globalThis.Bun`. All other production code consumes
 * the `ProcessRunner` port (`src/use-cases/ports/process-runner.ts`).
 */

import { spawn } from 'node:child_process';
import { err, ok } from '../domain/result.ts';
import { formatError } from '../domain/utilities/format-error.ts';
import type { ProcessRunner } from '../use-cases/ports/process-runner.ts';

export const createNodeProcessRunner = (): ProcessRunner => ({
  runInherit: async (command, args) =>
    new Promise((resolve) => {
      try {
        const child = spawn(command, [...args], { stdio: 'inherit' });
        child.on('exit', (exitCode) => resolve(ok({ exitCode: exitCode ?? 0 })));
        child.on('error', (e) => resolve(err({ type: 'spawn_failed', message: formatError(e) })));
      } catch (e) {
        resolve(err({ type: 'spawn_failed', message: formatError(e) }));
      }
    }),
});
