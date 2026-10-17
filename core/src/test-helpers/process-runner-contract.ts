import { describe, expect, it } from 'bun:test';
import type { ProcessRunner, ProcessRunOptions } from '../use-cases/ports/process-runner.ts';

// The behaviour every ProcessRunner adapter shares, run by each adapter's own
// test file. The runtimes disagreed once (a child killed by a signal read as
// exit 0 under Node and 143 under Bun); one suite keeps them in step.
const OPTIONS: ProcessRunOptions = { stdin: 'ignore', timeoutMs: 10000, maxStdoutBytes: 1024 * 1024 };

export const describeProcessRunnerContract = (name: string, create: () => ProcessRunner): void => {
  describe(`${name}: the ProcessRunner contract`, () => {
    it('hands back what the child printed and its exit code 0', async () => {
      const result = await create().run('sh', ['-c', 'printf hello'], OPTIONS);
      expect(result).toEqual({ ok: true, value: { exitCode: 0, stdout: 'hello' } });
    });

    it('hands back a non-zero exit code together with what the child printed, so the caller can read an error answer', async () => {
      const result = await create().run('sh', ['-c', 'printf \'{"errorCode":"x"}\'; exit 3'], OPTIONS);
      expect(result).toEqual({ ok: true, value: { exitCode: 3, stdout: '{"errorCode":"x"}' } });
    });

    it('refuses with spawn_failed when the command does not exist', async () => {
      const result = await create().run('definitely-not-a-real-command-zX9', [], OPTIONS);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('spawn_failed');
    });

    it('stops a child that outlives its deadline and reports timed_out', async () => {
      const started = Date.now();
      const result = await create().run('sleep', ['5'], { ...OPTIONS, timeoutMs: 200 });
      expect(result).toEqual({ ok: false, error: { type: 'timed_out', timeoutMs: 200 } });
      expect(Date.now() - started).toBeLessThan(3000);
    });

    it('reports a child killed by a signal as killed, never as a success', async () => {
      const result = await create().run('sh', ['-c', 'kill -TERM $$'], OPTIONS);
      expect(result).toEqual({ ok: false, error: { type: 'killed', signal: 'SIGTERM' } });
    });

    it('stops a child whose output passes the cap and reports output_too_large', async () => {
      const result = await create().run('yes', [], { ...OPTIONS, maxStdoutBytes: 1024 });
      expect(result).toEqual({ ok: false, error: { type: 'output_too_large', maxStdoutBytes: 1024 } });
    });

    it('gives the child an empty input when stdin is ignored, so a reader ends at once', async () => {
      const result = await create().run('cat', [], OPTIONS);
      expect(result).toEqual({ ok: true, value: { exitCode: 0, stdout: '' } });
    });
  });
};
