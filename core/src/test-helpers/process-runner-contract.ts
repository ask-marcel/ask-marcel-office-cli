import { describe, expect, it } from 'bun:test';
import type { ProcessRunner, ProcessRunOptions } from '../use-cases/ports/process-runner.ts';

// The behaviour every ProcessRunner adapter shares, run by each adapter's own
// test file. The runtimes disagreed once (a child killed by a signal read as
// exit 0 under Node and 143 under Bun); one suite keeps them in step.
const OPTIONS: ProcessRunOptions = { stdin: 'ignore', timeoutMs: 10000, maxStdoutBytes: 1024 * 1024 };

// The child is this runtime with a one-line script, so the suite needs no POSIX
// tool on PATH: PowerShell has no `yes`, and its `cat` and `sleep` are aliases.
const RUNTIME = process.execPath;
const script = (source: string): ReadonlyArray<string> => ['-e', source];

// On Windows a spawn call of Bun 1.4 sometimes blocks for 5 to 10 seconds, which
// is more than the default test timeout of 5 seconds.
const SPAWN_TEST_TIMEOUT_MS = 30_000;
const ON_WINDOWS = process.platform === 'win32';

export const describeProcessRunnerContract = (name: string, create: () => ProcessRunner): void => {
  describe(`${name}: the ProcessRunner contract`, () => {
    it(
      'hands back what the child printed and its exit code 0',
      async () => {
        const result = await create().run(RUNTIME, script("process.stdout.write('hello')"), OPTIONS);
        expect(result).toEqual({ ok: true, value: { exitCode: 0, stdout: 'hello' } });
      },
      SPAWN_TEST_TIMEOUT_MS
    );

    it(
      'hands back a non-zero exit code together with what the child printed, so the caller can read an error answer',
      async () => {
        const result = await create().run(RUNTIME, script('process.stdout.write(JSON.stringify({ errorCode: "x" })); process.exitCode = 3'), OPTIONS);
        expect(result).toEqual({ ok: true, value: { exitCode: 3, stdout: '{"errorCode":"x"}' } });
      },
      SPAWN_TEST_TIMEOUT_MS
    );

    it(
      'refuses with spawn_failed when the command does not exist',
      async () => {
        const result = await create().run('definitely-not-a-real-command-zX9', [], OPTIONS);
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.error.type).toBe('spawn_failed');
      },
      SPAWN_TEST_TIMEOUT_MS
    );

    it(
      'stops a child that outlives its deadline and reports timed_out',
      async () => {
        const started = Date.now();
        const result = await create().run(RUNTIME, script('setTimeout(() => {}, 5000)'), { ...OPTIONS, timeoutMs: 200 });
        expect(result).toEqual({ ok: false, error: { type: 'timed_out', timeoutMs: 200 } });
        // The blocked spawn call of Windows counts in the elapsed time, so the bound holds off Windows only.
        if (!ON_WINDOWS) expect(Date.now() - started).toBeLessThan(3000);
      },
      SPAWN_TEST_TIMEOUT_MS
    );

    // Windows has no signals: a shell that kills itself there ends with exit code 0.
    it.skipIf(ON_WINDOWS)('reports a child killed by a signal as killed, never as a success', async () => {
      const result = await create().run('sh', ['-c', 'kill -TERM $$'], OPTIONS);
      expect(result).toEqual({ ok: false, error: { type: 'killed', signal: 'SIGTERM' } });
    });

    it(
      'stops a child whose output passes the cap and reports output_too_large',
      async () => {
        const result = await create().run(RUNTIME, script("process.stdout.write('y'.repeat(100000)); setTimeout(() => {}, 5000)"), { ...OPTIONS, maxStdoutBytes: 1024 });
        expect(result).toEqual({ ok: false, error: { type: 'output_too_large', maxStdoutBytes: 1024 } });
      },
      SPAWN_TEST_TIMEOUT_MS
    );

    it(
      'gives the child an empty input when stdin is ignored, so a reader ends at once',
      async () => {
        const result = await create().run(RUNTIME, script('process.stdin.on("data", (chunk) => process.stdout.write(chunk))'), OPTIONS);
        expect(result).toEqual({ ok: true, value: { exitCode: 0, stdout: '' } });
      },
      SPAWN_TEST_TIMEOUT_MS
    );
  });
};
