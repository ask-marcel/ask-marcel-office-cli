import { describe, expect, it } from 'bun:test';
import { describeProcessRunnerContract } from '../test-helpers/process-runner-contract.ts';
import { createNodeProcessRunner } from './process-runner-node.ts';

describeProcessRunnerContract('Node process runner adapter', createNodeProcessRunner);

describe('Node process runner adapter', () => {
  it('returns spawn_failed when args spread throws synchronously inside the Promise executor', async () => {
    const runner = createNodeProcessRunner();
    // Force a TypeError out of `[...args]` by smuggling a non-iterable past
    // the type system. Exercises the executor-level catch branch.
    const result = await runner.run('echo', null as unknown as ReadonlyArray<string>, { stdin: 'ignore', timeoutMs: 10000, maxStdoutBytes: 1024 });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('spawn_failed');
  });
});
