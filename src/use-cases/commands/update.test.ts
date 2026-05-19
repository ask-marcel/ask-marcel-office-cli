import { describe, expect, it } from 'bun:test';
import { createProcessRunnerFake } from '../../test-helpers/process-runner-fake.ts';
import { execute } from './update.ts';

describe('update use-case', () => {
  it('invokes npm with the correct global-install args by default and returns the runner exit code in the success envelope', async () => {
    const runner = createProcessRunnerFake();
    const result = await execute(runner, 'npm');
    expect(result.ok).toBe(true);
    expect(runner.calls[0]).toEqual({ command: 'npm', args: ['i', '-g', 'ask-marcel-office-cli@latest'] });
    if (result.ok) expect(result.value.exitCode).toBe(0);
  });

  it('invokes bun with the bun-style add args when bun is the chosen manager', async () => {
    const runner = createProcessRunnerFake();
    const result = await execute(runner, 'bun');
    expect(result.ok).toBe(true);
    expect(runner.calls[0]).toEqual({ command: 'bun', args: ['add', '-g', 'ask-marcel-office-cli@latest'] });
    if (result.ok) expect(result.value.exitCode).toBe(0);
  });

  it('returns install_failed with the runner exit code when the install process exits non-zero', async () => {
    const runner = createProcessRunnerFake({ resultPerCall: [{ exitCode: 1 }] });
    const result = await execute(runner, 'npm');
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('install_failed');
      if (result.error.type === 'install_failed') expect(result.error.exitCode).toBe(1);
    }
  });

  it('returns spawn_failed when the runner cannot spawn the install command', async () => {
    const runner = createProcessRunnerFake({ throwOn: [0] });
    const result = await execute(runner, 'npm');
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'spawn_failed') {
      expect(result.error.message).toContain('fake spawn failed');
    }
  });
});
