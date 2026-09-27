import { describe, expect, it } from 'bun:test';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import { buildCli } from './cli.ts';

// The error line an agent reads after a mistyped name or flag.
const errorFor = async (args: ReadonlyArray<string>): Promise<string> => {
  const cli = buildCli({ auth: fakeAuthManager(), graph: fakeGraphClient(), logger: createLoggerFake(), processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
  const original = process.stdout.write.bind(process.stdout);
  let out = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    out += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', ...args]);
  } catch {
    /* commander exits on a parse error */
  } finally {
    process.stdout.write = original;
  }
  return (JSON.parse(out.trim()) as { error: string }).error;
};

describe('the did-you-mean line after a mistyped name or flag', () => {
  it('stays silent rather than offering an unrelated short flag', async () => {
    expect(await errorFor(['list-teams-chat-messages', '--chat-id', 'x', '--top', '5'])).toBe("unknown option '--top'");
    expect(await errorFor(['list-mail-messages', '--id', 'x'])).toBe("unknown option '--id'");
  });

  it('offers a lifecycle command its own flags', async () => {
    expect(await errorFor(['help-json', '--ters'])).toContain('Did you mean `--terse`?');
    expect(await errorFor(['login', '--forse'])).toContain('Did you mean `--force`?');
  });

  it('names the flag a mistyped required one was meant to be, beside the missing-flag message', async () => {
    const error = await errorFor(['list-incomplete-todo-tasks', '--list-id', 'L1']);
    expect(error).toContain("required option '--todo-task-list-id <value>' not specified");
    expect(error).toContain('`--list-id` is not a flag of this command. Did you mean `--todo-task-list-id`?');
  });

  it('reads a flag typed with its value, and a plural or one-letter guess at a command', async () => {
    expect(await errorFor(['list-mail-messages', '--selct=id'])).toContain('Did you mean `--select`?');
    expect(await errorFor(['get-mails'])).toContain('`get-mail-message`');
    expect(await errorFor(['l'])).toBe("unknown command 'l'");
  });

  it('lets a three-letter word stand for a longer one, and a plural for its singular, but not any word one letter off', async () => {
    expect(await errorFor(['list-cal-events'])).toContain('`list-calendar-events`');
    expect(await errorFor(['get-mails-settings'])).toContain('`get-mailbox-settings`');
    expect(await errorFor(['get-maix'])).toBe("unknown command 'get-maix'");
  });

  it('keeps a typo with only 40% of the name left out of the suggestions', async () => {
    expect(await errorFor(['help-json', '--xerbx'])).toBe("unknown option '--xerbx'");
  });

  it('answers a very long mistyped name without scoring it against every command', async () => {
    const started = performance.now();
    expect(await errorFor(['x'.repeat(20_000)])).toBe(`unknown command '${'x'.repeat(20_000)}'`);
    expect(performance.now() - started).toBeLessThan(2000);
  });
});
