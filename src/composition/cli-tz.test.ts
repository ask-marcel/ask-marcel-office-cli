import { afterEach, describe, expect, it } from 'bun:test';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import { currentDateZone, setDateZone } from '../use-cases/commands/date-zone.ts';
import { buildCli } from './cli.ts';
import { machineTimeZone } from './date-zone.ts';

const captureStdout = async (run: () => Promise<unknown>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run();
  } catch {
    /* commander may exit */
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

const cliWith = (): { cli: ReturnType<typeof buildCli>; logger: ReturnType<typeof createLoggerFake> } => {
  const logger = createLoggerFake();
  const cli = buildCli({ auth: fakeAuthManager(), graph: fakeGraphClient(), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
  return { cli, logger };
};

const savedEnv = process.env['ASKMARCEL_TZ'];

describe('the --tz option and the run zone', () => {
  afterEach(() => {
    setDateZone('UTC');
    if (savedEnv === undefined) delete process.env['ASKMARCEL_TZ'];
    else process.env['ASKMARCEL_TZ'] = savedEnv;
  });

  it('sets the zone from --tz before a subcommand runs', async () => {
    const { cli } = cliWith();
    await captureStdout(() => cli.parseAsync(['node', 'ask-marcel-office', '--tz', 'Asia/Shanghai', 'docs', 'list-mail-messages']));
    expect(currentDateZone()).toBe('Asia/Shanghai');
  });

  it('falls back to ASKMARCEL_TZ, then to the machine, and warns once about a bad environment value', async () => {
    process.env['ASKMARCEL_TZ'] = 'Europe/Amsterdam';
    const env = cliWith();
    await captureStdout(() => env.cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'list-mail-messages']));
    expect(currentDateZone()).toBe('Europe/Amsterdam');
    delete process.env['ASKMARCEL_TZ'];
    const machine = cliWith();
    await captureStdout(() => machine.cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'list-mail-messages']));
    expect(currentDateZone()).toBe(machineTimeZone());
    process.env['ASKMARCEL_TZ'] = 'Mars/Olympus';
    const bad = cliWith();
    await captureStdout(() => bad.cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'list-mail-messages']));
    expect(currentDateZone()).toBe(machineTimeZone());
    expect(bad.logger.calls.some((e) => e.level === 'warn' && e.event === 'date_zone_fallback')).toBe(true);
  });

  it('refuses a --tz that is not a known zone, naming the rule', async () => {
    const { cli } = cliWith();
    const out = await captureStdout(() => cli.parseAsync(['node', 'ask-marcel-office', '--tz', 'Mars/Olympus', 'docs', 'list-mail-messages']));
    expect(out).toContain('Not a known IANA time zone');
    expect(currentDateZone()).toBe('UTC');
  });
});
