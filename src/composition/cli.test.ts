import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { AuthError, AuthManager } from '../infra/auth.ts';
import type { GraphClient, GraphError } from '../infra/graph-client.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import { buildCli } from './cli.ts';

const captureStream = async (stream: 'stdout' | 'stderr', run: () => void | Promise<unknown>): Promise<string> => {
  const target = process[stream];
  const original = target.write.bind(target);
  let captured = '';
  const swap = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  target.write = swap;
  try {
    await run();
  } finally {
    target.write = original;
  }
  return captured;
};

const okAuth = (): AuthManager => ({
  getAccessToken: async () => ({ ok: true, value: accessTokenUnsafe('tok') }),
  getElevatedAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' as const } }),
  logout: async () => ({ ok: true, value: undefined }),
});

const cancelledAuth = (): AuthManager => ({
  getAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' } as AuthError }),
  getElevatedAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' as const } }),
  logout: async () => ({ ok: false, error: { type: 'auth_cancelled' } as AuthError }),
});

const failedAuth = (): AuthManager => ({
  getAccessToken: async () => ({ ok: false, error: { type: 'auth_failed', message: 'browser launch failed' } as AuthError }),
  getElevatedAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' as const } }),
  logout: async () => ({ ok: false, error: { type: 'auth_failed', message: 'rm denied' } as AuthError }),
});

const okGraph = (value: unknown): GraphClient => ({
  get: async () => ({ ok: true, value }),
  post: async () => ({ ok: true, value }),
  getBinary: async () => ({ ok: true, value }),
  getElevated: async () => ({ ok: true, value: {} }),
  getBinaryElevated: async () => ({ ok: true, value: {} }),
  fetchUrl: async () => ({ ok: true, value }),
  put: async () => ({ ok: true, value }),
  delete: async () => ({ ok: true, value }),
});

const errGraph = (error: GraphError): GraphClient => ({
  get: async () => ({ ok: false, error }),
  post: async () => ({ ok: false, error }),
  getBinary: async () => ({ ok: false, error }),
  getElevated: async () => ({ ok: true, value: {} }),
  getBinaryElevated: async () => ({ ok: true, value: {} }),
  fetchUrl: async () => ({ ok: false, error }),
  put: async () => ({ ok: false, error }),
  delete: async () => ({ ok: false, error }),
});

describe('buildCli command surface', () => {
  it('renders an authenticated envelope when login succeeds', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'login']));
    expect(out).toContain('"status":"authenticated"');
  });

  it('renders an Authentication cancelled error when the user closes the browser', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: cancelledAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake() });
    const out = await captureStream('stderr', () => cli.parseAsync(['node', 'ask-marcel', 'login']));
    expect(out).toContain('Authentication cancelled');
  });

  it('invokes onCommandError exactly once when a command fails', async () => {
    const logger = createLoggerFake();
    let errorReports = 0;
    const cli = buildCli({
      auth: cancelledAuth(),
      graph: okGraph({}),
      logger,
      processRunner: createProcessRunnerFake(),
      onCommandError: () => {
        errorReports += 1;
      },
    });
    await captureStream('stderr', () => cli.parseAsync(['node', 'ask-marcel', 'login']));
    expect(errorReports).toBe(1);
  });

  it('does not invoke onCommandError when a command succeeds', async () => {
    const logger = createLoggerFake();
    let errorReports = 0;
    const cli = buildCli({
      auth: okAuth(),
      graph: okGraph({}),
      logger,
      processRunner: createProcessRunnerFake(),
      onCommandError: () => {
        errorReports += 1;
      },
    });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'login']));
    expect(errorReports).toBe(0);
  });

  it('renders the underlying message when login fails for a non-cancellation reason', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: failedAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake() });
    const out = await captureStream('stderr', () => cli.parseAsync(['node', 'ask-marcel', 'login']));
    expect(out).toContain('browser launch failed');
  });

  it('renders a logged_out envelope when logout succeeds', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'logout']));
    expect(out).toContain('"status":"logged_out"');
  });

  it('renders the underlying message when logout fails', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: failedAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake() });
    const out = await captureStream('stderr', () => cli.parseAsync(['node', 'ask-marcel', 'logout']));
    expect(out).toContain('rm denied');
  });

  it('runs a generic Graph command and renders the value as JSON', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({ displayName: 'Vincent' }), logger, processRunner: createProcessRunnerFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'get-current-user']));
    expect(out).toContain('Vincent');
  });

  it('accepts both the canonical and the alias spelling of a command flag (e.g. --task-list-id alongside --todo-task-list-id)', async () => {
    let capturedPath = '';
    const captureGraph: GraphClient = {
      get: async (path: string) => {
        capturedPath = path;
        return { ok: true, value: { value: [] } };
      },
      post: async () => ({ ok: true, value: {} }),
      getBinary: async () => ({ ok: true, value: {} }),
      getElevated: async () => ({ ok: true, value: {} }),
      getBinaryElevated: async () => ({ ok: true, value: {} }),
      fetchUrl: async () => ({ ok: true, value: {} }),
      put: async () => ({ ok: true, value: {} }),
      delete: async () => ({ ok: true, value: {} }),
    };
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: captureGraph, logger, processRunner: createProcessRunnerFake() });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'list-todo-tasks', '--task-list-id', 'AAMkABC']));
    expect(capturedPath).toBe('/me/todo/lists/AAMkABC/tasks');
  });

  it('still accepts the canonical flag name when the user does not use the alias', async () => {
    let capturedPath = '';
    const captureGraph: GraphClient = {
      get: async (path: string) => {
        capturedPath = path;
        return { ok: true, value: { value: [] } };
      },
      post: async () => ({ ok: true, value: {} }),
      getBinary: async () => ({ ok: true, value: {} }),
      getElevated: async () => ({ ok: true, value: {} }),
      getBinaryElevated: async () => ({ ok: true, value: {} }),
      fetchUrl: async () => ({ ok: true, value: {} }),
      put: async () => ({ ok: true, value: {} }),
      delete: async () => ({ ok: true, value: {} }),
    };
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: captureGraph, logger, processRunner: createProcessRunnerFake() });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'list-todo-tasks', '--todo-task-list-id', 'AAMkXYZ']));
    expect(capturedPath).toBe('/me/todo/lists/AAMkXYZ/tasks');
  });

  it('renders the Graph error message when a generic Graph command fails', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      auth: okAuth(),
      graph: errGraph({ type: 'api_error', status: 404, message: 'not found' }),
      logger,
      processRunner: createProcessRunnerFake(),
    });
    const out = await captureStream('stderr', () => cli.parseAsync(['node', 'ask-marcel', 'get-current-user']));
    expect(out).toContain('not found');
  });

  it('runs npm install when the user invokes `update` and the manager is npm', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'npm' });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'update']));
    expect(runner.calls[0]).toEqual({ command: 'npm', args: ['i', '-g', 'ask-marcel-office-cli@latest'] });
    expect(out).toContain('"status":"updated"');
    expect(out).toContain('"via":"npm"');
  });

  it('runs `bun add -g` when the user invokes `update` and the manager is bun', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'bun' });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'update']));
    expect(runner.calls[0]).toEqual({ command: 'bun', args: ['add', '-g', 'ask-marcel-office-cli@latest'] });
    expect(out).toContain('"via":"bun"');
  });

  it('renders the install exit code when the update install exits non-zero', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake({ resultPerCall: [{ exitCode: 7 }] });
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'npm' });
    const out = await captureStream('stderr', () => cli.parseAsync(['node', 'ask-marcel', 'update']));
    expect(out).toContain('exited with code 7');
  });

  it('renders the spawn-failed message when the update install cannot be spawned', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake({ throwOn: [0] });
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'npm' });
    const out = await captureStream('stderr', () => cli.parseAsync(['node', 'ask-marcel', 'update']));
    expect(out).toContain('update failed');
  });

  it('auto-detects the package manager from the bin path when packageManager is not supplied', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake();
    const previousArgv = process.argv[1];
    process.argv[1] = '/Users/anyone/.bun/install/global/node_modules/ask-marcel-office-cli/dist/cli.js';
    try {
      const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner });
      await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'update']));
      expect(runner.calls[0]).toEqual({ command: 'bun', args: ['add', '-g', 'ask-marcel-office-cli@latest'] });
    } finally {
      process.argv[1] = previousArgv;
    }
  });

  it('prints Markdown for a single command when the user runs `docs <cmd>`', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'docs', 'get-current-user']));
    expect(out).toContain('# `get-current-user`');
    expect(out).toContain('## Example');
  });

  it('renders an unknown-command message when the user runs `docs` with a name that does not exist', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake() });
    const out = await captureStream('stderr', () => cli.parseAsync(['node', 'ask-marcel', 'docs', 'this-is-not-a-real-command']));
    expect(out).toContain('Unknown command');
    expect(out).toContain('this-is-not-a-real-command');
  });

  it('points users at --help when `docs <lifecycle-command>` is invoked (login/logout/update/docs/help are not in the registry)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake() });
    const out = await captureStream('stderr', () => cli.parseAsync(['node', 'ask-marcel', 'docs', 'login']));
    expect(out).toContain('lifecycle command');
    expect(out).toContain('ask-marcel login --help');
  });
});
