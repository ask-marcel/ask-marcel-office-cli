import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { AuthError, AuthManager } from '../infra/auth.ts';
import type { GraphClient, GraphError } from '../infra/graph-client.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
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
  getCachedTokenInfo: async () => ({ ok: true, value: { scopes: [], audience: undefined, expiresAt: undefined } }),
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
  getCachedTokenInfo: async () => ({ ok: false, error }),
});

describe('buildCli command surface', () => {
  it('renders an authenticated envelope when login succeeds', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'login']));
    expect(out).toContain('"status":"authenticated"');
  });

  it('renders an Authentication cancelled error when the user closes the browser', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: cancelledAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'login']));
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
      fs: createFileSystemFake(),
      onCommandError: () => {
        errorReports += 1;
      },
    });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'login']));
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
      fs: createFileSystemFake(),
      onCommandError: () => {
        errorReports += 1;
      },
    });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'login']));
    expect(errorReports).toBe(0);
  });

  it('renders the underlying message when login fails for a non-cancellation reason', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: failedAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'login']));
    expect(out).toContain('browser launch failed');
  });

  it('renders a logged_out envelope when logout succeeds', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'logout']));
    expect(out).toContain('"status":"logged_out"');
  });

  it('renders the underlying message when logout fails', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: failedAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'logout']));
    expect(out).toContain('rm denied');
  });

  it('runs a generic Graph command and renders the value as JSON', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({ displayName: 'Vincent' }), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
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
      getCachedTokenInfo: async () => ({ ok: true, value: { scopes: [], audience: undefined, expiresAt: undefined } }),
    };
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: captureGraph, logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
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
      getCachedTokenInfo: async () => ({ ok: true, value: { scopes: [], audience: undefined, expiresAt: undefined } }),
    };
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: captureGraph, logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
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
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'get-current-user']));
    expect(out).toContain('not found');
  });

  it('runs npm install when the user invokes `update` and the manager is npm', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'npm', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'update']));
    expect(runner.calls[0]).toEqual({ command: 'npm', args: ['i', '-g', 'ask-marcel-office-cli@latest'] });
    expect(out).toContain('"status":"updated"');
    expect(out).toContain('"via":"npm"');
  });

  it('runs `bun add -g` when the user invokes `update` and the manager is bun', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'bun', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'update']));
    expect(runner.calls[0]).toEqual({ command: 'bun', args: ['add', '-g', 'ask-marcel-office-cli@latest'] });
    expect(out).toContain('"via":"bun"');
  });

  it('renders the install exit code when the update install exits non-zero', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake({ resultPerCall: [{ exitCode: 7 }] });
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'npm', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'update']));
    expect(out).toContain('exited with code 7');
  });

  it('renders the spawn-failed message when the update install cannot be spawned', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake({ throwOn: [0] });
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'npm', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'update']));
    expect(out).toContain('update failed');
  });

  it('auto-detects the package manager from the bin path when packageManager is not supplied', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake();
    const previousArgv = process.argv[1];
    process.argv[1] = '/Users/anyone/.bun/install/global/node_modules/ask-marcel-office-cli/dist/cli.js';
    try {
      const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, fs: createFileSystemFake() });
      await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'update']));
      expect(runner.calls[0]).toEqual({ command: 'bun', args: ['add', '-g', 'ask-marcel-office-cli@latest'] });
    } finally {
      process.argv[1] = previousArgv;
    }
  });

  it('prints Markdown for a single command when the user runs `docs <cmd>`', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'docs', 'get-current-user']));
    expect(out).toContain('# `get-current-user`');
    expect(out).toContain('## Example');
  });

  it('renders an unknown-command message when the user runs `docs` with a name that does not exist', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'docs', 'this-is-not-a-real-command']));
    expect(out).toContain('Unknown command');
    expect(out).toContain('this-is-not-a-real-command');
  });

  it('prints rich Markdown docs when `docs <lifecycle-command>` is invoked, since lifecycle commands now have manifest entries', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'docs', 'login']));
    expect(out).toContain('# `login`');
    expect(out).toContain('**Category:** Lifecycle');
    expect(out).toContain('Authenticate against Microsoft Graph');
  });

  it('falls back to the unknown-command error when `docs help` is invoked (commander does not register `help` as a regular subcommand)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'docs', 'help']));
    expect(out).toContain('Unknown command');
    expect(out).toContain('help');
  });

  it("'help-json' subcommand prints the full machine-readable manifest (same shape as docs/commands.json) to stdout", async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), version: '1.0.0', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'help-json']));
    const parsed = JSON.parse(out.trim()) as { package: string; version: string; commands: ReadonlyArray<{ name: string }> };
    expect(parsed.package).toBe('ask-marcel-office-cli');
    expect(parsed.version).toBe('1.0.0');
    expect(parsed.commands.length).toBeGreaterThan(100);
    expect(parsed.commands.some((c) => c.name === 'list-drives')).toBe(true);
  });

  it('routes commander parser errors (unknown option) to the JSON envelope on stdout, not stderr plain text', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const stderrOut = await captureStream('stderr', async () => {
      const stdoutOut = await captureStream('stdout', async () => {
        try {
          await cli.parseAsync(['node', 'ask-marcel', 'list-drives', '--no-such-flag']);
        } catch {
          /* expected — commander throws on parser error after exitOverride */
        }
      });
      const parsed = JSON.parse(stdoutOut.trim()) as { ok: boolean; error: string };
      expect(parsed.ok).toBe(false);
      expect(parsed.error).toContain('unknown option');
      expect(parsed.error).toContain('--no-such-flag');
    });
    expect(stderrOut).toBe('');
  });

  it('routes commander parser errors (unknown subcommand) to the JSON envelope on stdout', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel', 'this-command-does-not-exist']);
      } catch {
        /* expected */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: boolean; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('this-command-does-not-exist');
  });

  it('routes commander parser errors (missing required option) to the JSON envelope on stdout', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel', 'get-mail-message']);
      } catch {
        /* expected */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: boolean; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error.toLowerCase()).toContain('required');
  });

  it('global --output-path writes inline base64 bytes to disk and replaces base64 with savedTo in the envelope (lets the LLM avoid round-tripping multi-MB PDFs through stdout)', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: inlinedPdf, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel', '--output-path', '/work/test-output/may-deck.pdf', 'download-drive-item-as-pdf', '--drive-id', 'd1', '--item-id', 'i1'])
    );
    const parsed = JSON.parse(out.trim()) as { ok: true; data: { contentType: string; size: number; savedTo: string; base64?: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.savedTo).toBe('/work/test-output/may-deck.pdf');
    expect(parsed.data.base64).toBeUndefined();
    const written = fs.snapshotBytes('/work/test-output/may-deck.pdf');
    expect(written).toBeDefined();
    if (written) expect(Array.from(written)).toEqual([0x25, 0x50, 0x44, 0x46, 0x2d]);
  });

  it('global --output-path writes a text body via writeText for markdown/plain-text returning commands', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const textGraph: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'notes.md' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'text/plain', size: 5, text: 'hello' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: textGraph, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel', '--output-path', '/work/test-output/notes.md', 'download-drive-item-as-markdown', '--drive-id', 'd1', '--item-id', 'i1'])
    );
    const parsed = JSON.parse(out.trim()) as { ok: true; data: { savedTo: string; text?: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.savedTo).toBe('/work/test-output/notes.md');
    expect(parsed.data.text).toBeUndefined();
    expect(fs.snapshot('/work/test-output/notes.md')).toBe('hello');
  });

  it('global --output-path is a no-op when the command returns plain JSON (no base64 / no text) — surfaces a clear error rather than silently no-op-ing', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({ displayName: 'Vincent' }), logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel', '--output-path', '/work/test-output/profile.json', 'get-current-user']);
      } catch {
        /* commander may throw after exitOverride for explicit failures */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('--output-path');
    expect(parsed.error).toContain('did not return inlined bytes');
    expect(fs.has('/work/test-output/profile.json')).toBe(false);
  });

  it('global --output-path surfaces a write_failed error envelope when the filesystem rejects the write (e.g. permission denied)', async () => {
    const logger = createLoggerFake();
    const fs: FileSystem = {
      readJson: async () => ({ ok: false, error: { type: 'not_found' as const } }),
      writeText: async () => ({ ok: true, value: undefined }),
      writeBytes: async () => ({ ok: false, error: { type: 'io_failed' as const, message: 'EACCES: permission denied, open' } }),
      deleteIfExists: async () => ({ ok: true, value: undefined }),
    };
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: inlinedPdf, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel', '--output-path', '/root/forbidden.pdf', 'download-drive-item-as-pdf', '--drive-id', 'd1', '--item-id', 'i1'])
    );
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('--output-path: write failed');
    expect(parsed.error).toContain('EACCES');
  });

  it('`help <unknown>` returns a JSON-envelope error rather than silently exiting (the audit-flagged v1.0.0 §1.2 inconsistency vs `<unknown>` and `docs <unknown>`)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'help', 'no-such-command']));
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('no-such-command');
  });

  it('`help <known>` prints the markdown docs (alias of `docs <known>`)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'help', 'list-drives']));
    expect(out).toContain('list-drives');
  });

  it('`help` with no argument prints the global --help text', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'help']));
    expect(out).toContain('Usage: ask-marcel');
  });

  it('omitting --output-path leaves the envelope unchanged (existing consumers still get base64 in JSON)', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: inlinedPdf, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel', 'download-drive-item-as-pdf', '--drive-id', 'd1', '--item-id', 'i1']));
    const parsed = JSON.parse(out.trim()) as { ok: true; data: { base64?: string; savedTo?: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.base64).toBe('JVBERi0=');
    expect(parsed.data.savedTo).toBeUndefined();
  });
});
