import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { AuthError, AuthManager } from '../infra/auth.ts';
import type { GraphClient, GraphError, TokenInfo } from '../infra/graph-client.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { buildMediaSamples } from '../test-helpers/office-fixtures.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { commands } from '../use-cases/commands/index.ts';
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

const okAuth = (): AuthManager => fakeAuthManager();

const cancelledAuth = (): AuthManager =>
  fakeAuthManager({
    getAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' } as AuthError }),
    logout: async () => ({ ok: false, error: { type: 'auth_cancelled' } as AuthError }),
  });

const failedAuth = (): AuthManager =>
  fakeAuthManager({
    getAccessToken: async () => ({ ok: false, error: { type: 'auth_failed', message: 'browser launch failed' } as AuthError }),
    logout: async () => ({ ok: false, error: { type: 'auth_failed', message: 'rm denied' } as AuthError }),
  });

const okGraph = (value: unknown): GraphClient =>
  fakeGraphClient({
    get: async () => ({ ok: true, value }),
    post: async () => ({ ok: true, value }),
    getBinary: async () => ({ ok: true, value }),
    fetchUrl: async () => ({ ok: true, value }),
    put: async () => ({ ok: true, value }),
    delete: async () => ({ ok: true, value }),
  });

const errGraph = (error: GraphError): GraphClient =>
  fakeGraphClient({
    get: async () => ({ ok: false, error }),
    post: async () => ({ ok: false, error }),
    getBinary: async () => ({ ok: false, error }),
    fetchUrl: async () => ({ ok: false, error }),
    put: async () => ({ ok: false, error }),
    delete: async () => ({ ok: false, error }),
    getCachedTokenInfo: async () => ({ ok: false, error }),
  });

const sampleTokenInfo = (over: Partial<TokenInfo> = {}): TokenInfo => ({
  scopes: ['Mail.Read'],
  audience: 'https://graph.microsoft.com',
  expiresAt: '2026-12-31T00:00:00.000Z',
  expiresInSeconds: 8938,
  elevated: { available: false, expiresInSeconds: undefined, scopes: [], refresh: 'interactive' },
  chatsvcagg: { available: true, expiresInSeconds: 5400, scopes: ['user_impersonation'], refresh: 'automatic' },
  ic3: { available: true, expiresInSeconds: 5400, scopes: ['Teams.AccessAsUser.All'], refresh: 'automatic' },
  ...over,
});

const graphWithTokenInfo = (info: TokenInfo): GraphClient => fakeGraphClient({ getCachedTokenInfo: async () => ({ ok: true, value: info }) });

describe('buildCli command surface', () => {
  it('renders the slim availability summary + two-pointer hint when login succeeds (under --output json)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: graphWithTokenInfo(sampleTokenInfo()), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'login']));
    const parsed = JSON.parse(out) as { data: { status: string; available: string[]; hint: string } };
    expect(parsed.data.status).toBe('authenticated');
    expect(parsed.data.available).toEqual(['basic', 'chatsvcagg', 'ic3']); // sampleTokenInfo has elevated unavailable → omitted
    expect(parsed.data.hint).toContain('scopes-check');
    expect(parsed.data.hint).toContain('login --force');
  });

  it('top-level description derives endpoint counts from the registry and labels mail-draft writes as mutations, never as searches (F-03)', () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const getCount = Object.values(commands).filter((c) => c.meta.graphMethod === 'GET').length;
    const mutating = Object.entries(commands)
      .filter(([, c]) => c.meta.mutates === true)
      .map(([n]) => n);
    const searchPosts = Object.entries(commands)
      .filter(([, c]) => c.meta.graphMethod === 'POST' && c.meta.mutates !== true)
      .map(([n]) => n);
    const description = cli.description();
    expect(description).toContain(`${getCount} GET endpoints`);
    // every search POST is listed in the "search POST" clause
    for (const name of searchPosts) expect(description).toContain(name);
    // every write command is named as a mutation, and there IS at least one
    expect(mutating.length).toBeGreaterThan(0);
    for (const name of mutating) expect(description).toContain(name);
    // the old bug: every POST was labelled "searches, not mutations", which
    // mislabelled create-mail-draft. That phrasing must be gone.
    expect(description).not.toContain('searches, not mutations');
  });

  it('renders the slim summary in text format with the two-pointer hint', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: graphWithTokenInfo(sampleTokenInfo()), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
    expect(out).toContain('status: authenticated');
    expect(out).toContain('scopes-check');
    expect(out).toContain('login --force');
  });

  it('includes a token in login available only when getCachedTokenInfo reports it available (elevated)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      auth: okAuth(),
      graph: graphWithTokenInfo(sampleTokenInfo({ elevated: { available: true, expiresInSeconds: 1800, scopes: ['Chat.ReadBasic'], refresh: 'interactive' } })),
      logger,
      processRunner: createProcessRunnerFake(),
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'login']));
    const parsed = JSON.parse(out) as { data: { available: string[] } };
    expect(parsed.data.available).toEqual(['basic', 'elevated', 'chatsvcagg', 'ic3']); // elevated now available → listed
  });

  it('login --force forwards the force flag to getAccessToken so a warm session re-captures all tokens', async () => {
    const calls: Array<{ force?: boolean } | undefined> = [];
    const capturingAuth: AuthManager = {
      ...okAuth(),
      getAccessToken: async (options?: { force?: boolean }) => {
        calls.push(options);
        return { ok: true as const, value: accessTokenUnsafe('tok') };
      },
    };
    const logger = createLoggerFake();
    const cli = buildCli({ auth: capturingAuth, graph: graphWithTokenInfo(sampleTokenInfo()), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'login', '--force']));
    expect(calls.some((c) => c?.force === true)).toBe(true);
  });

  it('surfaces the token-status read failure when getCachedTokenInfo fails after a successful login', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      auth: okAuth(),
      graph: errGraph({ type: 'auth_failed', message: 'no token cached' }),
      logger,
      processRunner: createProcessRunnerFake(),
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
    expect(out).toContain('no token cached');
  });

  it('renders a Graph error in text format with `error:` + `source: graph` (envelope-symmetry fix — round 2 — stamps the source even when the hint table did not match)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      auth: okAuth(),
      graph: errGraph({ type: 'api_error', status: 404, message: 'not found' }),
      logger,
      processRunner: createProcessRunnerFake(),
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'get-current-user']));
    // The api_error → 'graph' mapping in `sourceFromGraphError` makes `source`
    // present on every Graph-originating failure, hint-matched or not. The
    // 404 message "not found" has no rule match → no `hint:` line → but the
    // `source: graph` line IS stamped from the explicit fallback.
    expect(out).toBe('error: not found\nsource: graph\n');
  });

  it('threads the Retry-After interval from a throttled Graph 429 into the JSON error envelope so a crawler can honor the backoff', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      auth: okAuth(),
      graph: errGraph({ type: 'api_error', status: 429, message: 'TooManyRequests: Too many requests', code: 'TooManyRequests', retryAfterSeconds: 120 }),
      logger,
      processRunner: createProcessRunnerFake(),
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'get-current-user']));
    const parsed = JSON.parse(out.trim()) as { ok: false; retryAfterSeconds?: number };
    expect(parsed.ok).toBe(false);
    expect(parsed.retryAfterSeconds).toBe(120);
  });

  it('renders an Authentication cancelled error when the user closes the browser', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: cancelledAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
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
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
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
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
    expect(errorReports).toBe(0);
  });

  it('renders the underlying message when login fails for a non-cancellation reason', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: failedAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
    expect(out).toContain('browser launch failed');
  });

  it('renders a logged_out envelope when logout succeeds (under --output json)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'logout']));
    expect(out).toContain('"status":"logged_out"');
  });

  it('renders the underlying message when logout fails', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: failedAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'logout']));
    expect(out).toContain('rm denied');
  });

  it('runs a generic Graph command and renders the value as JSON', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({ displayName: 'Jordan' }), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'get-current-user']));
    expect(out).toContain('Jordan');
  });

  it('accepts both the canonical and the alias spelling of a command flag (e.g. --task-list-id alongside --todo-task-list-id)', async () => {
    let capturedPath = '';
    const captureGraph: GraphClient = fakeGraphClient({
      get: async (path: string) => {
        capturedPath = path;
        return { ok: true, value: { value: [] } };
      },
    });
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: captureGraph, logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'list-todo-tasks', '--task-list-id', 'AAMkABC']));
    expect(capturedPath).toBe('/me/todo/lists/AAMkABC/tasks');
  });

  it('convert-local-file reads the file through the CLI-wired filesystem and never touches Graph', async () => {
    let graphCalled = false;
    const tripwireGraph: GraphClient = fakeGraphClient({
      get: async () => {
        graphCalled = true;
        return { ok: true, value: {} };
      },
    });
    const fs = createFileSystemFake();
    fs.seed('/work/data.csv', 'name,age\nAlice,30');
    const cli = buildCli({ auth: okAuth(), graph: tripwireGraph, logger: createLoggerFake(), processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'convert-local-file', '--path', '/work/data.csv']));
    const parsed = JSON.parse(out.trim()) as { ok: boolean; data: { text: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.text).toContain('| Alice | 30 |');
    expect(graphCalled).toBe(false);
  });

  it('stamps a failed local-only command with source `cli`, not `graph` — it never touched Graph (F-02)', async () => {
    const fs = createFileSystemFake(); // nothing seeded → readBytes returns a not-found error
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger: createLoggerFake(), processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'convert-local-file', '--path', '/work/missing.csv']));
    const parsed = JSON.parse(out.trim()) as { ok: boolean; source?: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.source).toBe('cli');
  });

  it('accepts --id as an alias for --message-id on sole-message-id commands (P3)', async () => {
    let capturedPath = '';
    const captureGraph: GraphClient = fakeGraphClient({
      get: async (path: string) => {
        capturedPath = path;
        return { ok: true, value: { id: 'AAMkAGI2', subject: 'hi' } };
      },
    });
    const cli = buildCli({ auth: okAuth(), graph: captureGraph, logger: createLoggerFake(), processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'get-mail-message', '--id', 'AAMkAGI2']));
    // --id maps to messageId → the message path (a default $select is appended by the command).
    expect(capturedPath.startsWith('/me/messages/AAMkAGI2')).toBe(true);
  });

  it('accepts --id as an alias for --message-id on update-mail-draft, matching its read siblings (F-05)', async () => {
    let capturedPath = '';
    const captureGraph: GraphClient = fakeGraphClient({
      patch: async (path: string) => {
        capturedPath = path;
        return { ok: true, value: { id: 'AAMkAGI2', subject: 'Updated' } };
      },
    });
    const cli = buildCli({ auth: okAuth(), graph: captureGraph, logger: createLoggerFake(), processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'update-mail-draft', '--id', 'AAMkAGI2', '--subject', 'Updated']));
    expect(capturedPath).toBe('/me/messages/AAMkAGI2');
  });

  it('accepts --start/--end as aliases for --start-date-time/--end-date-time on the calendar-view family', async () => {
    const pathFor = async (args: ReadonlyArray<string>): Promise<string> => {
      let captured = '';
      const graph: GraphClient = fakeGraphClient({
        get: async (path: string) => {
          captured = path;
          return { ok: true, value: { value: [] } };
        },
      });
      const cli = buildCli({ auth: okAuth(), graph, logger: createLoggerFake(), processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
      await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'list-calendar-view', ...args]));
      return captured;
    };
    // The short --start/--end forms agents reach for must produce the identical Graph
    // call as the canonical --start-date-time/--end-date-time (alias → canonical key).
    const aliased = await pathFor(['--start', '2026-04-01T00:00:00Z', '--end', '2026-05-01T00:00:00Z']);
    const canonical = await pathFor(['--start-date-time', '2026-04-01T00:00:00Z', '--end-date-time', '2026-05-01T00:00:00Z']);
    expect(aliased).toBe(canonical);
    expect(aliased).toContain('/me/calendarView?startDateTime=');
  });

  it('rewrites a validation-error message to reference the alias the user typed', async () => {
    const captureGraph: GraphClient = fakeGraphClient({
      get: async () => ({ ok: true, value: { value: [] } }),
    });
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: captureGraph, logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', 'search-onenote-pages', '--query', '']);
      } catch {
        /* commander may exit on validation failure */
      }
    });
    expect(out).toContain('--query is empty');
    expect(out).not.toContain('--title-substring is empty');
  });

  it('help-json rejects an explicit --output text (manifest is JSON by contract — audit )', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'text', 'help-json']);
      } catch {
        /* commander may exit */
      }
    });
    expect(out).toContain('help-json always emits JSON');
    expect(out).not.toContain('"commands"');
  });

  it('help-json still works when --output is left at its default text value (no explicit flag)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', 'help-json']);
      } catch {
        /* commander may exit */
      }
    });
    expect(out).toContain('"commands"');
  });

  // — compact default `--help` + help-json projections
  // (--terse / --category). These tests fix the discoverability gap an LLM
  // hits today: `--help` returns ~60 KB so the model is forced to truncate
  // or dump-to-disk before it can read the listing; `help-json` returns
  // 370 KB unfiltered. v1.4.0 surface-consolidation: the `--verbose` opt-out
  // was dropped — it was a one-trick toggle on this top-level listing only,
  // and `help-json --terse` covers the same need with a structured payload.
  //
  // Notes on test mechanics:
  // - Triggering Commander's `--help` flag directly would call `process.exit(0)`
  // even with `exitOverride` (the override returns, Commander then calls
  // `process.exit(exitCode)`), killing the test process. The custom `help`
  // subcommand goes through `program.outputHelp()` which renders without
  // exiting, so we use it for assertion. Production behaviour for `--help`
  // is exercised by the bare-args test at the bottom of this file.
  it('compact default help listing truncates each subcommand description to its first sentence', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help']));
    // get-mail-message has a multi-sentence summary post-§A; the compact
    // help should carry only the first sentence. The post-first-sentence
    // "ships a slim default" prose is the marker that we successfully cut.
    expect(out).toContain('Get a single Outlook message by ID.');
    expect(out).not.toContain('ships a slim default');
    // Footer must point at the discovery surfaces.
    expect(out).toContain('help-json [--terse] [--category mail]');
  });

  it('compact help listing stays under the 20 KB token-budget ceiling (byte-count regression guard — replaces the dropped --verbose opt-out)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const compact = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help']));
    // First-sentence truncation keeps the listing well under 20 KB. The pre-
    // compaction full-summary form ran ~60 KB; if compaction silently
    // regresses (e.g. compactSummary cuts wrong), this guard fires.
    // 35 KB ceiling: current compact listing is ~28 KB; this guard fires if
    // a future change accidentally restores the pre-compaction full-summary
    // form (~60 KB). It leaves ~7 KB of headroom for new commands.
    expect(compact.length).toBeLessThan(35 * 1024);
  });

  it('--verbose is no longer a recognised top-level option (v1.4.0 surface-consolidation drop)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    let unknown = false;
    try {
      await cli.parseAsync(['node', 'ask-marcel-office', '--verbose', 'help']);
    } catch (e) {
      // Commander throws a CommanderError with code 'commander.unknownOption'
      // when it doesn't recognise a flag. exitOverride routes that through.
      unknown = (e as { code?: string }).code === 'commander.unknownOption';
    }
    expect(unknown).toBe(true);
  });

  it('help-json --terse strips per-command heavy fields (no options/example/graphPathTemplate/responseShape)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help-json', '--terse']);
      } catch {
        /* commander exits */
      }
    });
    expect(out).toContain('"commands"');
    expect(out).not.toContain('"graphPathTemplate"');
    expect(out).not.toContain('"responseShape"');
    expect(out).not.toContain('"options"');
  });

  it('help-json --category mail filters the manifest down to the mail category', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help-json', '--category', 'mail']);
      } catch {
        /* commander exits */
      }
    });
    expect(out).toContain('"get-mail-message"');
    // get-current-user is in the `user` category, must not leak in.
    expect(out).not.toContain('"get-current-user"');
  });

  it('help-json --terse --category mail composes: terse projection within a single category', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help-json', '--terse', '--category', 'mail']);
      } catch {
        /* commander exits */
      }
    });
    expect(out).toContain('"get-mail-message"');
    expect(out).not.toContain('"options"');
    expect(out).not.toContain('"graphPathTemplate"');
  });

  it('help-json --category with an unknown name surfaces a structured `ok:false` envelope listing the valid categories', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help-json', '--category', 'notarealcategory']);
      } catch {
        /* commander exits */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('Unknown --category "notarealcategory"');
    expect(parsed.error).toContain('mail');
    expect(parsed.error).toContain('drive');
  });

  it('rejects a duplicate --output flag', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({ id: 'u1' }), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output', 'text', 'get-current-user']);
      } catch {
        /* commander may exit */
      }
    });
    expect(out).toContain('--output cannot be passed more than once');
  });

  it('rejects a single-value flag passed more than once', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({ value: [] }), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', 'list-folder-files', '--drive-id', 'd1', '--item-id', 'i1', '--filter', 'A', '--filter', 'B']);
      } catch {
        /* commander may exit on validation failure */
      }
    });
    expect(out).toContain('--filter cannot be passed more than once');
    expect(out).toContain('previous: "A"');
    expect(out).toContain('new: "B"');
  });

  it('still accepts the canonical flag name when the user does not use the alias', async () => {
    let capturedPath = '';
    const captureGraph: GraphClient = fakeGraphClient({
      get: async (path: string) => {
        capturedPath = path;
        return { ok: true, value: { value: [] } };
      },
    });
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: captureGraph, logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'list-todo-tasks', '--todo-task-list-id', 'AAMkXYZ']));
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
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'get-current-user']));
    expect(out).toContain('not found');
  });

  it('runs npm install when the user invokes `update` and the manager is npm (under --output json)', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'npm', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'update']));
    expect(runner.calls[0]).toEqual({ command: 'npm', args: ['i', '-g', 'ask-marcel-office-cli@latest'] });
    expect(out).toContain('"status":"updated"');
    expect(out).toContain('"via":"npm"');
  });

  it('runs `bun add -g` when the user invokes `update` and the manager is bun (under --output json)', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'bun', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'update']));
    expect(runner.calls[0]).toEqual({ command: 'bun', args: ['add', '-g', 'ask-marcel-office-cli@latest'] });
    expect(out).toContain('"via":"bun"');
  });

  it('renders the install exit code when the update install exits non-zero', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake({ resultPerCall: [{ exitCode: 7 }] });
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'npm', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'update']));
    expect(out).toContain('exited with code 7');
  });

  it('renders the spawn-failed message when the update install cannot be spawned', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake({ throwOn: [0] });
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, packageManager: 'npm', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'update']));
    expect(out).toContain('update failed');
  });

  it('auto-detects the package manager from the bin path when packageManager is not supplied', async () => {
    const logger = createLoggerFake();
    const runner = createProcessRunnerFake();
    const previousArgv = process.argv[1];
    process.argv[1] = '/Users/anyone/.bun/install/global/node_modules/ask-marcel-office-cli/dist/cli.js';
    try {
      const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: runner, fs: createFileSystemFake() });
      await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'update']));
      expect(runner.calls[0]).toEqual({ command: 'bun', args: ['add', '-g', 'ask-marcel-office-cli@latest'] });
    } finally {
      process.argv[1] = previousArgv;
    }
  });

  it('prints Markdown for a single command when the user runs `docs <cmd>`', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'get-current-user']));
    expect(out).toContain('# `get-current-user`');
    expect(out).toContain('## Example');
  });

  it('renders an unknown-command message when the user runs `docs` with a name that does not exist', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'this-is-not-a-real-command']));
    expect(out).toContain('Unknown command');
    expect(out).toContain('this-is-not-a-real-command');
  });

  it('prints rich Markdown docs when `docs <lifecycle-command>` is invoked, since lifecycle commands now have manifest entries', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'login']));
    expect(out).toContain('# `login`');
    expect(out).toContain('**Category:** Lifecycle');
    expect(out).toContain('Authenticate against Microsoft Graph');
  });

  it('falls back to the unknown-command error when `docs help` is invoked (commander does not register `help` as a regular subcommand)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'help']));
    expect(out).toContain('Unknown command');
    expect(out).toContain('help');
  });

  it("'help-json' subcommand prints the full machine-readable manifest (same shape as docs/commands.json) to stdout", async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), version: '1.0.0', fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help-json']));
    const parsed = JSON.parse(out.trim()) as { package: string; version: string; commands: ReadonlyArray<{ name: string }> };
    expect(parsed.package).toBe('ask-marcel-office-cli');
    expect(parsed.version).toBe('1.0.0');
    expect(parsed.commands.length).toBeGreaterThan(100);
    expect(parsed.commands.some((c) => c.name === 'list-drives')).toBe(true);
  });

  // — `help-json` and `docs` used to bypass --output-path
  // entirely, neither writing the file nor surfacing a "not supported"
  // error. Both now honour the flag and write their text body. On error
  // (e.g. directory path) the same envelope as every other bytes-producing
  // command is surfaced (writeOrPrintText error path).
  it("'help-json' and 'docs' honour --output-path: write to disk + emit savedTo envelope; surface is_directory on a directory path", async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const make = (): ReturnType<typeof buildCli> => buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), version: '1.0.0', fs });
    const helpOut = await captureStream('stdout', () =>
      make().parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/work/test-output/manifest.json', 'help-json'])
    );
    const helpParsed = JSON.parse(helpOut.trim()) as { ok: true; data: { savedTo: string } };
    expect(helpParsed.data.savedTo).toBe('/work/test-output/manifest.json');
    const manifest = JSON.parse(fs.snapshot('/work/test-output/manifest.json') ?? '{}') as { commands: ReadonlyArray<unknown> };
    expect(manifest.commands.length).toBeGreaterThan(100);
    const docsOut = await captureStream('stdout', () =>
      make().parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/work/test-output/get-current-user.md', 'docs', 'get-current-user'])
    );
    const docsParsed = JSON.parse(docsOut.trim()) as { ok: true; data: { savedTo: string } };
    expect(docsParsed.data.savedTo).toBe('/work/test-output/get-current-user.md');
    expect(fs.snapshot('/work/test-output/get-current-user.md') ?? '').toContain('get-current-user');
    const failOut = await captureStream('stdout', async () => {
      try {
        await make().parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/work/test-output/', 'help-json']);
      } catch {
        /* expected */
      }
    });
    const failParsed = JSON.parse(failOut.trim()) as { ok: false; error: string };
    expect(failParsed.error).toContain('must be a file path, not a directory');
  });

  // — `--output-path` to a directory path used to surface
  // Node's `EISDIR: illegal operation on a directory`. Now it returns a
  // clear "must be a file path, not a directory" message.
  it('renders --output-path ending in / as "must be a file path, not a directory" rather than EISDIR', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: inlinedPdf, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync([
          'node',
          'ask-marcel-office',
          '--output',
          'json',
          '--output-path',
          '/work/test-output/',
          'download-drive-item-as-pdf',
          '--drive-id',
          'd1',
          '--item-id',
          'i1',
        ]);
      } catch {
        /* expected */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('must be a file path, not a directory');
  });

  // — `*-as-pdf` commands silently falling back to source
  // bytes used to write a corrupt "PDF" to disk. Now the CLI rejects the
  // write with a clear message naming the actual content type.
  it('renders --output-path: passthrough source bytes (not converted PDF) — save as <contentType>, not .pdf, when the response carries passthrough:true', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const passthroughGraph: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'doc-v1.docx' } }),
      getBinaryElevated: async () => ({
        ok: true,
        value: { contentType: 'application/octet-stream', size: 12, base64: 'JVBERi0=' },
      }),
    };
    const cli = buildCli({ auth: okAuth(), graph: passthroughGraph, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync([
          'node',
          'ask-marcel-office',
          '--output',
          'json',
          '--output-path',
          '/work/test-output/doc-v1.pdf',
          'download-drive-item-version',
          '--drive-id',
          'd1',
          '--item-id',
          'i1',
          '--version-id',
          '1.0',
          '--format',
          'pdf',
        ]);
      } catch {
        /* expected */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('passthrough');
    expect(parsed.error).toContain('application/octet-stream');
    expect(fs.has('/work/test-output/doc-v1.pdf')).toBe(false);
  });

  // — when --output-path parent dir is missing /
  // non-writable, the raw Node `ENOENT: no such file or directory, mkdir
  // '/x'` used to leak through. Now translated to a clear message.
  it('translates a Node ENOENT/mkdir error from --output-path into a clear "parent directory missing or not writable" message', async () => {
    const logger = createLoggerFake();
    const fs: FileSystem = {
      readJson: async () => ({ ok: false, error: { type: 'not_found' as const } }),
      readBytes: async () => ({ ok: false, error: { type: 'not_found' as const } }),
      chmod: async () => ({ ok: true as const, value: undefined }),
      writeText: async () => ({ ok: true, value: undefined }),
      writeBytes: async () => ({ ok: false, error: { type: 'io_failed' as const, message: "ENOENT: no such file or directory, mkdir '/nonexistent-dir-no-write'" } }),
      deleteIfExists: async () => ({ ok: true, value: undefined }),
      deleteDirIfExists: async () => ({ ok: true, value: undefined }),
    };
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: inlinedPdf, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync([
        'node',
        'ask-marcel-office',
        '--output',
        'json',
        '--output-path',
        '/nonexistent-dir-no-write/test.pdf',
        'download-drive-item-as-pdf',
        '--drive-id',
        'd1',
        '--item-id',
        'i1',
      ])
    );
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('parent directory missing or not writable');
    expect(parsed.error).toContain('/nonexistent-dir-no-write');
  });

  it("renders a commander 'required option not specified' error for `search-mail-messages` without --query (asserts meta.options.required: true survives manifest → Commander wiring)", async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'search-mail-messages']);
      } catch {
        /* expected — commander throws via exitOverride on missing required option */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('required option');
    expect(parsed.error).toContain('--query');
  });

  it('routes commander parser errors (unknown option) to the JSON envelope on stdout, not stderr plain text (under --output json)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const stderrOut = await captureStream('stderr', async () => {
      const stdoutOut = await captureStream('stdout', async () => {
        try {
          await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'list-drives', '--no-such-flag']);
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

  it('routes commander parser errors as plain "error: ..." lines on stdout by default (text mode), nothing on stderr', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const stderrOut = await captureStream('stderr', async () => {
      const stdoutOut = await captureStream('stdout', async () => {
        try {
          await cli.parseAsync(['node', 'ask-marcel-office', 'list-drives', '--no-such-flag']);
        } catch {
          /* expected */
        }
      });
      expect(stdoutOut.startsWith('error: ')).toBe(true);
      expect(stdoutOut).toContain('--no-such-flag');
      expect(stdoutOut.endsWith('\n')).toBe(true);
    });
    expect(stderrOut).toBe('');
  });

  it('routes commander parser errors (unknown subcommand) to the JSON envelope on stdout (under --output json)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'this-command-does-not-exist']);
      } catch {
        /* expected */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: boolean; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('this-command-does-not-exist');
  });

  it('routes commander parser errors (missing required option) to the JSON envelope on stdout (under --output json)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    // get-mail-attachment keeps commander `requiredOption` (its flags carry no
    // aliases). get-mail-message can't be used here any more — its `--id` alias
    // moves required-ness to schema validation (commander no longer throws).
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'get-mail-attachment']);
      } catch {
        /* expected */
      }
    });
    const parsed = JSON.parse(out.trim()) as { ok: boolean; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error.toLowerCase()).toContain('required');
  });

  it('rejects an invalid --output value with a plain "error: ..." line (text mode is the default for the error too)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'bogus', 'get-current-user']);
      } catch {
        /* expected */
      }
    });
    expect(out.startsWith('error: ')).toBe(true);
    expect(out).toContain("'bogus'");
  });

  it('global --output-path writes inline base64 bytes to disk and replaces base64 with savedTo in the envelope (under --output json)', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: inlinedPdf, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync([
        'node',
        'ask-marcel-office',
        '--output',
        'json',
        '--output-path',
        '/work/test-output/may-deck.pdf',
        'download-drive-item-as-pdf',
        '--drive-id',
        'd1',
        '--item-id',
        'i1',
      ])
    );
    const parsed = JSON.parse(out.trim()) as { ok: true; data: { contentType: string; size: number; savedTo: string; base64?: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.savedTo).toBe('/work/test-output/may-deck.pdf');
    expect(parsed.data.base64).toBeUndefined();
    const written = fs.snapshotBytes('/work/test-output/may-deck.pdf');
    expect(written).toBeDefined();
    if (written) expect(Array.from(written)).toEqual([0x25, 0x50, 0x44, 0x46, 0x2d]);
  });

  it('global --output-path writes a text body via writeText for markdown/plain-text returning commands (under --output json)', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const textGraph: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'notes.md' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'text/plain', size: 5, text: 'hello' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: textGraph, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync([
        'node',
        'ask-marcel-office',
        '--output',
        'json',
        '--output-path',
        '/work/test-output/notes.md',
        'download-drive-item-as-markdown',
        '--drive-id',
        'd1',
        '--item-id',
        'i1',
      ])
    );
    const parsed = JSON.parse(out.trim()) as { ok: true; data: { savedTo: string; text?: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.savedTo).toBe('/work/test-output/notes.md');
    expect(parsed.data.text).toBeUndefined();
    expect(fs.snapshot('/work/test-output/notes.md')).toBe('hello');
  });

  it('global --output-path is a no-op when the command returns plain JSON (no base64 / no text) — surfaces a clear error rather than silently no-op-ing (under --output json)', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({ displayName: 'Jordan' }), logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', async () => {
      try {
        await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/work/test-output/profile.json', 'get-current-user']);
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

  it('global --output-path surfaces a write_failed error envelope when the filesystem rejects the write (e.g. permission denied) (under --output json)', async () => {
    const logger = createLoggerFake();
    const fs: FileSystem = {
      readJson: async () => ({ ok: false, error: { type: 'not_found' as const } }),
      readBytes: async () => ({ ok: false, error: { type: 'not_found' as const } }),
      chmod: async () => ({ ok: true as const, value: undefined }),
      writeText: async () => ({ ok: true, value: undefined }),
      writeBytes: async () => ({ ok: false, error: { type: 'io_failed' as const, message: 'EACCES: permission denied, open' } }),
      deleteIfExists: async () => ({ ok: true, value: undefined }),
      deleteDirIfExists: async () => ({ ok: true, value: undefined }),
    };
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: inlinedPdf, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', '--output-path', '/root/forbidden.pdf', 'download-drive-item-as-pdf', '--drive-id', 'd1', '--item-id', 'i1'])
    );
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('--output-path: write failed');
    expect(parsed.error).toContain('EACCES');
  });

  it('`ask-marcel-office` with NO subcommand prints --help to stdout instead of silently exiting 1', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office']));
    expect(out).toContain('Usage: ask-marcel-office');
    expect(out).toContain('login');
    expect(out).toContain('list-drives');
  });

  it('`help <unknown>` returns a JSON-envelope error rather than silently exiting (under --output json; )', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'help', 'no-such-command']));
    const parsed = JSON.parse(out.trim()) as { ok: false; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toContain('no-such-command');
  });

  it('`help <known>` prints the markdown docs (alias of `docs <known>`)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help', 'list-drives']));
    expect(out).toContain('list-drives');
  });

  it('`help` with no argument prints the global --help text', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'help']));
    expect(out).toContain('Usage: ask-marcel-office');
  });

  it('omitting --output-path leaves the JSON envelope unchanged (existing consumers still get base64) (under --output json)', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 5, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: inlinedPdf, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'download-drive-item-as-pdf', '--drive-id', 'd1', '--item-id', 'i1'])
    );
    const parsed = JSON.parse(out.trim()) as { ok: true; data: { base64?: string; savedTo?: string } };
    expect(parsed.ok).toBe(true);
    expect(parsed.data.base64).toBe('JVBERi0=');
    expect(parsed.data.savedTo).toBeUndefined();
  });

  it('omitting --output-path on a binary command in text mode replaces base64 with a "use --output-path" hint so multi-MB blobs do not flood stdout', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const inlinedPdf: GraphClient = {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'q3.docx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/pdf', size: 12345, base64: 'JVBERi0=' } }),
    };
    const cli = buildCli({ auth: okAuth(), graph: inlinedPdf, logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'download-drive-item-as-pdf', '--drive-id', 'd1', '--item-id', 'i1']));
    expect(out).toBe('binary: application/pdf, 12345 bytes — use --output-path to save\n');
  });

  const mediaGraph = async (): Promise<GraphClient> => {
    const bytes = await buildMediaSamples();
    let binary = '';
    for (const byte of bytes) binary += String.fromCharCode(byte);
    return {
      ...okGraph({}),
      get: async () => ({ ok: true, value: { name: 'deck.pptx' } }),
      getBinary: async () => ({ ok: true, value: { contentType: 'application/octet-stream', size: bytes.byteLength, base64: btoa(binary) } }),
    };
  };

  it('extract-drive-item-images --output-dir writes every image to the directory and replaces base64 with savedTo', async () => {
    const logger = createLoggerFake();
    const fs = createFileSystemFake();
    const cli = buildCli({ auth: okAuth(), graph: await mediaGraph(), logger, processRunner: createProcessRunnerFake(), fs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'extract-drive-item-images', '--drive-id', 'd1', '--item-id', 'i1', '--output-dir', '/work/imgs'])
    );
    expect(out).toContain('savedTo');
    expect(out).toContain('/work/imgs/word_media_image1.png'); // full path flattened, not basename
    expect(out).not.toContain('base64');
    expect(Array.from(fs.snapshotBytes('/work/imgs/word_media_image1.png') ?? [])).toEqual([0x89, 0x50, 0x4e, 0x47]);
  });

  it('--output-dir on a command that returns no media array emits a clear error pointing at the image-extraction commands', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({ displayName: 'Jordan' }), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'get-current-user', '--output-dir', '/work/imgs']));
    expect(out).toContain('did not return a media array');
    expect(out).toContain('extract-drive-item-images');
    expect(out).toContain('extract-mail-attachment-images');
  });

  it('rejects an empty --output-dir explicitly', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: await mediaGraph(), logger, processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'extract-drive-item-images', '--drive-id', 'd1', '--item-id', 'i1', '--output-dir', ''])
    );
    expect(out).toContain('directory argument is empty');
  });

  it('surfaces a filesystem write failure from --output-dir', async () => {
    const logger = createLoggerFake();
    const failingFs: FileSystem = {
      readJson: async () => ({ ok: false, error: { type: 'not_found' } }),
      readBytes: async () => ({ ok: false, error: { type: 'not_found' } }),
      chmod: async () => ({ ok: true, value: undefined }),
      writeText: async () => ({ ok: true, value: undefined }),
      writeBytes: async () => ({ ok: false, error: { type: 'io_failed', message: 'ENOSPC: no space left on device' } }),
      deleteIfExists: async () => ({ ok: true, value: undefined }),
      deleteDirIfExists: async () => ({ ok: true, value: undefined }),
    };
    const cli = buildCli({ auth: okAuth(), graph: await mediaGraph(), logger, processRunner: createProcessRunnerFake(), fs: failingFs });
    const out = await captureStream('stdout', () =>
      cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'extract-drive-item-images', '--drive-id', 'd1', '--item-id', 'i1', '--output-dir', '/work/imgs'])
    );
    expect(out).toContain('write failed');
    expect(out).toContain('ENOSPC');
  });
});
