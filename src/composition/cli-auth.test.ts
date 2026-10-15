// The auth cases of the CLI suite: `login`, `status` and `logout`, and the lifecycle docs
// of `login`. Moved out of cli.test.ts with their text unchanged (package split, step 16),
// so the auth package takes this file whole; the helpers are copies of that file's.
import { describe, expect, it } from 'bun:test';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import type { AuthError, AuthLadder } from '../infra/auth.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import type { TokenInfo } from '../use-cases/ports/token-report.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
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

const okAuth = (): AuthLadder => fakeAuthManager();

const cancelledAuth = (): AuthLadder =>
  fakeAuthManager({
    getAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' } as AuthError }),
    logout: async () => ({ ok: false, error: { type: 'auth_cancelled' } as AuthError }),
  });

const failedAuth = (): AuthLadder =>
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

const authWithTokenInfo = (info: TokenInfo): AuthLadder => fakeAuthManager({ getTokenInfo: async () => ({ ok: true, value: info }) });

describe('buildCli command surface', () => {
  it('renders the slim availability summary + two-pointer hint when login succeeds (under --output json)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: authWithTokenInfo(sampleTokenInfo()), graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'login']));
    const parsed = JSON.parse(out) as { data: { status: string; available: string[]; hint: string } };
    expect(parsed.data.status).toBe('authenticated');
    expect(parsed.data.available).toEqual(['basic', 'chatsvcagg', 'ic3']); // sampleTokenInfo has elevated unavailable → omitted
    expect(parsed.data.hint).toContain('ask-marcel-office status');
    expect(parsed.data.hint).toContain('login --force');
  });

  it('renders the slim summary in text format with the two-pointer hint', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: authWithTokenInfo(sampleTokenInfo()), graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
    expect(out).toContain('status: authenticated');
    expect(out).toContain('ask-marcel-office status');
    expect(out).toContain('login --force');
  });

  it('includes a token in login available only when the token report shows it available (elevated)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      auth: authWithTokenInfo(sampleTokenInfo({ elevated: { available: true, expiresInSeconds: 1800, scopes: ['Chat.ReadBasic'], refresh: 'interactive' } })),
      graph: okGraph({}),
      logger,
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'login']));
    const parsed = JSON.parse(out) as { data: { available: string[] } };
    expect(parsed.data.available).toEqual(['basic', 'elevated', 'chatsvcagg', 'ic3']); // elevated now available → listed
  });

  it('reports the tokens that the sign-in just captured, read from the login-configured manager and not from the command-path one', async () => {
    const elevated = { available: true, expiresInSeconds: 1800, scopes: ['Chat.ReadBasic'], refresh: 'interactive' as const };
    const cli = buildCli({
      auth: authWithTokenInfo(sampleTokenInfo()),
      makeLoginAuth: () => authWithTokenInfo(sampleTokenInfo({ elevated })),
      graph: okGraph({}),
      logger: createLoggerFake(),
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'login']));
    const parsed = JSON.parse(out) as { data: { available: string[] } };
    expect(parsed.data.available).toEqual(['basic', 'elevated', 'chatsvcagg', 'ic3']);
  });

  it('login --force forwards the force flag to getAccessToken so a warm session re-captures all tokens', async () => {
    const calls: Array<{ force?: boolean } | undefined> = [];
    const capturingAuth: AuthLadder = {
      ...okAuth(),
      getAccessToken: async (options?: { force?: boolean }) => {
        calls.push(options);
        return { ok: true as const, value: accessTokenUnsafe('tok') };
      },
    };
    const logger = createLoggerFake();
    const cli = buildCli({ auth: capturingAuth, graph: okGraph({}), logger, fs: createFileSystemFake() });
    await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'login', '--force']));
    expect(calls.some((c) => c?.force === true)).toBe(true);
  });

  it('surfaces the token-status read failure when the token report fails after a successful login', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({
      auth: fakeAuthManager({ getTokenInfo: async () => ({ ok: false, error: { type: 'auth_failed', message: 'no token cached' } }) }),
      graph: okGraph({}),
      logger,
      fs: createFileSystemFake(),
    });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
    expect(out).toContain('no token cached');
  });

  it('prints the token report for status under --output json', async () => {
    const auth = fakeAuthManager({ getTokenInfo: async () => ({ ok: true, value: sampleTokenInfo() }) });
    const cli = buildCli({ auth, graph: okGraph({}), logger: createLoggerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'status']));
    const parsed = JSON.parse(out) as { ok: boolean; data: Record<string, { available?: boolean }> };
    expect(parsed.ok).toBe(true);
    expect(Object.keys(parsed.data)).toEqual(['basic', 'elevated', 'chatsvcagg', 'ic3', 'hint']);
    expect(parsed.data['basic']?.available).toBe(true);
    expect(parsed.data['elevated']?.available).toBe(false);
  });

  it('gives the error code not_authenticated when status finds no session in the token cache', async () => {
    const message = 'Not signed in: there is no cached token to inspect. Run `ask-marcel-office login`.';
    const auth = fakeAuthManager({ getTokenInfo: async () => ({ ok: false, error: { type: 'auth_failed', message, code: 'not_authenticated' } }) });
    const cli = buildCli({ auth, graph: okGraph({}), logger: createLoggerFake(), fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'status']));
    const parsed = JSON.parse(out.trim()) as { ok: boolean; errorCode?: string; error: string };
    expect(parsed.ok).toBe(false);
    expect(parsed.errorCode).toBe('not_authenticated');
    expect(parsed.error).toContain('ask-marcel-office login');
  });

  it('renders an Authentication cancelled error when the user closes the browser', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: cancelledAuth(), graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
    expect(out).toContain('Authentication cancelled');
  });

  // "Authentication cancelled" alone reads as "the user closed the window",
  // which is the one thing it is NOT when an endpoint-security agent blocks the
  // local control connection: the browser opens, sits on about:blank, and is
  // never driveable. That case cost four rounds of diagnosis on 2026-09-07, so
  // the message names it and points at the trace.
  it('points a cancelled login at the EDR / control-connection cause, not just the user closing the window', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: cancelledAuth(), graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
    expect(out).toContain('about:blank');
    expect(out).toContain('security');
    expect(out).toContain('ASKMARCEL_TRACE=1');
  });

  // A SentinelOne policy on one Windows box blocked the CDP pipe only under Bun;
  // the same login ran fine under Node (npm install, node shebang). The message
  // should offer that runtime switch, since it is cheaper than a security
  // exclusion.
  it('offers the run-under-Node workaround on a cancelled login', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: cancelledAuth(), graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
    expect(out).toContain('Node');
    expect(out).toContain('Bun');
  });

  it('invokes onCommandError exactly once when a command fails', async () => {
    const logger = createLoggerFake();
    let errorReports = 0;
    const cli = buildCli({
      auth: cancelledAuth(),
      graph: okGraph({}),
      logger,
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
    const cli = buildCli({ auth: failedAuth(), graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'login']));
    expect(out).toContain('browser launch failed');
  });

  it('renders a logged_out envelope when logout succeeds (under --output json)', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', 'logout']));
    expect(out).toContain('"status":"logged_out"');
  });

  it('renders the underlying message when logout fails', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: failedAuth(), graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'logout']));
    expect(out).toContain('rm denied');
  });

  it('prints rich Markdown docs when `docs <lifecycle-command>` is invoked, since lifecycle commands now have manifest entries', async () => {
    const logger = createLoggerFake();
    const cli = buildCli({ auth: okAuth(), graph: okGraph({}), logger, fs: createFileSystemFake() });
    const out = await captureStream('stdout', () => cli.parseAsync(['node', 'ask-marcel-office', 'docs', 'login']));
    expect(out).toContain('# `login`');
    expect(out).toContain('**Category:** Lifecycle');
    expect(out).toContain('Authenticate against Microsoft Graph');
  });
});
