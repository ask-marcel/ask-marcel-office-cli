import { afterEach, describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import type { FetchMock } from '../test-helpers/fetch-mock.ts';
import { installFetchMock } from '../test-helpers/fetch-mock.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import type { ProcessRunnerCall } from '../test-helpers/process-runner-fake.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import { buildDeps } from './build-deps.ts';
import { createEnvThenHelperTokenSource } from './token-source.ts';

const HOME = '/virtual/home';
const CACHE = '/virtual/home/.ask-marcel/token-cache.json';
const LOCATOR = '/virtual/home/.ask-marcel/token-helper.json';
const HELPER = '/opt/ask-marcel/dist/token.js';
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const jwt = (claims: Record<string, unknown>): AccessToken => accessTokenUnsafe(`${segment({ alg: 'RS256' })}.${segment(claims)}.sig`);

const HELPER_TOKEN = jwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com', from: 'helper' });

// The helper answers every tier with its token, and the chat tiers with a region.
const helperRunner = (): ReturnType<typeof createProcessRunnerFake> =>
  createProcessRunnerFake((call: ProcessRunnerCall) => {
    const chat = call.args.includes('ic3') || call.args.includes('chatsvcagg');
    return ok({ exitCode: 0, stdout: `${JSON.stringify({ accessToken: HELPER_TOKEN, expiresOn: inAnHour(), ...(chat ? { region: 'amer' } : {}) })}\n` });
  });

let mock: FetchMock | undefined;
afterEach(() => mock?.restore());

const graphAnswers = (): FetchMock =>
  installFetchMock([{ match: () => true, respond: () => new Response(JSON.stringify({ id: 'me' }), { headers: { 'content-type': 'application/json' } }) }]);

const authorizationOf = (fetches: FetchMock, at: number): string | undefined => {
  const headers = fetches.calls[at]?.init?.headers as Record<string, string> | undefined;
  return headers?.['Authorization'];
};

describe('the token source the CLI and the MCP server sign with', () => {
  it('signs with the token the helper at ASKMARCEL_TOKEN_COMMAND prints, when that variable is set', async () => {
    mock = graphAnswers();
    const runner = helperRunner();
    const deps = buildDeps({ home: HOME, env: { ASKMARCEL_TOKEN_COMMAND: HELPER }, fs: createFileSystemFake(), processRunner: runner, interactive: false });
    expect(await deps.graph.get('/me')).toEqual(ok({ id: 'me' }));
    expect(authorizationOf(mock, 0)).toBe(`Bearer ${HELPER_TOKEN}`);
    expect(runner.calls.map((call) => [call.command, ...call.args, call.options.stdin])).toEqual([[HELPER, '--tier', 'basic', 'ignore']]);
  });

  it('takes a tier variable ahead of the helper once the helper is selected', async () => {
    mock = graphAnswers();
    const runner = helperRunner();
    const own = jwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com', from: 'env' });
    const deps = buildDeps({
      home: HOME,
      env: { ASKMARCEL_TOKEN_COMMAND: HELPER, ASKMARCEL_TOKEN_BASIC: own },
      fs: createFileSystemFake(),
      processRunner: runner,
      interactive: false,
    });
    await deps.graph.get('/me');
    expect(authorizationOf(mock, 0)).toBe(`Bearer ${own}`);
    expect(runner.calls).toEqual([]);
  });

  it('keeps the in-process token cache as the default: without ASKMARCEL_TOKEN_COMMAND neither a tier variable nor a helper is used', async () => {
    mock = graphAnswers();
    const runner = helperRunner();
    const cached = jwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com', from: 'cache' });
    const fs = createFileSystemFake();
    fs.seed(CACHE, JSON.stringify({ access_token: cached, expires_on: inAnHour(), refresh_token: 'rt' }));
    fs.seed(LOCATOR, JSON.stringify({ execPath: '/usr/local/bin/node', entry: HELPER, version: '2.8.0' }));
    fs.seed(HELPER, '');
    const own = jwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com', from: 'env' });
    const deps = buildDeps({ home: HOME, env: { ASKMARCEL_TOKEN_BASIC: own }, fs, processRunner: runner, interactive: false });
    await deps.graph.get('/me');
    expect(authorizationOf(mock, 0)).toBe(`Bearer ${cached}`);
    expect(runner.calls).toEqual([]);
  });

  it('runs the helper once for an ic3 chat read, which needs both the region and the token', async () => {
    mock = graphAnswers();
    const runner = helperRunner();
    const deps = buildDeps({ home: HOME, env: { ASKMARCEL_TOKEN_COMMAND: HELPER }, fs: createFileSystemFake(), processRunner: runner, interactive: true });
    await deps.graph.teamsChatIc3('/v1/users/ME/conversations');
    expect(mock.calls[0]?.url).toBe('https://teams.microsoft.com/api/chatsvc/amer/v1/users/ME/conversations');
    expect(runner.calls.map((call) => [...call.args, call.options.stdin])).toEqual([['--tier', 'ic3', 'inherit']]);
  });

  it('finds the helper through the locator file auth writes when ASKMARCEL_TOKEN_COMMAND is not set', async () => {
    const runner = helperRunner();
    const fs = createFileSystemFake();
    fs.seed(LOCATOR, JSON.stringify({ execPath: '/usr/local/bin/node', entry: HELPER, version: '2.8.0' }));
    fs.seed(HELPER, '');
    fs.seed('/usr/local/bin/node', '');
    const source = createEnvThenHelperTokenSource({ home: HOME, env: {}, fs, runner, interactive: false, platform: 'linux', execPath: '/usr/bin/bun' });
    expect(await source.graphToken('elevated')).toEqual(ok(HELPER_TOKEN));
    expect(runner.calls.map((call) => [call.command, ...call.args])).toEqual([['/usr/local/bin/node', HELPER, '--tier', 'elevated']]);
  });

  it('runs a Windows shim script with this runtime, which it defaults to', async () => {
    const runner = helperRunner();
    const fs = createFileSystemFake();
    fs.seed('C:\\npm\\ask-marcel-office-auth.cmd', '"%dp0%\\auth\\cli.js" %*');
    const source = createEnvThenHelperTokenSource({ home: HOME, env: { PATH: 'C:\\npm' }, fs, runner, interactive: false, platform: 'win32' });
    await source.guestToken(tenantIdUnsafe('11111111-2222-3333-4444-555555555555'));
    expect(runner.calls[0]?.command).toBe(process.execPath);
  });
});
