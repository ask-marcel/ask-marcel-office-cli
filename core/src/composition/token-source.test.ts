import { describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '../test-helpers/jwt.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import { createEnvThenHelperTokenSource } from './token-source.ts';

/*
 * The token source of the packages that do not sign in themselves: a tier's
 * environment variable when it is set, otherwise the token helper of the auth
 * package, found through the locator file under the home folder.
 */

const HOME = '/virtual/home';
const LOCATOR = '/virtual/home/.ask-marcel/token-helper.json';
const HELPER = '/opt/ask-marcel/auth/dist/token.js';
const NODE = '/usr/local/bin/node';
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const graphToken = (from: string): AccessToken => accessTokenUnsafe(unsignedJwt({ exp: inAnHour(), aud: 'https://graph.microsoft.com', from }));

const HELPER_TOKEN = graphToken('helper');

const answering = (): ReturnType<typeof createProcessRunnerFake> =>
  createProcessRunnerFake(() => ok({ exitCode: 0, stdout: `${JSON.stringify({ accessToken: HELPER_TOKEN, expiresOn: inAnHour() })}\n` }));

const refusing = (): ReturnType<typeof createProcessRunnerFake> =>
  createProcessRunnerFake(() =>
    ok({ exitCode: 1, stdout: `${JSON.stringify({ errorCode: 'secondary_token_unavailable', tier: 'ic3', message: 'no token', remedy: 'run login' })}\n` })
  );

const messageOf = (result: { readonly ok: boolean; readonly error?: unknown }): string => {
  const error = result.error as { readonly message?: string } | undefined;
  return error?.message ?? '';
};

describe('the token source of a package that takes its tokens from the auth package', () => {
  it('takes the token of a tier variable and starts no helper', async () => {
    const runner = answering();
    const own = graphToken('env');
    const source = createEnvThenHelperTokenSource({ home: HOME, env: { ASKMARCEL_TOKEN_BASIC: own }, fs: createFileSystemFake(), runner, interactive: false });
    expect(await source.graphToken('basic')).toEqual(ok(own));
    expect(runner.calls).toEqual([]);
  });

  it('runs the helper that the locator file under the home folder names', async () => {
    const runner = answering();
    const fs = createFileSystemFake();
    fs.seed(LOCATOR, JSON.stringify({ execPath: NODE, entry: HELPER, version: '1.0.0' }));
    fs.seed(HELPER, '');
    fs.seed(NODE, '');
    const source = createEnvThenHelperTokenSource({ home: HOME, env: {}, fs, runner, interactive: false, platform: 'linux', execPath: '/usr/bin/bun' });
    expect(await source.graphToken('elevated')).toEqual(ok(HELPER_TOKEN));
    expect(runner.calls.map((call) => [call.command, ...call.args])).toEqual([[NODE, HELPER, '--tier', 'elevated']]);
  });

  it('ends a refused token with the commands of the package that need it', async () => {
    const source = createEnvThenHelperTokenSource({
      home: HOME,
      env: { ASKMARCEL_TOKEN_COMMAND: HELPER },
      fs: createFileSystemFake(),
      runner: refusing(),
      interactive: false,
      neededBy: { elevated: [], chatsvcagg: [], ic3: ['list-teams-chat-history'] },
    });
    expect(messageOf(await source.substrateToken('ic3'))).toEndWith('(Commands that need this token: list-teams-chat-history.)');
  });

  it('adds no list of commands when the package names none', async () => {
    const source = createEnvThenHelperTokenSource({ home: HOME, env: { ASKMARCEL_TOKEN_COMMAND: HELPER }, fs: createFileSystemFake(), runner: refusing(), interactive: false });
    expect(messageOf(await source.substrateToken('ic3'))).toEndWith('to read its message and remedy.');
  });

  it('looks for the auth bin on PATH with the platform of this process when no locator names a helper', async () => {
    const runner = answering();
    const source = createEnvThenHelperTokenSource({ home: HOME, env: {}, fs: createFileSystemFake(), runner, interactive: true });
    expect(await source.graphToken('basic')).toEqual(ok(HELPER_TOKEN));
    expect(runner.calls.map((call) => [call.command, ...call.args, call.options.stdin])).toEqual([['ask-marcel-office-auth', 'token', '--tier', 'basic', 'inherit']]);
  });
});
