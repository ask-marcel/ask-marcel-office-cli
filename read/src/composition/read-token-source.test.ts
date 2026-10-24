import { afterEach, beforeEach, describe, expect, it } from 'bun:test';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import type { ProcessRunnerCall } from '@ask-marcel/office-core/test-helpers/process-runner-fake.ts';
import { createProcessRunnerFake } from '@ask-marcel/office-core/test-helpers/process-runner-fake.ts';
import type { CommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { commands } from '../use-cases/commands/index.ts';
import { buildReadTokenSource, createEnvThenHelperTokenSource, READ_NEEDED_BY } from './read-token-source.ts';

// The read package signs only through a tier variable or the token helper of
// the auth package (LESSONS 2026-10-18): it holds no token cache and opens no
// browser. The tokens here are built at run time, with no signature. The
// library case sets the tier variable of this process and gives it back after.
// The cases with no home or no interactive option read the home folder and
// the stdin of this process, and change neither.

const HOME = '/virtual/home';
const HELPER = '/opt/ask-marcel/auth/dist/token.js';
const RUNTIME = '/opt/node/bin/node';
const VARIABLE = 'ASKMARCEL_TOKEN_BASIC';
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const graphToken = (from: string): string => unsignedJwt({ aud: 'https://graph.microsoft.com', exp: inAnHour(), from });

// The helper answers each tier with its token, the way the auth package does.
const answeringHelper = (token: string): ReturnType<typeof createProcessRunnerFake> =>
  createProcessRunnerFake((_call: ProcessRunnerCall) => ok({ exitCode: 0, stdout: `${JSON.stringify({ accessToken: token, expiresOn: inAnHour() })}\n` }));

// The helper refuses the ic3 token, the way the auth package does when no browser sign-in got it.
const refusingHelper = (): ReturnType<typeof createProcessRunnerFake> =>
  createProcessRunnerFake(() =>
    ok({ exitCode: 1, stdout: `${JSON.stringify({ errorCode: 'secondary_token_unavailable', tier: 'ic3', message: 'no token', remedy: 'run login' })}\n` })
  );

const registryNamesWhere = (matches: (meta: CommandMeta) => boolean): ReadonlyArray<string> =>
  Object.entries(commands).flatMap(([name, command]) => (matches(command.meta) ? [name] : []));

const inNameOrder = (names: ReadonlyArray<string>): ReadonlyArray<string> => names.toSorted((a, b) => a.localeCompare(b));

let saved: string | undefined;

beforeEach(() => {
  saved = process.env[VARIABLE];
});

afterEach(() => {
  if (saved === undefined) delete process.env[VARIABLE];
  else process.env[VARIABLE] = saved;
});

describe('the token source of the read package', () => {
  it('names, for each secondary token, the read commands that need it, in name order, from the read registry', () => {
    expect(READ_NEEDED_BY).toEqual({
      elevated: inNameOrder(registryNamesWhere((meta) => meta.needsElevatedToken === true)),
      chatsvcagg: inNameOrder(registryNamesWhere((meta) => meta.needsSubstrateToken === 'chatsvcagg')),
      ic3: inNameOrder(registryNamesWhere((meta) => meta.needsSubstrateToken === 'ic3')),
    });
    expect(READ_NEEDED_BY.ic3).toContain('list-teams-chat-history');
    expect(READ_NEEDED_BY.chatsvcagg).toContain('list-teams-chat-messages');
    expect(READ_NEEDED_BY.elevated.length).toBeGreaterThan(1);
  });

  it('takes the tier variable ahead of the helper', async () => {
    const own = graphToken('env');
    const runner = answeringHelper(graphToken('helper'));
    const source = buildReadTokenSource({
      env: { ASKMARCEL_TOKEN_BASIC: own, ASKMARCEL_TOKEN_COMMAND: HELPER },
      home: HOME,
      fs: createFileSystemFake(),
      runner,
      interactive: false,
    });
    expect(await source.graphToken('basic')).toEqual(ok(accessTokenUnsafe(own)));
    expect(runner.calls).toEqual([]);
  });

  it('runs the helper at ASKMARCEL_TOKEN_COMMAND when no tier variable is set, with no stdin in a session with no person', async () => {
    const token = graphToken('helper');
    const runner = answeringHelper(token);
    const source = buildReadTokenSource({ env: { ASKMARCEL_TOKEN_COMMAND: HELPER }, home: HOME, fs: createFileSystemFake(), runner, interactive: false });
    expect(await source.graphToken('basic')).toEqual(ok(accessTokenUnsafe(token)));
    expect(runner.calls.map((call) => [call.command, ...call.args, call.options.stdin])).toEqual([[HELPER, '--tier', 'basic', 'ignore']]);
  });

  it('runs the helper that the auth package recorded under the home folder of this process when the caller names no home', async () => {
    const fs = createFileSystemFake();
    fs.seed(join(homedir(), '.ask-marcel', 'token-helper.json'), JSON.stringify({ execPath: RUNTIME, entry: HELPER, version: '1.0.0' }));
    fs.seed(RUNTIME, '');
    fs.seed(HELPER, '');
    const runner = answeringHelper(graphToken('helper'));
    await buildReadTokenSource({ env: {}, fs, runner, interactive: false }).graphToken('basic');
    expect(runner.calls.map((call) => [call.command, ...call.args])).toEqual([[RUNTIME, HELPER, '--tier', 'basic']]);
  });

  it('gives the helper the stdin of a terminal only when this process has one, when the caller does not say', async () => {
    const runner = answeringHelper(graphToken('helper'));
    await buildReadTokenSource({ env: { ASKMARCEL_TOKEN_COMMAND: HELPER }, home: HOME, fs: createFileSystemFake(), runner }).graphToken('basic');
    expect(runner.calls.map((call) => call.options.stdin)).toEqual([process.stdin.isTTY === true ? 'inherit' : 'ignore']);
  });

  it('ends a refused ic3 token with the read commands that need it', async () => {
    const source = buildReadTokenSource({ env: { ASKMARCEL_TOKEN_COMMAND: HELPER }, home: HOME, fs: createFileSystemFake(), runner: refusingHelper(), interactive: false });
    const result = await source.substrateToken('ic3');
    const error = result.ok ? undefined : result.error;
    expect(error?.type).toBe('auth_failed');
    if (error?.type === 'auth_failed')
      expect([error.code, error.message.endsWith(`(Commands that need this token: ${READ_NEEDED_BY.ic3.join(', ')}.)`)]).toEqual(['secondary_token_unavailable', true]);
  });

  it('gives a library caller a token source with no parameter to pass, which signs with the tier variable of this process', async () => {
    const token = graphToken('library');
    process.env[VARIABLE] = token;
    expect(await createEnvThenHelperTokenSource().graphToken('basic')).toEqual(ok(accessTokenUnsafe(token)));
  });
});
