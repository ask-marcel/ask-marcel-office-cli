import { describe, expect, it } from 'bun:test';
import { err, ok } from '../domain/result.ts';
import { tenantIdUnsafe } from '../domain/tenant-id.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import { buildDeps } from './build-deps.ts';
import { createEnvThenHelperTokenSource } from './token-source.ts';

// Auth owns the remedy of a token failure; the package that asked for the token
// adds the commands of its own that need that token (plan, "Token protocol").

const HOME = '/virtual/home';
const HELPER = '/opt/ask-marcel/dist/token.js';

// The helper refuses every tier the way the auth package does.
const refusingRunner = (): ReturnType<typeof createProcessRunnerFake> =>
  createProcessRunnerFake(() =>
    ok({ exitCode: 1, stdout: `${JSON.stringify({ errorCode: 'secondary_token_unavailable', tier: 'ic3', message: 'no token', remedy: 'run login' })}\n` })
  );

const NEEDED_BY = { elevated: ['get-chat', 'list-chats'], chatsvcagg: ['list-teams-chat-messages'], ic3: ['list-teams-chat-history', 'get-teams-chat-message'] };

const sourceRefusing = (neededBy?: typeof NEEDED_BY): ReturnType<typeof createEnvThenHelperTokenSource> =>
  createEnvThenHelperTokenSource({
    env: { ASKMARCEL_TOKEN_COMMAND: HELPER },
    home: HOME,
    fs: createFileSystemFake(),
    runner: refusingRunner(),
    interactive: false,
    ...(neededBy ? { neededBy } : {}),
  });

const messageOf = (result: { readonly ok: boolean; readonly error?: unknown }): string => {
  const error = result.error as { readonly message?: string } | undefined;
  return error?.message ?? '';
};

describe('a token failure names the commands of this package that need the token', () => {
  it('ends a refused ic3 token with the commands that need ic3, and keeps the code auth gave', async () => {
    const result = await sourceRefusing(NEEDED_BY).substrateToken('ic3');
    expect(result.ok).toBe(false);
    expect(messageOf(result)).toEndWith('to read its message and remedy. (Commands that need this token: list-teams-chat-history, get-teams-chat-message.)');
    if (!result.ok) expect(result.error).toMatchObject({ type: 'auth_failed', code: 'secondary_token_unavailable' });
  });

  it('names the chatsvcagg commands for the chatsvcagg token and its region, and the elevated ones for the elevated token', async () => {
    const source = sourceRefusing(NEEDED_BY);
    expect(messageOf(await source.substrateToken('chatsvcagg'))).toEndWith('(Commands that need this token: list-teams-chat-messages.)');
    expect(messageOf(await source.substrateRegion('chatsvcagg'))).toEndWith('(Commands that need this token: list-teams-chat-messages.)');
    expect(messageOf(await source.graphToken('elevated'))).toEndWith('(Commands that need this token: get-chat, list-chats.)');
  });

  it('adds no list to the basic and the guest token, which every command needs', async () => {
    const source = sourceRefusing(NEEDED_BY);
    expect(messageOf(await source.graphToken('basic'))).toEndWith('to read its message and remedy.');
    expect(messageOf(await source.guestToken(tenantIdUnsafe('6f1e3a92-4b7c-4d51-9e2f-8a3b5c7d1e04')))).toEndWith('to read its message and remedy.');
  });

  it('adds no list when the package names no command', async () => {
    expect(messageOf(await sourceRefusing().substrateToken('ic3'))).toEndWith('to read its message and remedy.');
  });

  it('adds no list to a failure that is not a missing token: a helper path that is not absolute, a bad token in the environment, a helper that never started', async () => {
    const source = (env: Record<string, string>, runner = refusingRunner()): ReturnType<typeof createEnvThenHelperTokenSource> =>
      createEnvThenHelperTokenSource({ env, home: HOME, fs: createFileSystemFake(), runner, interactive: false, neededBy: NEEDED_BY });
    const notAbsolute = await source({ ASKMARCEL_TOKEN_COMMAND: '1' }).graphToken('elevated');
    expect(notAbsolute).toMatchObject({ ok: false, error: { code: 'token_helper_unavailable' } });
    expect(messageOf(notAbsolute)).not.toContain('Commands that need');
    const badToken = await source({ ASKMARCEL_TOKEN_IC3: 'not-a-token' }).substrateToken('ic3');
    expect(badToken).toMatchObject({ ok: false, error: { code: 'env_token_invalid' } });
    expect(messageOf(badToken)).not.toContain('Commands that need');
    const neverStarted = createProcessRunnerFake(() => err({ type: 'spawn_failed', message: 'ENOENT' }));
    const spawnFailed = await source({ ASKMARCEL_TOKEN_COMMAND: HELPER }, neverStarted).substrateToken('chatsvcagg');
    expect(spawnFailed.ok).toBe(false);
    expect(messageOf(spawnFailed)).not.toContain('Commands that need');
  });

  it('lists the registry commands that need the elevated token when the CLI signs through the helper', async () => {
    const deps = buildDeps({ home: HOME, env: { ASKMARCEL_TOKEN_COMMAND: HELPER }, fs: createFileSystemFake(), processRunner: refusingRunner(), interactive: false });
    const result = await deps.graph.getElevated('/me/chats');
    expect(messageOf(result)).toEndWith('(Commands that need this token: diff-drive-item-versions, download-drive-item-version, get-chat, get-user, list-chats.)');
  });
});
