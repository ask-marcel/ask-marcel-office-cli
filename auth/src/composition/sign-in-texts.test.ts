import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import type { AuthError, AuthLadder } from '../infra/auth.ts';
import { createAuthManagerFromApi } from '../infra/auth-browser.ts';
import type { BrowserAuth, BrowserAuthApi, ContextLike, PageLike } from '../infra/browser-auth.ts';
import { createBrowserAuthFromApi } from '../infra/browser-auth.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { buildAuthCli } from './auth-program.ts';

// The sign-in texts that tell a person what to install or run name the auth
// package and the auth bin (package split, phase 4), never the single package
// `ask-marcel-office-cli` or its bin, which no longer has the sign-in commands.

const INSTALL_UNDER_NODE = 'npm i -g @ask-marcel/office-auth';

const stdoutOf = async (run: () => Promise<unknown>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run();
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

const cancelledAuth = (): AuthLadder => fakeAuthManager({ getAccessToken: async () => ({ ok: false, error: { type: 'auth_cancelled' } as AuthError }) });

// A browser whose every elevated launch hangs past its budget.
const hangingBrowser = (): BrowserAuth => {
  const timedOut = { ok: false as const, reason: 'launch_timeout' as const };
  return {
    acquireElevatedToken: async () => timedOut,
    acquireChatsvcaggToken: async () => timedOut,
    acquireIc3Token: async () => timedOut,
    acquireBothTokens: async () => ({ teams: null, elevated: timedOut, chatsvcagg: timedOut, ic3: timedOut }),
    close: async () => {},
  };
};

// A browser profile that is still signed in: every page shows Teams and no
// token comes back, so the sign-in clears the session before it waits.
const signedInProfile = (): BrowserAuthApi => {
  const page: PageLike = { on: () => {}, goto: async () => undefined, url: () => 'https://teams.microsoft.com/v2/', evaluate: async () => undefined, close: async () => {} };
  const context: ContextLike = { newPage: async () => page, clearCookies: async () => {}, close: async () => {} };
  return { launchPersistentContext: async () => context };
};

describe('the sign-in texts', () => {
  it('tell a person whose login was cancelled under Bun to install the auth package under Node', async () => {
    const cli = buildAuthCli({ auth: cancelledAuth(), logger: createLoggerFake(), fs: createFileSystemFake() });

    const out = await stdoutOf(() => cli.parseAsync(['node', 'ask-marcel-office-auth', 'login']));

    expect(out).toContain(`\`${INSTALL_UNDER_NODE}\`, then \`ask-marcel-office-auth login\``);
    expect(out).not.toContain('ask-marcel-office-cli');
  });

  it('tell a person whose elevated browser launch timed out to install the auth package under Node', async () => {
    const auth = createAuthManagerFromApi(hangingBrowser(), '/home/.ask-marcel/token-cache.json', '/home/.ask-marcel/browser-profile', createLoggerFake(), createFileSystemFake());

    const result = await auth.getElevatedAccessToken();

    expect(result.ok).toBe(false);
    const message = result.ok || result.error.type !== 'auth_failed' ? '' : result.error.message;
    expect(message).toContain('elevated browser launch timed out');
    expect(message).toContain(`\`${INSTALL_UNDER_NODE}\``);
    expect(message).not.toContain('ask-marcel-office-cli');
  });

  it('tell a person whose browser session is about to be cleared that a plain login of the auth bin keeps it', async () => {
    const progress: string[] = [];
    const browser = createBrowserAuthFromApi(signedInProfile(), {
      logger: createLoggerFake(),
      fs: createFileSystemFake(),
      trace: () => {},
      onProgress: (line) => progress.push(line),
      profileDir: '/virtual/.ask-marcel/browser-profile',
      initialSettleMs: 1,
      postReloginSettleMs: 1,
      pollIntervalMs: 1,
      pollDeadlineMs: 20,
      navigationTimeoutMs: 50,
      elevatedPreferenceGraceMs: 1,
    });

    await browser.acquireBothTokens('https://teams.microsoft.com');

    const wipe = progress.filter((line) => line.includes('"Stay signed in"'));
    expect(wipe).toHaveLength(1);
    expect(wipe[0]).toContain('A plain `ask-marcel-office-auth login` refreshes the elevated token without clearing it.');
  });
});
