/*
 * The auth lifecycle commands of the single package's CLI: `login`, `status`
 * and `logout`. The CLI builder (`cli-shell.ts`) is core and knows no auth;
 * `cli.ts` hands it these through `addLifecycleCommands`, so they land under
 * the Lifecycle group and print through the same output as every command. After
 * the package split they are the auth package's (docs/plans/2026-10-01-package-split.md).
 */
import type { Command } from 'commander';
import type { AuthLadder } from '../infra/auth.ts';
import type { NameSource } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import { fillNames, fillNamesIn } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import { buildLoginSummary } from '../use-cases/commands/login-status.ts';
import * as login from '../use-cases/commands/login.ts';
import * as logout from '../use-cases/commands/logout.ts';
import * as status from '../use-cases/commands/status.ts';
import type { LoginAuthFactory } from './build-deps.ts';
import type { CliOutput } from '@ask-marcel/office-core/composition/cli-shell.ts';

type AuthCliDeps = {
  readonly auth: AuthLadder;
  /**
   * Builds the AuthManager for an interactive `login` run (it may recapture
   * secondary tokens via the browser, unlike the command-path `auth`).
   * Supplied by the composition root (`build-deps.ts`); when omitted (tests),
   * the login action falls back to the injected `auth`, so a single fake
   * drives both `login.execute` and `getLastElevatedOutcome`.
   */
  readonly makeLoginAuth?: LoginAuthFactory;
  /** The package the commands are part of: its bins fill the placeholders of their help and their reports. */
  readonly names: NameSource;
};

/**
 * `auth_cancelled` covers two very different situations and reads as only one
 * of them. Closing the sign-in window is the common case; the other is a
 * browser that launched and could never be driven, because Playwright talks to
 * it over a local DevTools (CDP) connection and an endpoint-security agent can
 * block that pipe. The second looks like a browser stuck on about:blank with no
 * sign-in page, and a bare "Authentication cancelled" sends the reader hunting
 * their own actions instead of their security software (four rounds of it on
 * 2026-09-07, against a SentinelOne policy).
 */
const AUTH_CANCELLED_MESSAGE =
  'Authentication cancelled. If you closed the sign-in window, run `login` again. If a browser opened but stayed on about:blank with no sign-in page, it launched but could not be driven: endpoint-security / EDR software blocking the local DevTools (CDP) connection is the usual cause. If you are running under Bun, try the Node install instead (`npm i -g ask-marcel-office-cli`, then `{auth-bin} login`) — some EDR policies block Bun from driving the browser but trust Node; otherwise add a security exclusion for the browser Playwright launches. Re-run with ASKMARCEL_TRACE=1 to see which browser failed and why.';

const addAuthCommands = (program: Command, output: CliOutput, deps: AuthCliDeps): void => {
  const { auth, names } = deps;
  const fill = (text: string): string => fillNames(text, names);
  const loginCmd = program
    .command('login')
    .description(
      'Authenticate against Microsoft Graph via the Teams web client (cached token → refresh → browser). Already signed in? Reports all four cached tokens (basic / elevated / chatsvcagg / ic3) with their time-left and refresh route; --force re-captures every token via the browser.'
    )
    .option(
      '--force',
      'Ignore the cache and re-capture every token via the browser. Rarely needed, and NOT free: it clears the browser session first so the token grant re-fires, which deletes the 90-day "Stay signed in" cookie and usually means entering your password again. A plain `login` already re-captures the elevated (M365) token when it is missing and leaves that session intact, so reach for `--force` only when a tier is stuck and a plain `login` has not fixed it.'
    )
    .action(async () => {
      const force = loginCmd.opts<{ force?: boolean }>().force ?? false;
      // The composition root supplies `makeLoginAuth` (a login-configured
      // manager that may recapture secondary tokens via the browser); tests
      // omit it and the injected `auth` fake drives both execute and the
      // token-status read.
      const loginAuth = deps.makeLoginAuth ? deps.makeLoginAuth() : auth;
      const result = await login.execute(loginAuth, { force });
      if (!result.ok) {
        output.fail(result.error.type === 'auth_cancelled' ? AUTH_CANCELLED_MESSAGE : result.error.message);
        return;
      }
      // Slim confirmation: which tokens are available now, and where to look next.
      // The decode-only read never opens a browser. Full per-token scopes + expiry
      // live in `status`; login just points there (and to `login --force`).
      const info = await loginAuth.getTokenInfo();
      if (!info.ok) {
        output.fail(info.error.message);
        return;
      }
      const summary = buildLoginSummary({
        elevatedAvailable: info.value.elevated.available,
        chatsvcaggAvailable: info.value.chatsvcagg.available,
        ic3Available: info.value.ic3.available,
      });
      output.print(fillNamesIn(summary, names));
    });
  loginCmd.addHelpText(
    'after',
    fill(
      [
        '',
        'Examples:      {auth-bin} login          (sign in the first time, or show all four token statuses if already signed in)',
        '               {auth-bin} login --force  (re-capture every token via the browser, ignoring the cache)',
        'Token cache:   ~/.ask-marcel/token-cache.json (access + refresh tokens, JSON, 0600).',
        'Browser data:  ~/.ask-marcel/browser-profile/ (Playwright persistent context).',
        'Scopes:        granted by Microsoft to the Teams web client (CLIENT_ID 5e3ce6c0-...);',
        '               this CLI cannot request additional scopes. To inspect the granted set,',
        '               run `{auth-bin} status`.',
        'Stuck flow:    `{auth-bin} logout` then re-run; the browser fallback opens a fresh Edge / Chrome window.',
      ].join('\n  ')
    )
  );

  const statusCmd = program
    .command('status')
    .description(
      'Show the four cached tokens (basic, elevated, chatsvcagg and ic3). For each token, it shows if the token is available and the seconds before its expiry. It also shows the scopes, the refresh method, and the data that the token lets you read. It reads only the token cache, and it does not make a Graph call or open a browser.'
    )
    .action(async () => {
      const result = await status.execute(auth);
      if (result.ok) output.print(fillNamesIn(result.value, names));
      else output.fail(result.error.message, result.error.code, 'cli');
    });
  statusCmd.addHelpText(
    'after',
    fill(['', 'Example:       {auth-bin} status', 'Not signed in: the error code is `not_authenticated`. Run `{auth-bin} login` first.'].join('\n  '))
  );

  const logoutCmd = program
    .command('logout')
    .description('Clear the cached tokens and the browser profile so the next command forces a fresh sign-in.')
    .action(async () => {
      const result = await logout.execute(auth);
      if (result.ok) output.print({ status: 'logged_out' });
      else output.fail(result.error.type === 'auth_cancelled' ? 'Logout cancelled' : result.error.message);
    });
  logoutCmd.addHelpText(
    'after',
    fill(
      [
        '',
        'Example:       {auth-bin} logout',
        'Removes:       ~/.ask-marcel/token-cache.json (access + refresh tokens).',
        '               ~/.ask-marcel/browser-profile/, or $ASKMARCEL_BROWSER_PROFILE when set (Playwright persistent context).',
        'Next sign-in:  usually asks for your password again, since the "Stay signed in" cookie lived in the profile.',
        'Verify clean:  ls ~/.ask-marcel/  (token-cache.json and browser-profile/ should both be gone).',
      ].join('\n  ')
    )
  );
};

export { addAuthCommands };
export type { AuthCliDeps };
