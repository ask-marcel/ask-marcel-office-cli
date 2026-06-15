import { exec } from 'node:child_process';
import { promisify } from 'node:util';
import type { AccessToken } from '../domain/access-token.ts';
import { accessToken, accessTokenUnsafe } from '../domain/access-token.ts';
import { isTokenFresh } from '../domain/jwt-utils.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { Logger } from '../use-cases/ports/logger.ts';
import { createTokenCallbackServer } from './token-callback-server.ts';

const execAsync = promisify(exec);

type SystemBrowserAuthDeps = {
  readonly logger: Logger;
  readonly timeoutMs?: number;
  readonly extensionTimeoutMs?: number;
};

type SystemBrowserTokenResult = {
  readonly accessToken: AccessToken;
  readonly refreshToken: string | null;
  readonly elevatedAccessToken?: AccessToken | null;
  readonly chatsvcaggAccessToken?: AccessToken | null;
  readonly ic3AccessToken?: AccessToken | null;
  readonly chatsvcaggRegion?: string;
};

type SystemBrowserAuthError =
  | { type: 'server_bind_failed'; message: string }
  | { type: 'browser_open_failed'; message: string }
  | { type: 'extension_timeout'; message: string }
  | { type: 'invalid_token'; message: string }
  | { type: 'cancelled'; message: string };

const TEAMS_URL = 'https://teams.microsoft.com/';

const CHROME_BUNDLE_ID = 'com.google.chrome';
const EDGE_BUNDLE_ID = 'com.microsoft.edgemac';

const detectDefaultBrowserMac = async (): Promise<string | null> => {
  try {
    const { stdout } = await execAsync('defaults read com.apple.LaunchServices/com.apple.launchservices.secure LSHandlers 2>/dev/null');
    // Each dict block has LSHandlerRoleAll before LSHandlerURLScheme.
    // Track the current block's role, then return it when we see the https scheme.
    let currentRole: string | null = null;
    for (const line of stdout.split('\n')) {
      const roleMatch = line.match(/LSHandlerRoleAll\s*=\s*"([^"]+)"/);
      if (roleMatch) {
        currentRole = roleMatch[1];
      }
      if (line.includes('LSHandlerURLScheme') && line.includes('https') && currentRole && currentRole !== '-') {
        return currentRole;
      }
      // Reset on dict boundary
      if (line.trim() === '},') {
        currentRole = null;
      }
    }
  } catch {
    // Ignore detection failures
  }
  return null;
};

const openBrowser = async (url: string): Promise<void> => {
  const platform = process.platform;
  let command: string;

  if (platform === 'darwin') {
    const defaultBrowser = await detectDefaultBrowserMac();
    if (defaultBrowser === CHROME_BUNDLE_ID) {
      // Direct executable invocation: open -a doesn't pass args to running instances
      command = `"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --incognito "${url}" 2>/dev/null || open -a "Google Chrome" --args --incognito "${url}"`;
    } else if (defaultBrowser === EDGE_BUNDLE_ID) {
      command = `"/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge" --inprivate "${url}" 2>/dev/null || open -a "Microsoft Edge" --args --inprivate "${url}"`;
    } else {
      // Default browser is Safari or unknown: try Chrome incognito, then Edge InPrivate, then plain open
      command = `"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --incognito "${url}" 2>/dev/null || "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge" --inprivate "${url}" 2>/dev/null || open "${url}"`;
    }
  } else if (platform === 'win32') {
    // Windows: use PowerShell to open default browser with incognito
    // Try Edge first (most common on Windows), then Chrome, then default
    command = `powershell -Command "Start-Process 'msedge' -ArgumentList '-inprivate','${url}' -ErrorAction SilentlyContinue" 2>nul || powershell -Command "Start-Process 'chrome' -ArgumentList '--incognito','${url}' -ErrorAction SilentlyContinue" 2>nul || start "" "${url}"`;
  } else {
    // Linux: Chrome incognito
    command = `google-chrome --incognito "${url}" 2>/dev/null || chromium --incognito "${url}" 2>/dev/null || xdg-open "${url}"`;
  }

  await execAsync(command);
};

const validateToken = (raw: string): Result<AccessToken, SystemBrowserAuthError> => {
  const validated = accessToken(raw);
  if (!validated.ok) {
    return err({ type: 'invalid_token', message: 'extension returned an invalid access token' });
  }
  return ok(validated.value);
};

// Substrate tokens (chatsvcagg, IC3) carry non-Graph audiences
// (e.g. https://chatsvcagg.teams.microsoft.com, https://ic3.teams.office.com).
// The standard accessToken() validator rejects them via isGraphToken().
// This validator only checks JWT format + expiry.
const validateSubstrateToken = (raw: string): Result<AccessToken, SystemBrowserAuthError> => {
  if (!raw.startsWith('eyJ')) {
    return err({ type: 'invalid_token', message: 'substrate token is not a valid JWT' });
  }
  if (!isTokenFresh(raw)) {
    return err({ type: 'invalid_token', message: 'substrate token is expired' });
  }
  return ok(accessTokenUnsafe(raw));
};

const authenticateViaSystemBrowser = async (deps: SystemBrowserAuthDeps): Promise<Result<SystemBrowserTokenResult, SystemBrowserAuthError>> => {
  const { logger, timeoutMs = 5 * 60 * 1000, extensionTimeoutMs = 5 * 60 * 1000 } = deps;

  logger.info('system_browser.starting', { timeoutMs, extensionTimeoutMs });

  // 1. Start the token callback server (OS assigns random port)
  const callbackServer = createTokenCallbackServer(logger, timeoutMs);

  // We need to start the server first to get the port
  const startPromise = callbackServer.start();

  // Wait a bit for the server to bind
  await new Promise((resolve) => setTimeout(resolve, 100));

  const port = callbackServer.port;
  if (port === 0) {
    const result = await startPromise;
    if (!result.ok) {
      return err({ type: 'server_bind_failed', message: result.error.message });
    }
  }

  const actualPort = callbackServer.port;
  logger.info('system_browser.server_ready', { port: actualPort });

  // 2. Open system browser with the port in URL
  const teamsUrl = `${TEAMS_URL}?ask_marcel_port=${actualPort}`;
  logger.info('system_browser.opening', { url: teamsUrl });
  
  // Show progress to user
  process.stderr.write('Opening browser for authentication...\n');
  process.stderr.write('Please ensure the Ask Marcel Companion extension is installed in your browser.\n');
  process.stderr.write('Waiting for token from browser extension (this may take a few minutes)...\n');

  try {
    await openBrowser(teamsUrl);
    logger.info('system_browser.opened');
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    logger.info('system_browser.open_failed', { message });
    await callbackServer.stop();
    return err({ type: 'browser_open_failed', message });
  }

  // 3. Wait for token callback from extension
  logger.info('system_browser.waiting_for_token', { extensionTimeoutMs });

  // Wait for the full timeout period (user needs time to login)
  const result = await startPromise;

  if (!result.ok) {
    // Extension timeout or server error
    logger.info('system_browser.failed', { reason: result.error.type });
    if (result.error.type === 'timeout') {
      const message = 'Browser extension did not respond within timeout. ' +
        'Make sure the Ask Marcel Companion extension is installed and enabled in your browser. ' +
        'See browser-extension/README.md for installation instructions. ' +
        'Alternatively, use --use-playwright flag for Playwright-based authentication.';
      return err({ type: 'extension_timeout', message });
    }
    return err({ type: 'server_bind_failed', message: result.error.message });
  }

  const payload = result.value;

  // 4. Validate tokens
  const accessResult = validateToken(payload.access_token);
  if (!accessResult.ok) return accessResult;

  const response: SystemBrowserTokenResult = {
    accessToken: accessResult.value,
    refreshToken: payload.refresh_token ?? null,
  };

  // Validate optional elevated tokens
  if (payload.elevated_access_token) {
    const elevatedResult = validateToken(payload.elevated_access_token);
    if (elevatedResult.ok) {
      (response as { elevatedAccessToken?: AccessToken }).elevatedAccessToken = elevatedResult.value;
    }
  }

  if (payload.chatsvcagg_access_token) {
    const chatsvcaggResult = validateSubstrateToken(payload.chatsvcagg_access_token);
    if (chatsvcaggResult.ok) {
      (response as { chatsvcaggAccessToken?: AccessToken }).chatsvcaggAccessToken = chatsvcaggResult.value;
    }
  }

  if (payload.ic3_access_token) {
    const ic3Result = validateSubstrateToken(payload.ic3_access_token);
    if (ic3Result.ok) {
      (response as { ic3AccessToken?: AccessToken }).ic3AccessToken = ic3Result.value;
    }
  }

  if (payload.chatsvcagg_region) {
    (response as { chatsvcaggRegion?: string }).chatsvcaggRegion = payload.chatsvcagg_region;
  }

  logger.info('system_browser.success', {
    hasAccessToken: true,
    hasRefreshToken: !!response.refreshToken,
    hasElevated: !!response.elevatedAccessToken,
    hasChatsvcagg: !!response.chatsvcaggAccessToken,
  });

  return ok(response);
};

export { authenticateViaSystemBrowser };
export type { SystemBrowserAuthDeps, SystemBrowserAuthError, SystemBrowserTokenResult };
