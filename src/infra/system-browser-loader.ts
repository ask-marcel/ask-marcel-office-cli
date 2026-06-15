import type { Logger } from '../use-cases/ports/logger.ts';
import type { SystemBrowserAuthFn } from './auth.ts';

// production-wiring: the default system-browser auth function. Its body
// dynamically imports the system-browser-auth flow, which spins up a real
// localhost callback server and opens the user's browser — impossible to
// unit-test without those side effects. Coverage-skipped for the same reason
// as playwright-loader.ts; tests inject a fake systemBrowserAuthFn instead, so
// this default is never exercised under test (the `??` short-circuits past it).
export const defaultSystemBrowserAuth =
  (logger: Logger, skipSystemBrowser: boolean): SystemBrowserAuthFn =>
  async () => {
    if (skipSystemBrowser) {
      return { ok: false as const, error: { type: 'skipped' as const, message: 'system browser skipped' } };
    }
    const { authenticateViaSystemBrowser } = await import('./system-browser-auth.ts');
    return authenticateViaSystemBrowser({ logger, extensionTimeoutMs: 10_000 });
  };
