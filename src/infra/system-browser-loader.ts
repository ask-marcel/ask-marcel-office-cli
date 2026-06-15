/*
 * Real system-browser auth loader — production wiring only.
 *
 * This file is the dynamic-import boundary between `auth.ts` and the
 * `system-browser-auth.ts` runtime path (which opens the user's default
 * browser + spins a localhost token-callback server). It exists so that
 * `auth.ts` stays 100% unit-testable: tests inject a fake `SystemBrowserAuthFn`
 * into `createAuthManagerFromApi`, while production builds the default one
 * here.
 *
 * The body cannot be unit-tested without actually launching a browser, so
 * `scripts/check-coverage.ts` SKIPS this file — see its `production-wiring`
 * skip rule (mirrors `playwright-loader.ts`).
 */

import type { SystemBrowserAuthFn } from './auth.ts';
import type { Logger } from '../use-cases/ports/logger.ts';

export const createDefaultSystemBrowserAuth = (logger: Logger, skipSystemBrowser: boolean): SystemBrowserAuthFn => {
  const run: SystemBrowserAuthFn = async () => {
    if (skipSystemBrowser) {
      return { ok: false as const, error: { type: 'skipped' as const, message: 'system browser skipped' } };
    }
    const { authenticateViaSystemBrowser } = await import('./system-browser-auth.ts');
    return authenticateViaSystemBrowser({ logger, extensionTimeoutMs: 10_000 });
  };
  return run;
};
