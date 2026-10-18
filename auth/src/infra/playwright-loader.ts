/*
 * Real Playwright loader — production wiring only.
 *
 * This file is the dynamic-import boundary between the rest of the codebase
 * and the `playwright` runtime dependency. It exists so that `browser-auth.ts`
 * itself stays 100% unit-testable: tests pass a fake loader to
 * `createPlaywrightApi`, while production passes `loadPlaywright`.
 *
 * The body is one line, by design; it cannot be unit-tested without
 * actually launching a browser. `scripts/check-coverage.ts` SKIPS this file
 * for that reason — see the `production-wiring` skip rule there.
 */

import type { ChromiumLike } from './browser-auth.ts';

export const loadPlaywright = async (): Promise<{ readonly chromium: ChromiumLike }> => import('playwright');
