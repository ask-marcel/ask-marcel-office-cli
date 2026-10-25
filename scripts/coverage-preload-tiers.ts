/*
 * Coverage preload for a workspace folder (core/, auth/, read/, write/).
 *
 * `bun test --coverage` reports only the files that the tests import, so an
 * untested source file is absent from the table and the per-tier gate passes
 * over it. This preload imports every source file of the five tiers, so each
 * one shows in the table, at 0% when no test reaches it.
 *
 * It finds the files itself: a Bun.Glob over src/{domain,use-cases,infra,
 * presenter,composition} of the folder that runs it (the working directory),
 * minus the tests and the type-only ports. It keeps no list of files by hand:
 * a list like that misses the files nobody added to it.
 *
 * Shared at the root and run with the folder as the working directory:
 * `bun run ../scripts/check-coverage.ts --preload ../scripts/coverage-preload-tiers.ts`.
 * scripts/coverage-preload-tiers-selftest.ts proves that it shows an untested file.
 */

import { join } from 'node:path';

const TIERS = ['domain', 'use-cases', 'infra', 'presenter', 'composition'];

const sources = new Bun.Glob(`src/{${TIERS.join(',')}}/**/*.ts`);

for (const path of sources.scanSync({ cwd: process.cwd() })) {
  if (path.endsWith('.test.ts') || path.includes('/ports/')) continue;
  await import(join(process.cwd(), path));
}
