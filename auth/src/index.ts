/*
 * Public library API of @ask-marcel/office-auth (dist/index.js and
 * dist/index.d.ts). In phase 4 it gives the cache-path resolver: the folder of
 * the token cache and the browser profile under one home, the same paths the
 * auth bin uses (D3). The token functions come with the auth library step.
 *
 * Core is private and never published: the build inlines its code, and its
 * types reach the published declarations only through re-exports such as the
 * ones below (docs/plans/2026-10-01-package-split.md, Library API per package).
 */

export type { Result } from '@ask-marcel/office-core/domain/result.ts';
export { resolveAuthPaths } from '@ask-marcel/office-core/infra/auth-paths.ts';
export type { AuthPaths } from '@ask-marcel/office-core/infra/auth-paths.ts';
