/*
 * Public library API of @ask-marcel/office-auth (dist/index.js and
 * dist/index.d.ts), still a skeleton of the package split until its code moves
 * in (phase 4).
 *
 * Core is private and never published: the build inlines its code, and its
 * types reach the published declarations only through re-exports such as the
 * one below (docs/plans/2026-10-01-package-split.md, Library API per package).
 */

export type { Result } from '@ask-marcel/office-core/domain/result.ts';
