/*
 * The names of the auth package `@ask-marcel/office-auth`: its bin renders
 * `{bin}` and `{auth-bin}`, and the bins of read and write render theirs
 * (bin-names.ts). Its own module, so the token helper entry (src/token.ts)
 * fills its remedies with these names and loads no registry.
 */
import type { NameSource } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import { PACKAGE_BINS } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';

export const AUTH_PACKAGE_NAMES: NameSource = { binName: PACKAGE_BINS.auth, bins: PACKAGE_BINS };
