/*
 * The names of the read package `@ask-marcel/office-read`: its bin renders
 * `{bin}` and `{read-bin}`, and the bins of auth and write render theirs
 * (bin-names.ts), because the auth and write commands live in their own
 * packages. Its own module, apart from `single-package.ts`. The token helper
 * does not read it: it fills its remedies with the auth names
 * (auth-package-names.ts).
 */
import type { NameSource } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import { PACKAGE_BINS } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';

export const SINGLE_PACKAGE_NAMES: NameSource = { binName: PACKAGE_BINS.read, bins: PACKAGE_BINS };
