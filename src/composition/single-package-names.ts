/*
 * The names of the single package `ask-marcel-office-cli`: one bin that holds
 * the read commands, so the read bin placeholder (bin-names.ts) renders as that
 * bin. The auth and write placeholders render as the auth and write bins
 * (`PACKAGE_BINS.auth`, `PACKAGE_BINS.write`), because the auth and write
 * commands live in their own packages. Its own module, apart from
 * `single-package.ts`. The token helper does not read it: it fills its remedies
 * with the auth names (auth-package-names.ts).
 */
import type { NameSource } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import { PACKAGE_BINS } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';

const BIN = 'ask-marcel-office';

export const SINGLE_PACKAGE_NAMES: NameSource = { binName: BIN, bins: { auth: PACKAGE_BINS.auth, read: BIN, write: PACKAGE_BINS.write } };
