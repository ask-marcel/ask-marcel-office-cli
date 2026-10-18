/*
 * The names of the single package `ask-marcel-office-cli`: one bin that holds
 * the read and write commands, so the read and write bin placeholders
 * (bin-names.ts) render as that bin. The auth placeholder renders as the auth
 * bin (`PACKAGE_BINS.auth`), because the auth commands live in the auth
 * package. Its own module, apart from `single-package.ts`. The token helper
 * does not read it: it fills its remedies with the auth names
 * (auth-package-names.ts).
 */
import type { NameSource } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import { PACKAGE_BINS } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';

const BIN = 'ask-marcel-office';

export const SINGLE_PACKAGE_NAMES: NameSource = { binName: BIN, bins: { auth: PACKAGE_BINS.auth, read: BIN, write: BIN } };
