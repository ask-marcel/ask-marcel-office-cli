/*
 * The names of the single package `ask-marcel-office-cli`: one bin that holds
 * the auth, read and write commands, so every bin placeholder (bin-names.ts)
 * renders as that bin. Its own module, apart from `single-package.ts`. The token
 * helper does not read it: it fills its remedies with the auth names
 * (auth-package-names.ts).
 */
import type { NameSource } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';

const BIN = 'ask-marcel-office';

export const SINGLE_PACKAGE_NAMES: NameSource = { binName: BIN, bins: { auth: BIN, read: BIN, write: BIN } };
