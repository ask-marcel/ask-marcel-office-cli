/*
 * The names of the single package `ask-marcel-office-cli`: one bin that holds
 * the auth, read and write commands, so every bin placeholder (bin-names.ts)
 * renders as that bin. Its own module, apart from `single-package.ts`, because
 * the token helper entry (src/token.ts) fills its remedies with these names and
 * must not load the registry.
 */
import type { NameSource } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';

const BIN = 'ask-marcel-office';

export const SINGLE_PACKAGE_NAMES: NameSource = { binName: BIN, bins: { auth: BIN, read: BIN, write: BIN } };
