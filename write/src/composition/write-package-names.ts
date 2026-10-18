/*
 * The names of the write package `@ask-marcel/office-write`: its bin renders
 * `{bin}` and `{write-bin}`, and the bins of auth and read render theirs
 * (bin-names.ts). Its own module, apart from `write-package.ts`, as auth's.
 */
import type { NameSource } from '@ask-marcel/office-core/use-cases/commands/bin-names.ts';
import { PACKAGE_BINS } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';

export const WRITE_PACKAGE_NAMES: NameSource = { binName: PACKAGE_BINS.write, bins: PACKAGE_BINS };
