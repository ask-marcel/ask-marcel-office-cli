import { PACKAGE_ROLES } from './command-directory.ts';
import type { CommandPackage } from './command-package.ts';

/*
 * The placeholders a command reference carries in meta, hints, remedies and
 * examples (docs/plans/2026-10-01-package-split.md, "Cross-package names"):
 * `{auth-bin}`, `{read-bin}` and `{write-bin}` stand for the bin of the package
 * that holds the command (see command-directory.ts), and `{bin}` for the bin of
 * the package that renders the text, before the commands every package has
 * (`docs`, `help-json`, `mcp`, `<command> --help`). The renderers fill them at
 * render time; in the single package all four are `ask-marcel-office`.
 */

/** What filling the placeholders needs from the package: its bin and how it names the three bins of the split. */
export type NameSource = Pick<CommandPackage, 'binName' | 'bins'>;

export const fillNames = (text: string, names: NameSource): string =>
  PACKAGE_ROLES.reduce((filled, role) => filled.replaceAll(`{${role}-bin}`, names.bins[role]), text.replaceAll('{bin}', names.binName));

/** The text filled, or as it is when no package is named (a presenter call with no package behind it). */
export const fillNamesOrKeep = (text: string, names: NameSource | undefined): string => (names === undefined ? text : fillNames(text, names));

const isPlainRecord = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);

/**
 * Fills every string of a value the CLI built itself (a token report, a login
 * summary). Never for data from Graph: a document or a mail can hold `{bin}`.
 */
export const fillNamesIn = (value: unknown, names: NameSource): unknown => {
  if (typeof value === 'string') return fillNames(value, names);
  if (Array.isArray(value)) return value.map((item) => fillNamesIn(item, names));
  if (!isPlainRecord(value)) return value;
  return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, fillNamesIn(item, names)]));
};
