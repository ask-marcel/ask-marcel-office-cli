import { didYouMean } from '../../domain/closest-names.ts';
import type { PackageRole } from './command-directory.ts';
import { COMMAND_DIRECTORY, commandRole, PACKAGE_NAMES } from './command-directory.ts';
import type { CommandPackage } from './command-package.ts';

/*
 * The answer to a command name a package does not have, read off the command
 * directory (docs/plans/2026-10-01-package-split.md, "Cross-package names"): a
 * removed name gets what replaced it, a command of another package gets that
 * package and how to reach it there, anything else the closest names. A command
 * lives elsewhere only when its package has a bin other than this one, so the
 * single package, whose one bin holds all three, sends nobody elsewhere.
 *
 * A terminal reaches a command through its bin; an MCP client through the server
 * of its package. The auth server has only the `login` and `status` tools, so an
 * MCP caller is never sent there for `logout` or `token`, and nobody is offered
 * `token`: read and write run the token helper themselves.
 */

/** What the answer reads from the package: its bins, and its npm name for the install command. */
export type CommandPlace = Pick<CommandPackage, 'binName' | 'bins' | 'packageName'>;

/** Which front end asked: a terminal runs a bin, an MCP client registers a server. */
export type CommandSurface = 'cli' | 'mcp';

// A removed command that another one replaced, and that one's package.
const REPLACED_COMMANDS: ReadonlyMap<string, { readonly name: string; readonly role: PackageRole }> = new Map([['scopes-check', { name: 'status', role: 'auth' }]]);

const REMOVED_UPDATE = 'update';

const TOKEN_HELPER = 'token';

// The auth commands the auth MCP server serves as tools (auth-mcp.ts).
const AUTH_TOOLS: ReadonlySet<string> = new Set(['login', 'status']);

const isElsewhere = (role: PackageRole, place: CommandPlace): boolean => place.bins[role] !== place.binName;

const isMcpTool = (name: string, role: PackageRole): boolean => role !== 'auth' || AUTH_TOOLS.has(name);

/** A command as the caller reaches it: bare in this package, behind its bin or on its server in another. */
const reference = (name: string, role: PackageRole, place: CommandPlace, surface: CommandSurface): string => {
  if (!isElsewhere(role, place)) return `\`${name}\``;
  if (surface === 'cli') return `\`${place.bins[role]} ${name}\``;
  return role === 'auth' ? `the \`${name}\` tool of the ${place.bins[role]} server` : `\`${name}\` on the ${place.bins[role]} server`;
};

const livesIn = (name: string, role: PackageRole, place: CommandPlace, surface: CommandSurface): string => {
  const bin = place.bins[role];
  if (name === TOKEN_HELPER) return ` \`${name}\` is the token helper of ${PACKAGE_NAMES[role]}: the read and write packages run \`${bin} ${name}\` themselves.`;
  if (surface === 'cli') return ` \`${name}\` lives in ${PACKAGE_NAMES[role]}: install it and run \`${bin} ${name}\`.`;
  if (!isMcpTool(name, role)) return ` \`${name}\` lives in ${PACKAGE_NAMES[role]} and has no MCP tool: run \`${bin} ${name}\` in a terminal.`;
  const call = role === 'auth' ? `call its \`${name}\` tool there` : 'call it there';
  return ` \`${name}\` lives in ${PACKAGE_NAMES[role]}: register its MCP server (\`${bin} mcp\`) and ${call}.`;
};

const elsewhereRole = (name: string, place: CommandPlace): PackageRole | undefined => {
  const role = commandRole(name);
  return role !== undefined && isElsewhere(role, place) ? role : undefined;
};

// A command of another package worth offering for a near miss: one the caller can reach from this surface.
const isOffered = (name: string, role: PackageRole, place: CommandPlace, surface: CommandSurface): boolean =>
  name !== TOKEN_HELPER && isElsewhere(role, place) && (surface === 'cli' || isMcpTool(name, role));

// The closest names among this package's commands and the directory's commands of other packages.
const closest = (name: string, available: ReadonlyArray<string>, place: CommandPlace, surface: CommandSurface): string => {
  const elsewhere = [...COMMAND_DIRECTORY].filter(([candidate, role]) => isOffered(candidate, role, place, surface)).map(([candidate]) => candidate);
  return didYouMean(name, [...available, ...elsewhere], 'Did you mean', (candidate) => {
    const role = elsewhereRole(candidate, place);
    return role === undefined ? `\`${candidate}\`` : reference(candidate, role, place, surface);
  });
};

const replaced = (name: string, replacement: { readonly name: string; readonly role: PackageRole }, place: CommandPlace, surface: CommandSurface): string => {
  const by = reference(replacement.name, replacement.role, place, surface);
  return ` ${by.charAt(0).toUpperCase()}${by.slice(1)} replaced \`${name}\`.`;
};

/** What follows "unknown command": led by a space, or '' when nothing is close. */
export const commandSuggestion = (name: string, available: ReadonlyArray<string>, place: CommandPlace, surface: CommandSurface): string => {
  const replacement = REPLACED_COMMANDS.get(name);
  if (replacement !== undefined) return replaced(name, replacement, place, surface);
  if (name === REMOVED_UPDATE)
    return ` \`${name}\` was removed: to upgrade, install the package again with \`npm i -g ${place.packageName}@latest\` (or \`bun add -g ${place.packageName}@latest\`).`;
  const role = elsewhereRole(name, place);
  return role === undefined ? closest(name, available, place, surface) : livesIn(name, role, place, surface);
};

/** `Unknown command "x". <suggestion> <remedy>`, the one sentence every front end answers an unknown name with. */
export const unknownCommandMessage = (name: string, available: ReadonlyArray<string>, remedy: string, place: CommandPlace, surface: CommandSurface): string =>
  `Unknown command "${name}".${commandSuggestion(name, available, place, surface)} ${remedy}`;
