/*
 * How the read package signs (LESSONS 2026-10-18): a tier's environment
 * variable when it is set, otherwise the token helper of the auth package. The
 * read package holds no token cache, opens no browser and writes no locator:
 * only the auth bin writes them.
 */
import { homedir } from 'node:os';
import type { EnvThenHelperOptions } from '@ask-marcel/office-core/composition/token-source.ts';
import { createEnvThenHelperTokenSource as createCoreTokenSource } from '@ask-marcel/office-core/composition/token-source.ts';
import { createBunFileSystem } from '@ask-marcel/office-core/infra/filesystem-bun.ts';
import { createNodeFileSystem } from '@ask-marcel/office-core/infra/filesystem-node.ts';
import { createBunProcessRunner } from '@ask-marcel/office-core/infra/process-runner-bun.ts';
import { createNodeProcessRunner } from '@ask-marcel/office-core/infra/process-runner-node.ts';
import type { SecondaryTokenCommands } from '@ask-marcel/office-core/infra/token-source-needed-by.ts';
import type { CommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { TokenSource } from '@ask-marcel/office-core/use-cases/ports/token-source.ts';
import { commands } from '../use-cases/commands/index.ts';

// The secondary-token error messages name the read commands that need each
// token, in name order. Deriving the lists from the registry flags keeps the
// messages truthful as commands gain or lose a flag. The token source adds
// them after the remedy of the auth package.
const namesWhere = (matches: (meta: CommandMeta) => boolean): ReadonlyArray<string> =>
  Object.entries(commands)
    .filter(([, command]) => matches(command.meta))
    .map(([name]) => name)
    .toSorted((a, b) => a.localeCompare(b));

export const READ_NEEDED_BY: SecondaryTokenCommands = {
  elevated: namesWhere((meta) => meta.needsElevatedToken === true),
  chatsvcagg: namesWhere((meta) => meta.needsSubstrateToken === 'chatsvcagg'),
  ic3: namesWhere((meta) => meta.needsSubstrateToken === 'ic3'),
};

// The platform and the runtime are always this process's (core's defaults).
export type ReadTokenSourceOptions = Partial<Omit<EnvThenHelperOptions, 'neededBy' | 'platform' | 'execPath'>>;

const onBun = (): boolean => typeof globalThis.Bun !== 'undefined';

/** The token source of the read package. Each option not given takes this process's value. */
export const buildReadTokenSource = (options: ReadTokenSourceOptions = {}): TokenSource =>
  createCoreTokenSource({
    env: options.env ?? process.env,
    home: options.home ?? homedir(),
    fs: options.fs ?? (onBun() ? createBunFileSystem() : createNodeFileSystem()),
    runner: options.runner ?? (onBun() ? createBunProcessRunner() : createNodeProcessRunner()),
    interactive: options.interactive ?? process.stdin.isTTY === true,
    neededBy: READ_NEEDED_BY,
  });

/**
 * The library's token source: the tier variables of this process, else the
 * token helper of the auth package. It takes no parameter, so the package's
 * types name no file system and no process runner.
 */
export const createEnvThenHelperTokenSource = (): TokenSource => buildReadTokenSource();
