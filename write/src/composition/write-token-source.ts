/*
 * How the write package signs (E3, LESSONS 2026-10-18): a tier's environment
 * variable when it is set, otherwise the token helper of the auth package. The
 * write package holds no token cache, opens no browser and writes no locator:
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
import { commands } from '../use-cases/commands/write-commands.ts';

// The write commands that need each secondary token, in registry order, for the
// tail of a token failure. None needs one today (meta-invariants-write.test.ts),
// but the lists stay derived, so a write command that gains a flag is named.
const namesWhere = (matches: (meta: CommandMeta) => boolean): ReadonlyArray<string> => Object.entries(commands).flatMap(([name, command]) => (matches(command.meta) ? [name] : []));

export const WRITE_NEEDED_BY: SecondaryTokenCommands = {
  elevated: namesWhere((meta) => meta.needsElevatedToken === true),
  chatsvcagg: namesWhere((meta) => meta.needsSubstrateToken === 'chatsvcagg'),
  ic3: namesWhere((meta) => meta.needsSubstrateToken === 'ic3'),
};

// The platform and the runtime are always this process's (core's defaults).
export type WriteTokenSourceOptions = Partial<Omit<EnvThenHelperOptions, 'neededBy' | 'platform' | 'execPath'>>;

const onBun = (): boolean => typeof globalThis.Bun !== 'undefined';

/** The token source of the write package. Each option not given takes this process's value. */
export const buildWriteTokenSource = (options: WriteTokenSourceOptions = {}): TokenSource =>
  createCoreTokenSource({
    env: options.env ?? process.env,
    home: options.home ?? homedir(),
    fs: options.fs ?? (onBun() ? createBunFileSystem() : createNodeFileSystem()),
    runner: options.runner ?? (onBun() ? createBunProcessRunner() : createNodeProcessRunner()),
    interactive: options.interactive ?? process.stdin.isTTY === true,
    neededBy: WRITE_NEEDED_BY,
  });

/**
 * The library's token source: the tier variables of this process, else the
 * token helper of the auth package. It takes no parameter, so the package's
 * types name no file system and no process runner.
 */
export const createEnvThenHelperTokenSource = (): TokenSource => buildWriteTokenSource();
