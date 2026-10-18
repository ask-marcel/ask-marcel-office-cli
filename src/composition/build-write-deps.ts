/*
 * The composition root of the write package: a logger, the file system and the
 * write graph. The graph signs through the tier variables, else the token
 * helper of the auth package (`write-token-source.ts`). The deps hold no
 * process runner, and nothing here writes the token cache or the locator: only
 * the auth bin writes them (LESSONS 2026-10-18).
 */
import { createBunFileSystem } from '@ask-marcel/office-core/infra/filesystem-bun.ts';
import { createNodeFileSystem } from '@ask-marcel/office-core/infra/filesystem-node.ts';
import { createWinstonLogger } from '@ask-marcel/office-core/infra/logger.ts';
import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import { createWriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import type { AtomicFileWrites, FileExistence, FileSystem } from '@ask-marcel/office-core/use-cases/ports/filesystem.ts';
import type { Logger } from '@ask-marcel/office-core/use-cases/ports/logger.ts';
import type { ProcessRunner } from '@ask-marcel/office-core/use-cases/ports/process-runner.ts';
import { buildWriteTokenSource } from './write-token-source.ts';

export type BuildWriteDepsConfig = {
  readonly logLevel?: string;
  readonly home?: string;
  readonly env?: Readonly<Record<string, string | undefined>>;
  readonly fs?: FileSystem & AtomicFileWrites & FileExistence;
  // The runner of the token helper. It stays inside the token source.
  readonly processRunner?: ProcessRunner;
  // Whether a person is at a terminal; the helper then gets its stdin. Defaults to `process.stdin.isTTY`.
  readonly interactive?: boolean;
};

export type BuiltWriteDeps = Readonly<{ logger: Logger; graph: WriteGraph; fs: FileSystem & AtomicFileWrites }>;

const defaultFileSystem = (): FileSystem & AtomicFileWrites & FileExistence => (typeof globalThis.Bun === 'undefined' ? createNodeFileSystem() : createBunFileSystem());

export const buildWriteDeps = (config: BuildWriteDepsConfig = {}): BuiltWriteDeps => {
  const env = config.env ?? process.env;
  const fs = config.fs ?? defaultFileSystem();
  const logger = createWinstonLogger({ logLevel: config.logLevel ?? env.ASKMARCEL_LOG_LEVEL ?? 'error' });
  // A value not given stays undefined, and the token source takes this process's.
  const tokens = buildWriteTokenSource({ env, fs, home: config.home, runner: config.processRunner, interactive: config.interactive });
  return { logger, fs, graph: createWriteGraph(tokens) };
};
