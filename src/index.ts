/*
 * Public library API for `ask-marcel-office-cli`.
 *
 * Anything re-exported here is part of the package's public surface and
 * follows semver. Anything NOT re-exported here is implementation detail
 * (presenter, individual command files, format-error, ports, fakes).
 *
 * Consumers:
 *   import { commands, createGraphClient, buildDeps } from 'ask-marcel-office-cli';
 */

export { accessToken, accessTokenUnsafe } from '@ask-marcel/office-core/domain/access-token.ts';
export type { AccessToken, AccessTokenError } from '@ask-marcel/office-core/domain/access-token.ts';
export { envVar, envVarUnsafe } from '@ask-marcel/office-core/domain/env-var.ts';
export type { EnvVar, EnvVarError } from '@ask-marcel/office-core/domain/env-var.ts';
export { decodeJwtPayload, isGraphToken, isTokenFresh } from '@ask-marcel/office-core/domain/jwt-utils.ts';
export { err, ok } from '@ask-marcel/office-core/domain/result.ts';
export type { Result } from '@ask-marcel/office-core/domain/result.ts';

export { createAuthManager } from './infra/auth-browser.ts';
export type { AuthError, AuthManager } from './infra/auth.ts';
export { createBunFileSystem } from '@ask-marcel/office-core/infra/filesystem-bun.ts';
export { createNodeFileSystem } from '@ask-marcel/office-core/infra/filesystem-node.ts';
export { createGraphClient } from './infra/graph-client.ts';
export type { FetchFn, GraphClient, GraphError } from './infra/graph-client.ts';
export { createWinstonLogger } from '@ask-marcel/office-core/infra/logger.ts';
export type { WinstonLoggerConfig } from '@ask-marcel/office-core/infra/logger.ts';
export { createBunProcessRunner } from '@ask-marcel/office-core/infra/process-runner-bun.ts';
export { createNodeProcessRunner } from '@ask-marcel/office-core/infra/process-runner-node.ts';

export type { AtomicFileWrites, FileExistence, FileSystem, FileSystemError } from '@ask-marcel/office-core/use-cases/ports/filesystem.ts';
export type { Logger, LogMeta } from '@ask-marcel/office-core/use-cases/ports/logger.ts';
export type { ProcessRunner, ProcessRunnerError, ProcessRunOptions, ProcessRunResult } from '@ask-marcel/office-core/use-cases/ports/process-runner.ts';

export { commands } from './use-cases/commands/index.ts';
export type { Command, CommandExecute, CommandSchema } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';

export { buildDeps } from './composition/build-deps.ts';
export type { BuildDepsConfig, BuiltDeps } from './composition/build-deps.ts';
