import { createEnvTokenSource } from '../infra/env-token-source.ts';
import { createHelperTokenSource } from '../infra/helper-token-source.ts';
import { resolveAuthPaths } from '../infra/auth-paths.ts';
import { locateTokenHelper } from '../infra/token-helper-locator.ts';
import { runTokenHelper } from '../infra/token-helper-run.ts';
import type { FileExistence, FileSystem } from '../use-cases/ports/filesystem.ts';
import type { ProcessRunner } from '../use-cases/ports/process-runner.ts';
import type { TokenSource } from '../use-cases/ports/token-source.ts';
import { tokenHelperLocatorPath } from './token-helper.ts';

export type EnvThenHelperOptions = {
  readonly env: Readonly<Record<string, string | undefined>>;
  readonly home: string;
  readonly fs: Pick<FileSystem, 'readJson' | 'readBytes'> & FileExistence;
  readonly runner: ProcessRunner;
  readonly interactive: boolean;
  // This process's platform and runtime unless a test names others.
  readonly platform?: string;
  readonly execPath?: string;
};

// The package split's token source for read and write: a tier's environment
// variable when it is set, otherwise the token helper wherever it is found.
// The single package uses it only when ASKMARCEL_TOKEN_COMMAND is set.
export const createEnvThenHelperTokenSource = (options: EnvThenHelperOptions): TokenSource => {
  const { env, fs, runner, interactive } = options;
  const locatorPath = tokenHelperLocatorPath(resolveAuthPaths(options.home, env));
  const platform = options.platform ?? process.platform;
  const execPath = options.execPath ?? process.execPath;
  const locate = (): ReturnType<typeof locateTokenHelper> => locateTokenHelper({ env, locatorPath, fs, platform, execPath });
  const helper = createHelperTokenSource({ ask: (request) => runTokenHelper({ locate, runner, interactive }, request) });
  return createEnvTokenSource(env, helper);
};
