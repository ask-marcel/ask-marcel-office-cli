import type { Result } from './result.ts';
import { err, ok } from './result.ts';

export type EnvVar = string & { readonly __brand: 'EnvVar' };

export type EnvVarError = { type: 'missing_env_var'; key: string };

export const envVar = (key: string, raw: string | undefined): Result<EnvVar, EnvVarError> => {
  if (!raw) return err({ type: 'missing_env_var', key });
  return ok(raw as EnvVar);
};

export const envVarUnsafe = (raw: string): EnvVar => raw as EnvVar;
