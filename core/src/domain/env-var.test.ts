import { describe, expect, it } from 'bun:test';
import { envVar, envVarUnsafe } from './env-var.ts';

describe('envVar brand factory', () => {
  it('accepts a non-empty environment value', () => {
    const result = envVar('LOG_LEVEL', 'info');
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(envVarUnsafe('info'));
  });

  it('rejects an undefined environment value with the missing-env-var error', () => {
    const result = envVar('LOG_LEVEL', undefined);
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('missing_env_var');
      expect(result.error.key).toBe('LOG_LEVEL');
    }
  });

  it('treats an empty string as missing', () => {
    const result = envVar('SECRET', '');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('missing_env_var');
  });
});
