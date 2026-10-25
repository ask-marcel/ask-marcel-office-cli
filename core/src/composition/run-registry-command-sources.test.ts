import { describe, expect, it } from 'bun:test';
import { err } from '../domain/result.ts';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import type { GraphError } from '../infra/graph-request.ts';
import { runRegistryCommand } from './run-registry-command.ts';

const deps = { graph: fakeGraphClient(), fs: createFileSystemFake(), commandPackage: fakeCommandPackage() };

const failWith = async (error: GraphError): Promise<unknown> => {
  const command = { ...fakeCommand(), execute: async () => err(error) };
  return runRegistryCommand(deps, { name: 'list-widgets', command, params: {}, surface: 'cli' });
};

describe('the source of a failed Graph command', () => {
  it('names Graph as the source when Graph throttles the call, and keeps the wait it asked for', async () => {
    const result = await failWith({ type: 'api_error', status: 429, message: 'Too many requests', code: 'TooManyRequests', retryAfterSeconds: 7 });
    expect(result).toEqual({ ok: false, error: { message: 'Too many requests', code: 'TooManyRequests', source: 'graph', retryAfterSeconds: 7 } });
  });

  it('names the CLI as the source when the sign-in is gone, because the remedy is to sign in again', async () => {
    const result = await failWith({ type: 'auth_failed', message: 'Sign in again', code: 'no_token' });
    expect(result).toEqual({ ok: false, error: { message: 'Sign in again', code: 'no_token', source: 'cli' } });
  });
});
