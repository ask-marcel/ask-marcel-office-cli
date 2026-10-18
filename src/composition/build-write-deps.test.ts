import { afterEach, describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { FetchMock } from '@ask-marcel/office-core/test-helpers/fetch-mock.ts';
import { installFetchMock } from '@ask-marcel/office-core/test-helpers/fetch-mock.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { unsignedJwt } from '@ask-marcel/office-core/test-helpers/jwt.ts';
import { createProcessRunnerFake } from '@ask-marcel/office-core/test-helpers/process-runner-fake.ts';
import { buildWriteDeps } from './build-write-deps.ts';

// The composition root of the write package: a logger, the file system and the
// write graph signed through the tier variables or the token helper. It holds
// no process runner, no token cache and no locator: only the auth bin writes them.

const HOME = '/virtual/home';
const HELPER = '/opt/ask-marcel/auth/dist/token.js';
const NODE = '/usr/local/bin/node';
const inAnHour = (): number => Math.floor(Date.now() / 1000) + 3600;
const TOKEN = unsignedJwt({ aud: 'https://graph.microsoft.com', exp: inAnHour(), from: 'helper' });

let fetches: FetchMock | undefined;
afterEach(() => fetches?.restore());

const captureStderr = (run: () => void): string => {
  const original = process.stderr.write.bind(process.stderr);
  let captured = '';
  process.stderr.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    run();
  } finally {
    process.stderr.write = original;
  }
  return captured;
};

const graphAnswers = (): FetchMock =>
  installFetchMock([{ match: () => true, respond: () => new Response(JSON.stringify({ id: 'draft-1' }), { headers: { 'content-type': 'application/json' } }) }]);

describe('the deps of the write package', () => {
  it('wires a logger, the file system and the write graph, and no process runner', () => {
    const fs = createFileSystemFake();
    const deps = buildWriteDeps({ home: HOME, env: {}, fs, logLevel: 'error' });
    expect(Object.keys(deps).toSorted((a, b) => a.localeCompare(b))).toEqual(['fs', 'graph', 'logger']);
    expect(deps.fs).toBe(fs);
    expect(Object.keys(deps.graph).toSorted((a, b) => a.localeCompare(b))).toEqual(['delete', 'fetchUrl', 'get', 'getBinary', 'patch', 'post', 'put']);
  });

  it('signs a draft with the token that the helper at ASKMARCEL_TOKEN_COMMAND prints, and writes no file', async () => {
    fetches = graphAnswers();
    const fs = createFileSystemFake();
    const runner = createProcessRunnerFake(() => ok({ exitCode: 0, stdout: `${JSON.stringify({ accessToken: TOKEN, expiresOn: inAnHour() })}\n` }));
    const deps = buildWriteDeps({ home: HOME, env: { ASKMARCEL_TOKEN_COMMAND: HELPER }, fs, processRunner: runner, interactive: false });
    expect(await deps.graph.post('/me/messages', { subject: 'Agenda' })).toEqual(ok({ id: 'draft-1' }));
    const headers = fetches.calls[0]?.init?.headers as Record<string, string> | undefined;
    expect(headers?.['Authorization']).toBe(`Bearer ${TOKEN}`);
    expect(runner.calls.map((call) => [call.command, ...call.args, call.options.stdin])).toEqual([[HELPER, '--tier', 'basic', 'ignore']]);
    expect([fs.has(`${HOME}/.ask-marcel/token-helper.json`), fs.has(`${HOME}/.ask-marcel/token-cache.json`)]).toEqual([false, false]);
  });

  it('signs a draft with the helper that the locator under the given home names, read through the given file system', async () => {
    fetches = graphAnswers();
    const fs = createFileSystemFake();
    fs.seed(`${HOME}/.ask-marcel/token-helper.json`, JSON.stringify({ execPath: NODE, entry: HELPER, version: '1.0.0' }));
    fs.seed(HELPER, '');
    fs.seed(NODE, '');
    const runner = createProcessRunnerFake(() => ok({ exitCode: 0, stdout: `${JSON.stringify({ accessToken: TOKEN, expiresOn: inAnHour() })}\n` }));
    const deps = buildWriteDeps({ home: HOME, env: {}, fs, processRunner: runner, interactive: false });
    expect(await deps.graph.post('/me/messages', { subject: 'Agenda' })).toEqual(ok({ id: 'draft-1' }));
    expect(runner.calls.map((call) => [call.command, ...call.args])).toEqual([[NODE, HELPER, '--tier', 'basic']]);
  });

  it('takes the log level from ASKMARCEL_LOG_LEVEL of the env that it is given', () => {
    const deps = buildWriteDeps({ home: HOME, env: { ASKMARCEL_LOG_LEVEL: 'info' }, fs: createFileSystemFake() });
    expect(captureStderr(() => deps.logger.info('write_deps_info_line'))).toContain('write_deps_info_line');
  });

  it('logs errors only when it is given no level', () => {
    const deps = buildWriteDeps({ home: HOME, env: {}, fs: createFileSystemFake() });
    const stderr = captureStderr(() => {
      deps.logger.warn('write_deps_warning_line');
      deps.logger.error('write_deps_error_line');
    });
    expect(stderr).not.toContain('write_deps_warning_line');
    expect(stderr).toContain('write_deps_error_line');
  });

  it('builds from the values of this process when it is given no config', () => {
    expect(Object.keys(buildWriteDeps()).toSorted((a, b) => a.localeCompare(b))).toEqual(['fs', 'graph', 'logger']);
  });
});
