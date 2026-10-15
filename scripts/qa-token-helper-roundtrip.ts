#!/usr/bin/env bun
/*
 * The token helper round trip, as read and write will run it (package split,
 * Checks and CI: the windows-latest job). Offline: a fixture token cache, no
 * browser, no network.
 *
 *   1. Builds the helper's entry (src/token.ts) to a temporary dist/token.js.
 *   2. Writes a token cache with a fresh basic token under a temporary HOME
 *      (and USERPROFILE, which Windows reads first).
 *   3. Runs `node dist/token.js --tier basic` with stdin ignored, as auth runs
 *      when a person is not there. It must print the cached token, and it
 *      writes the locator file (~/.ask-marcel/token-helper.json) on the way.
 *   4. Asks the package split's token source for the basic token, through the
 *      locator file, the way read and write do: the env-then-helper source
 *      (src/composition/token-source.ts) spawns `execPath entry --tier basic`
 *      with no shell. Once with the Node process runner (an npm install run by
 *      node) and once with the Bun one. The helper inherits the caller's
 *      environment, so this part runs in a child of this script whose HOME is
 *      the temporary one (`--caller <home>`).
 *
 * Every ASKMARCEL_* variable is removed from the environment of the helper and
 * of the caller: ASKMARCEL_TOKEN_BASIC would answer before the helper runs,
 * and ASKMARCEL_TOKEN_COMMAND would name another helper than the locator file.
 *
 * On Windows this is the path that npm `.cmd` shims break: the locator names
 * node.exe and the script, so the spawn needs neither a shell nor PATH.
 *
 * Run: bun scripts/qa-token-helper-roundtrip.ts (node must be on PATH).
 * Exit non-zero on any failure.
 */

import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { createEnvThenHelperTokenSource } from '../src/composition/token-source.ts';
import { createNodeFileSystem } from '../src/infra/filesystem-node.ts';
import { createBunProcessRunner } from '../src/infra/process-runner-bun.ts';
import { createNodeProcessRunner } from '../src/infra/process-runner-node.ts';
import type { ProcessRunner } from '../src/use-cases/ports/process-runner.ts';

const segment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const inAnHour = Math.floor(Date.now() / 1000) + 3600;
const FIXTURE_TOKEN = `${segment({ alg: 'RS256' })}.${segment({ aud: 'https://graph.microsoft.com', exp: inAnHour, name: 'fixture' })}.sig`;

const fail = (message: string): never => {
  throw new Error(message);
};

const callerAt = process.argv.indexOf('--caller');
const work = callerAt === -1 ? mkdtempSync(join(tmpdir(), 'token-roundtrip-')) : '';
const home = join(work, 'home');
const cleanEnv = Object.fromEntries(Object.entries(process.env).filter(([key]) => !key.startsWith('ASKMARCEL_')));
const env = { ...cleanEnv, HOME: home, USERPROFILE: home };

const buildHelper = async (): Promise<string> => {
  const built = await Bun.build({
    entrypoints: [resolve(import.meta.dir, '../src/token.ts')],
    outdir: join(work, 'dist'),
    naming: 'token.js',
    target: 'node',
    format: 'esm',
    external: ['playwright'],
  });
  if (!built.success) fail(`the helper does not build: ${built.logs.map((log) => String(log.message)).join('; ')}`);
  return join(work, 'dist', 'token.js');
};

const writeFixtureCache = (): void => {
  mkdirSync(join(home, '.ask-marcel'), { recursive: true });
  writeFileSync(join(home, '.ask-marcel', 'token-cache.json'), JSON.stringify({ access_token: FIXTURE_TOKEN, expires_on: inAnHour, refresh_token: 'fixture-refresh-token' }));
};

const runDirectly = (entry: string): void => {
  const proc = Bun.spawnSync(['node', entry, '--tier', 'basic'], { env, stdin: 'ignore', stdout: 'pipe', stderr: 'pipe' });
  const line = proc.stdout.toString();
  if (proc.exitCode !== 0) fail(`node dist/token.js --tier basic exited ${proc.exitCode}: ${proc.stderr.toString().slice(0, 300)}`);
  if (!line.includes(FIXTURE_TOKEN)) fail('node dist/token.js --tier basic did not print the cached token');
};

const roundTrip = async (name: string, runner: ProcessRunner, callerHome: string): Promise<void> => {
  const cached = (await Bun.file(join(callerHome, '.ask-marcel', 'token-cache.json')).json()) as { readonly access_token: string };
  const source = createEnvThenHelperTokenSource({ env: cleanEnv, home: callerHome, fs: createNodeFileSystem(), runner, interactive: false });
  const token = await source.graphToken('basic');
  if (!token.ok) fail(`${name}: the token source failed: ${JSON.stringify(token.error)}`);
  if (token.ok && token.value !== cached.access_token) fail(`${name}: the token source answered another token`);
  process.stdout.write(`qa-token-helper-roundtrip: ${name} runner: the helper answered the cached basic token through the locator file\n`);
};

const asCaller = async (callerHome: string): Promise<void> => {
  await roundTrip('node', createNodeProcessRunner(), callerHome);
  await roundTrip('bun', createBunProcessRunner(), callerHome);
};

const runCaller = (): void => {
  const proc = Bun.spawnSync([process.execPath, import.meta.path, '--caller', home], { env, stdin: 'ignore', stdout: 'inherit', stderr: 'inherit' });
  if (proc.exitCode !== 0) fail('the caller could not get the basic token from the helper');
};

try {
  if (callerAt !== -1) {
    await asCaller(process.argv[callerAt + 1] ?? fail('--caller needs the home'));
  } else {
    const entry = await buildHelper();
    writeFixtureCache();
    runDirectly(entry);
    runCaller();
    process.stdout.write(`qa-token-helper-roundtrip: passed on ${process.platform}\n`);
  }
} catch (e) {
  process.stderr.write(`qa-token-helper-roundtrip: FAILED, ${e instanceof Error ? e.message : String(e)}\n`);
  process.exitCode = 1;
} finally {
  if (work !== '') rmSync(work, { recursive: true, force: true });
}
