#!/usr/bin/env bun
/*
 * Bundle smoke of @ask-marcel/office-auth: the built auth/dist/ run the way a
 * person and the read and write packages run it, under node and bun.
 *
 *   1. dist/token.js, the token helper, is a node script without commander,
 *      winston, the update notifier or the MCP SDK and zod, which keeps it
 *      near bare runtime startup;
 *   2. a cache hit prints the one token line, and a home with no session fails
 *      fast with not_authenticated (exit 1, stdin is not a terminal, so no
 *      browser opens);
 *   3. each run records the locator: auth/dist/token.js at the package version;
 *   4. `dist/cli.js token` prints the same line, and `dist/cli.js --version`
 *      answers the package version;
 *   5. `dist/cli.js mcp` writes only JSON-RPC frames to stdout (a tolerant SDK
 *      client skips a stray line, so this reads the raw stdout), has exactly the
 *      tools `login` and `status`, and no tool text holds the cached token.
 *
 * Each probe gets a new temporary HOME (and USERPROFILE) and no ASKMARCEL_*
 * variable, so a run never touches the real ~/.ask-marcel/ and a token variable
 * of the caller cannot answer for the fixture (.claude/LESSONS.md, 2026-10-04
 * and 2026-10-16). Offline: the fixture token is a cache hit.
 *
 * Run from auth/ after the build: `bun run build && bun ../scripts/qa-auth-bundle-smoke.ts`
 * (or `bun run smoke`). Run it as `bun <file>`, not `bun run <file>`, so that
 * bun puts no `node` shim on PATH and the node probes run on node.
 * Exit non-zero on any failure.
 */
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import type { SpawnSyncReturns } from 'node:child_process';

const ROOT = resolve(import.meta.dir, '../auth');
if (!existsSync(join(ROOT, 'dist/cli.js')) || !existsSync(join(ROOT, 'dist/token.js'))) {
  console.error(`auth bundle smoke: no dist/cli.js and dist/token.js under ${ROOT}; run \`bun run build\` in auth/ first`);
  process.exit(1);
}
const PACKAGE_VERSION = (JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')) as { version?: string }).version;
const runtimeVersion = (rt: string): string => spawnSync(rt, ['--version']).stdout?.toString().trim() || 'missing';
console.log(`auth bundle smoke: ${ROOT} at ${PACKAGE_VERSION}, node ${runtimeVersion('node')}, bun ${runtimeVersion('bun')}`);

const HOMES = mkdtempSync(join(tmpdir(), 'auth-smoke-'));
const CALLER_ENV = Object.fromEntries(Object.entries(process.env).filter(([key]) => !key.startsWith('ASKMARCEL_')));
const homeEnv = (home: string): NodeJS.ProcessEnv => ({ ...CALLER_ENV, HOME: home, USERPROFILE: home });

// A probe passes only when its process ended by itself with the expected status.
const PROBE_TIMEOUT_MS = 20000;
const exited = (p: SpawnSyncReturns<Buffer>, status: number): boolean => p.error === undefined && p.signal === null && p.status === status;
const exitNote = (p: SpawnSyncReturns<Buffer>): string => p.error?.message ?? (p.signal === null ? `exit ${p.status}` : `signal ${p.signal}`);

// base64url, as a token is sent: the AccessToken brand refuses padded base64.
const TOKEN_EXP = Math.floor(Date.now() / 1000) + 3600;
const jwtSegment = (value: unknown): string => Buffer.from(JSON.stringify(value)).toString('base64url');
const TOKEN_FIXTURE = `${jwtSegment({ alg: 'none' })}.${jwtSegment({ exp: TOKEN_EXP, aud: 'https://graph.microsoft.com' })}.sig`;

const newHome = (withToken: boolean): string => {
  const home = mkdtempSync(join(HOMES, 'home-'));
  mkdirSync(join(home, '.ask-marcel'), { recursive: true });
  if (withToken) writeFileSync(join(home, '.ask-marcel', 'token-cache.json'), JSON.stringify({ access_token: TOKEN_FIXTURE, expires_on: TOKEN_EXP, refresh_token: '' }));
  return home;
};

// The locator that a run left under `home` names auth/dist/token.js at this version.
const recordsHelper = (home: string): boolean => {
  const path = join(home, '.ask-marcel', 'token-helper.json');
  if (!existsSync(path)) return false;
  const locator = JSON.parse(readFileSync(path, 'utf8')) as { entry?: string; version?: string };
  return locator.entry === join(ROOT, 'dist/token.js') && locator.version === PACKAGE_VERSION;
};

const run = (rt: string, args: ReadonlyArray<string>, home: string, input?: string): { p: SpawnSyncReturns<Buffer>; ms: number } => {
  const started = performance.now();
  const p = spawnSync(rt, args, {
    cwd: ROOT,
    env: homeEnv(home),
    ...(input === undefined ? { stdio: ['ignore', 'pipe', 'pipe'] as const } : { input }),
    timeout: PROBE_TIMEOUT_MS,
    maxBuffer: 64 * 1024 * 1024,
  });
  return { p, ms: Math.round(performance.now() - started) };
};

// The helper prints exactly one JSON line on stdout.
const oneLine = (p: SpawnSyncReturns<Buffer>): Record<string, unknown> | undefined => {
  const lines = (p.stdout?.toString() ?? '').split('\n').filter((l) => l !== '');
  if (lines.length !== 1) return undefined;
  try {
    return JSON.parse(lines[0] ?? '') as Record<string, unknown>;
  } catch {
    return undefined;
  }
};

type Check = readonly [string, boolean, string];

const probeToken = (rt: string): ReadonlyArray<Check> => {
  const bundle = readFileSync(join(ROOT, 'dist/token.js'), 'utf8');
  const heavy = ['commander', 'winston', 'update-notifier', 'modelcontextprotocol', 'zod'].filter((name) => bundle.includes(name));
  const warm = newHome(true);
  const hit = run(rt, ['dist/token.js', '--tier', 'basic'], warm);
  const hitLine = oneLine(hit.p);
  const viaCliHome = newHome(true);
  const viaCli = run(rt, ['dist/cli.js', 'token', '--tier', 'basic'], viaCliHome);
  const coldHome = newHome(false);
  const cold = run(rt, ['dist/token.js', '--tier', 'basic'], coldHome);
  const versionHome = newHome(false);
  const version = run(rt, ['dist/cli.js', '--version'], versionHome);
  return [
    [
      'dist/token.js is a node script without the CLI modules',
      bundle.startsWith('#!/usr/bin/env node') && heavy.length === 0,
      heavy.length === 0 ? 'clean' : `carries ${heavy.join(', ')}`,
    ],
    ['a cache hit prints the token line', exited(hit.p, 0) && hitLine?.['accessToken'] === TOKEN_FIXTURE && hitLine['expiresOn'] === TOKEN_EXP, `${exitNote(hit.p)}, ${hit.ms} ms`],
    ['the helper records auth/dist/token.js at the package version', recordsHelper(warm), String(PACKAGE_VERSION)],
    ['no session fails fast with not_authenticated, no browser', exited(cold.p, 1) && oneLine(cold.p)?.['errorCode'] === 'not_authenticated', `${exitNote(cold.p)}, ${cold.ms} ms`],
    [
      'cli.js token prints the same line and records the helper',
      exited(viaCli.p, 0) && oneLine(viaCli.p)?.['accessToken'] === TOKEN_FIXTURE && recordsHelper(viaCliHome),
      exitNote(viaCli.p),
    ],
    [
      'cli.js --version answers the package version and records the helper',
      exited(version.p, 0) && version.p.stdout.toString().trim() === PACKAGE_VERSION && recordsHelper(versionHome),
      exitNote(version.p),
    ],
  ];
};

const MCP_INIT = { jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'qa', version: '1' } } };
const MCP_LIST = { jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} };
// The status tool reads the cached token's report: a call that needs no browser.
const MCP_STATUS = { jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'status', arguments: {} } };

type Frame = {
  readonly jsonrpc?: string;
  readonly id?: number;
  readonly result?: { readonly tools?: ReadonlyArray<{ readonly name: string }>; readonly content?: ReadonlyArray<{ readonly text?: string }> };
};

const frames = (stdout: string): ReadonlyArray<Frame> | string => {
  const lines = stdout.split('\n').filter((l) => l.trim() !== '');
  if (lines.length === 0) return 'no stdout';
  const parsed = lines.map((line): Frame | undefined => {
    try {
      return JSON.parse(line) as Frame;
    } catch {
      return undefined;
    }
  });
  const polluted = parsed.filter((frame) => frame?.jsonrpc !== '2.0').length;
  return polluted === 0 ? (parsed as ReadonlyArray<Frame>) : `STDOUT POLLUTED by ${polluted} non-JSON-RPC line(s)`;
};

const probeMcp = (rt: string): ReadonlyArray<Check> => {
  const home = newHome(true);
  const { p } = run(rt, ['dist/cli.js', 'mcp'], home, `${[MCP_INIT, MCP_LIST, MCP_STATUS].map((m) => JSON.stringify(m)).join('\n')}\n`);
  const stdout = p.stdout?.toString() ?? '';
  const read = frames(stdout);
  const ok = typeof read !== 'string';
  const tools = ok ? (read.find((frame) => frame.id === 2)?.result?.tools ?? []).map((tool) => tool.name) : [];
  const statusText = ok ? (read.find((frame) => frame.id === 3)?.result?.content ?? []).map((part) => part.text ?? '').join('') : '';
  const everything = `${stdout}${p.stderr?.toString() ?? ''}`;
  const leaked = everything.includes(TOKEN_FIXTURE) || everything.includes(TOKEN_FIXTURE.split('.')[1] ?? TOKEN_FIXTURE);
  return [
    ['the server exits cleanly when its input closes', exited(p, 0), exitNote(p)],
    ['stdout holds only JSON-RPC frames', ok, ok ? `${read.length} line(s)` : read],
    ['the tools are exactly login and status', JSON.stringify(tools) === JSON.stringify(['login', 'status']), tools.join(', ')],
    ['the status tool reports the cached basic token', statusText.includes('basic'), statusText.split('\n')[0] ?? ''],
    ['no tool text and no stderr line holds the token', !leaked, leaked ? 'the token is in the output' : 'clean'],
    ['the server records auth/dist/token.js', recordsHelper(home), String(PACKAGE_VERSION)],
  ];
};

let fails = 0;
for (const rt of ['node', 'bun']) {
  for (const [title, checks] of [
    [`token helper @ ${rt}`, probeToken(rt)],
    [`mcp stdio @ ${rt}`, probeMcp(rt)],
  ] as const) {
    console.log(`=== ${title} ===`);
    for (const [label, passed, note] of checks) {
      if (!passed) fails++;
      console.log(`  ${passed ? '✓' : '✗'} ${label} (${note})`);
    }
  }
}
rmSync(HOMES, { recursive: true, force: true });
console.log(`\n${fails === 0 ? 'AUTH BUNDLE OK under node + bun ✓' : `!! ${fails} auth bundle FAILURES`}`);
process.exit(fails > 0 ? 1 : 0);
