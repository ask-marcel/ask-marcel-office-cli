#!/usr/bin/env bun
/*
 * Bundle smoke of @ask-marcel/office-write: the built write/dist/ run the way a
 * person and an MCP client run it, under node and bun.
 *
 *   1. dist/cli.js and dist/index.js hold no read code: no converter library
 *      (unpdf, mammoth, jszip, turndown), no playwright, no read graph
 *      (createReadGraph) and no member of the read tiers (getElevated,
 *      teamsChatIc3);
 *   2. `--version` answers the package version; `--help` lists the 7 write
 *      commands, `help-json` and `docs` (not `mcp`: the bin serves it before
 *      commander); `help-json` gives the package and 10 entries, `mcp` among
 *      them; `docs create-mail-draft` prints its page;
 *   3. a missing required flag and an unknown flag each exit 1 with the
 *      commander code and no crash line, before any token request (the run has
 *      a bad tier token, so a token request would answer env_token_invalid);
 *   4. a full call with `ASKMARCEL_TOKEN_BASIC=bad` answers env_token_invalid,
 *      exit 1 (never a helper search: a stale auth/dist can be on PATH);
 *   5. `list-drives` names @ask-marcel/office-read, exit 1;
 *   6. `dist/cli.js mcp` writes only JSON-RPC frames to stdout (a tolerant SDK
 *      client skips a stray line, so this reads the raw stdout) and has exactly
 *      the tools list-commands, get-command-docs and run-write-command;
 *   7. no probe leaves a file under its HOME except the update-notifier
 *      configstore of this package and bun's own cache (the bin writes no token
 *      cache and no locator);
 *   8. the library loads by name.
 *
 * Each probe gets a new temporary HOME (and USERPROFILE) and no ASKMARCEL_*
 * variable except the bad tier token where it says so (.claude/LESSONS.md,
 * 2026-10-16). Offline: no probe reaches Graph.
 *
 * Run from write/ after the build: `bun run build && bun ../scripts/qa-write-bundle-smoke.ts`
 * (or `bun run smoke`). Run it as `bun <file>`, not `bun run <file>`, so that
 * bun puts no `node` shim on PATH and the node probes run on node.
 * Exit non-zero on any failure.
 */
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, relative, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import type { SpawnSyncReturns } from 'node:child_process';

const ROOT = resolve(import.meta.dir, '../write');
if (!existsSync(join(ROOT, 'dist/cli.js')) || !existsSync(join(ROOT, 'dist/index.js'))) {
  console.error(`write bundle smoke: no dist/cli.js and dist/index.js under ${ROOT}; run \`bun run build\` in write/ first`);
  process.exit(1);
}
const PACKAGE = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')) as { name?: string; version?: string };
const runtimeVersion = (rt: string): string => spawnSync(rt, ['--version']).stdout?.toString().trim() || 'missing';
console.log(`write bundle smoke: ${ROOT} at ${PACKAGE.version}, node ${runtimeVersion('node')}, bun ${runtimeVersion('bun')}`);

const WRITE_COMMANDS = [
  'convert-calendar-event-attachment-to-pdf',
  'convert-group-post-attachment-to-pdf',
  'convert-mail-attachment-to-pdf',
  'create-forward-draft',
  'create-mail-draft',
  'create-reply-draft',
  'update-mail-draft',
];
const READ_MARKERS = ['unpdf', 'mammoth', 'jszip', 'turndown', 'playwright', 'createReadGraph', 'getElevated', 'teamsChatIc3'];
const BAD_TOKEN = { ASKMARCEL_TOKEN_BASIC: 'bad' };
const DRAFT_FLAGS = ['--subject', 'S', '--body-content', 'B', '--to-recipients', 'alice@example.com'];

const HOMES = mkdtempSync(join(tmpdir(), 'write-smoke-'));
const CALLER_ENV = Object.fromEntries(Object.entries(process.env).filter(([key]) => !key.startsWith('ASKMARCEL_')));
// The configstore file that update-notifier writes for this package, and the
// transpiler cache that bun itself keeps under HOME (.bun/install/cache).
const ALLOWED_FILE = /^(\.config\/configstore\/update-notifier-@ask-marcel\/office-write\.json|\.bun\/install\/cache\/.+)$/;

// A probe passes only when its process ended by itself with the expected status.
const PROBE_TIMEOUT_MS = 20000;
const exited = (p: SpawnSyncReturns<Buffer>, status: number): boolean => p.error === undefined && p.signal === null && p.status === status;
const exitNote = (p: SpawnSyncReturns<Buffer>): string => p.error?.message ?? (p.signal === null ? `exit ${p.status}` : `signal ${p.signal}`);

const filesUnder = (dir: string): ReadonlyArray<string> =>
  readdirSync(dir, { recursive: true, withFileTypes: true })
    .filter((entry) => entry.isFile())
    .map((entry) => relative(dir, join(entry.parentPath, entry.name)));

type Run = { readonly p: SpawnSyncReturns<Buffer>; readonly out: string; readonly err: string; readonly strayFiles: ReadonlyArray<string> };

const run = (rt: string, args: ReadonlyArray<string>, extraEnv: Readonly<Record<string, string>> = {}, input?: string): Run => {
  const home = mkdtempSync(join(HOMES, 'home-'));
  const p = spawnSync(rt, ['dist/cli.js', ...args], {
    cwd: ROOT,
    env: { ...CALLER_ENV, ...extraEnv, HOME: home, USERPROFILE: home },
    ...(input === undefined ? { stdio: ['ignore', 'pipe', 'pipe'] as const } : { input }),
    timeout: PROBE_TIMEOUT_MS,
    maxBuffer: 64 * 1024 * 1024,
  });
  const strayFiles = filesUnder(home).filter((path) => !ALLOWED_FILE.test(path));
  return { p, out: p.stdout?.toString() ?? '', err: p.stderr?.toString() ?? '', strayFiles };
};

// The JSON envelope of a failed command (`--output json`).
const errorCode = (out: string): string => {
  try {
    return String((JSON.parse(out) as { errorCode?: unknown }).errorCode ?? '');
  } catch {
    return `no JSON envelope: ${out.slice(0, 120)}`;
  }
};

type Check = readonly [string, boolean, string];

const probeBundles = (): ReadonlyArray<Check> =>
  ['dist/cli.js', 'dist/index.js'].map((file): Check => {
    const text = readFileSync(join(ROOT, file), 'utf8');
    const found = READ_MARKERS.filter((marker) => text.includes(marker));
    return [`${file} holds no read code`, found.length === 0, found.length === 0 ? 'clean' : `holds ${found.join(', ')}`];
  });

const probeCli = (rt: string): ReadonlyArray<Check> => {
  const version = run(rt, ['--version']);
  const help = run(rt, ['--help']);
  const helpListed = (name: string): boolean => help.out.split('\n').some((line) => line.startsWith(`  ${name} `));
  const missingHelp = [...WRITE_COMMANDS, 'help-json', 'docs'].filter((name) => !helpListed(name));
  const helpJson = run(rt, ['help-json']);
  const manifest = ((): { package?: string; commands?: ReadonlyArray<{ name: string }> } => {
    try {
      return JSON.parse(helpJson.out) as { package?: string; commands?: ReadonlyArray<{ name: string }> };
    } catch {
      return {};
    }
  })();
  const names = (manifest.commands ?? []).map((c) => c.name);
  const docs = run(rt, ['docs', 'create-mail-draft']);
  const missing = run(rt, ['create-mail-draft', '--subject', 'S', '--output', 'json'], BAD_TOKEN);
  const unknown = run(rt, ['create-mail-draft', ...DRAFT_FLAGS, '--subjet', 'x', '--output', 'json'], BAD_TOKEN);
  const bad = run(rt, ['create-mail-draft', ...DRAFT_FLAGS, '--output', 'json'], BAD_TOKEN);
  const readName = run(rt, ['list-drives']);
  const all = [version, help, helpJson, docs, missing, unknown, bad, readName];
  const stray = [...new Set(all.flatMap((r) => r.strayFiles))];
  return [
    ['--version answers the package version', exited(version.p, 0) && version.out.trim() === PACKAGE.version, `${exitNote(version.p)}, ${version.out.trim()}`],
    [
      '--help lists the 7 write commands, help-json and docs, and not mcp',
      exited(help.p, 0) && missingHelp.length === 0 && !helpListed('mcp'),
      missingHelp.length === 0 ? exitNote(help.p) : `missing ${missingHelp.join(', ')}`,
    ],
    [
      'help-json gives the package and 10 entries, mcp among them',
      exited(helpJson.p, 0) && manifest.package === PACKAGE.name && names.length === 10 && names.includes('mcp'),
      `${manifest.package ?? 'no package'}, ${names.length} entries`,
    ],
    ['docs create-mail-draft prints its page', exited(docs.p, 0) && docs.out.startsWith('# `create-mail-draft`'), exitNote(docs.p)],
    [
      'a missing required flag exits 1 before any token request',
      exited(missing.p, 1) && errorCode(missing.out) === 'commander.missingMandatoryOptionValue' && missing.err === '',
      `${exitNote(missing.p)}, ${errorCode(missing.out)}`,
    ],
    [
      'an unknown flag exits 1 before any token request',
      exited(unknown.p, 1) && errorCode(unknown.out) === 'commander.unknownOption' && unknown.err === '',
      `${exitNote(unknown.p)}, ${errorCode(unknown.out)}`,
    ],
    ['a bad tier token answers env_token_invalid', exited(bad.p, 1) && errorCode(bad.out) === 'env_token_invalid', `${exitNote(bad.p)}, ${errorCode(bad.out)}`],
    ['list-drives names @ask-marcel/office-read', exited(readName.p, 1) && readName.out.includes('@ask-marcel/office-read'), exitNote(readName.p)],
    ['no file under HOME but the update-notifier configstore', stray.length === 0, stray.length === 0 ? 'clean' : stray.join(', ')],
  ];
};

const MCP_INIT = { jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'qa', version: '1' } } };
const MCP_LIST = { jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} };
// A tool call that needs no token: the docs of one write command.
const MCP_DOCS = { jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'get-command-docs', arguments: { command: 'create-mail-draft' } } };

type Frame = {
  readonly jsonrpc?: string;
  readonly id?: number;
  readonly result?: {
    readonly serverInfo?: { readonly name?: string };
    readonly tools?: ReadonlyArray<{ readonly name: string }>;
    readonly content?: ReadonlyArray<{ readonly text?: string }>;
  };
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
  const served = run(rt, ['mcp'], {}, `${[MCP_INIT, MCP_LIST, MCP_DOCS].map((m) => JSON.stringify(m)).join('\n')}\n`);
  const read = frames(served.out);
  const ok = typeof read !== 'string';
  const server = ok ? (read.find((frame) => frame.id === 1)?.result?.serverInfo?.name ?? '') : '';
  const tools = ok ? (read.find((frame) => frame.id === 2)?.result?.tools ?? []).map((tool) => tool.name).toSorted() : [];
  const docsText = ok ? (read.find((frame) => frame.id === 3)?.result?.content ?? []).map((part) => part.text ?? '').join('') : '';
  return [
    ['the server exits cleanly when its input closes', exited(served.p, 0), exitNote(served.p)],
    ['stdout holds only JSON-RPC frames', ok, ok ? `${read.length} line(s)` : read],
    ['the server is ask-marcel-office-write', server === 'ask-marcel-office-write', server],
    [
      'the tools are exactly list-commands, get-command-docs and run-write-command',
      JSON.stringify(tools) === JSON.stringify(['get-command-docs', 'list-commands', 'run-write-command']),
      tools.join(', '),
    ],
    ['get-command-docs answers the create-mail-draft page', docsText.startsWith('# `create-mail-draft`'), docsText.split('\n')[0] ?? ''],
    ['no file under HOME but the update-notifier configstore', served.strayFiles.length === 0, served.strayFiles.length === 0 ? 'clean' : served.strayFiles.join(', ')],
  ];
};

const probeLibrary = (rt: string): ReadonlyArray<Check> => {
  const script = `const m = await import('${PACKAGE.name}'); process.stdout.write(Object.keys(m).sort().join(',') + ' ' + Object.keys(m.commands).length);`;
  const home = mkdtempSync(join(HOMES, 'home-'));
  const p = spawnSync(rt, rt === 'node' ? ['--input-type=module', '-e', script] : ['-e', script], {
    cwd: ROOT,
    env: { ...CALLER_ENV, HOME: home, USERPROFILE: home },
    stdio: ['ignore', 'pipe', 'pipe'],
    timeout: PROBE_TIMEOUT_MS,
  });
  const out = p.stdout?.toString() ?? '';
  return [
    [
      'the library loads by name with its exports and the 7 commands',
      exited(p, 0) && out === 'commands,createEnvThenHelperTokenSource,createWriteClient,createWriteGraph 7',
      out || (p.stderr?.toString().slice(0, 200) ?? exitNote(p)),
    ],
  ];
};

let fails = 0;
const report = (title: string, checks: ReadonlyArray<Check>): void => {
  console.log(`=== ${title} ===`);
  for (const [label, passed, note] of checks) {
    if (!passed) fails++;
    console.log(`  ${passed ? '✓' : '✗'} ${label} (${note})`);
  }
};
report('bundles', probeBundles());
for (const rt of ['node', 'bun']) {
  report(`cli @ ${rt}`, probeCli(rt));
  report(`mcp stdio @ ${rt}`, probeMcp(rt));
  report(`library @ ${rt}`, probeLibrary(rt));
}
rmSync(HOMES, { recursive: true, force: true });
console.log(`\n${fails === 0 ? 'WRITE BUNDLE OK under node + bun ✓' : `!! ${fails} write bundle FAILURES`}`);
process.exit(fails > 0 ? 1 : 0);
