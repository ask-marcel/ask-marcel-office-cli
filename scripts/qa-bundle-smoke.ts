#!/usr/bin/env bun
/*
 * QA Phase A7 — bundle-interop smoke across EVERY converter (docs/QA-PLAYBOOK.md §A7).
 *
 * `bun test` runs the conversion use-cases from SOURCE and can NEVER catch
 * bundler-interop breakage — the `.msg` "Object is not a constructor" bug shipped
 * green through every source gate and only a bundle smoke found it. Each format
 * loads a different vendored library (mammoth/sheetjs/jszip/unpdf/word-extractor/
 * msgreader/fast-xml-parser), so each is its own interop surface.
 *
 * This generates a real on-disk fixture for every format from the repo's own
 * fixture builders, then runs the built `dist/cli.js` on each under BOTH `node`
 * and `bun`, asserting `ok`. It then checks the read bin's own answers: its
 * names, the refusal of a write or sign-in command, a bad tier token, an
 * unknown flag, the files a run leaves under HOME, and that the shipped
 * dist/cli.js and dist/index.js hold no write code (D9) and that the library
 * holds no CLI or MCP module. The library loads by name with exactly its
 * exports, and its token source gets a token from a token helper under node
 * and under bun.
 * Run from read/: `bun run build && bun ../scripts/qa-bundle-smoke.ts` (or
 * `bun run smoke`). `--package-root <dir>` runs the same probes against another
 * copy of the package (its `dist/` under <dir>), which is how
 * qa-packed-install-smoke.ts checks the tarball as installed outside the repo.
 * Exit non-zero on any failure.
 */
import { chmodSync, existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import type { SpawnSyncReturns } from 'node:child_process';
import {
  buildRichDocx, buildRichXlsx, buildRichPptx, buildRichOdt, buildRichOds, buildRichOdp,
  buildPdfWithText, buildPdfWithImage, buildSampleDoc, buildLegacyXls, buildSampleMsg, buildSampleEml, buildSampleZipArchive,
} from '../read/src/test-helpers/office-fixtures.ts';

const DIR = '/tmp/qa/smoke';
mkdirSync(DIR, { recursive: true });
const write = async (name: string, bytes: Uint8Array): Promise<string> => { const p = `${DIR}/${name}`; await Bun.write(p, bytes); return p; };

// generate a fixture per format (await handles both sync + async builders)
const F: Record<string, string> = {};
F.docx = await write('f.docx', await buildRichDocx());
F.xlsx = await write('f.xlsx', await buildRichXlsx());
F.pptx = await write('f.pptx', await buildRichPptx());
F.odt = await write('f.odt', await buildRichOdt());
F.ods = await write('f.ods', await buildRichOds());
F.odp = await write('f.odp', await buildRichOdp());
F.pdf = await write('f.pdf', await buildPdfWithText());
F.doc = await write('f.doc', await buildSampleDoc());
F.xls = await write('f.xls', await buildLegacyXls());
F.msg = await write('f.msg', await buildSampleMsg());
F.eml = await write('f.eml', buildSampleEml());
F.zip = await write('f.zip', await buildSampleZipArchive());
F.csv = await write('f.csv', new TextEncoder().encode('a,b,c\n1,2,3\n'));
// HTML goes through turndown in the bundle since 2.8, head dropped and data: images placeholdered.
F.html = await write('f.html', new TextEncoder().encode('<html><head><title>t</title></head><body><h1>Q3</h1><table><tr><th>A</th></tr><tr><td>1</td></tr></table><img src="data:image/png;base64,AAAA" alt="c"></body></html>'));
const pdfImg = await write('img.pdf', await buildPdfWithImage());
// The vendored image-bearing docx: an image extractor answering media=0 here is a failure, not a pass.
const imageDocx = resolve('src/test-helpers/assets/image-sample.docx');

// Every probe runs with the package root as its working directory, so `dist/cli.js`
// and the library probe's `./dist/index.js` resolve inside the copy under test. A
// flag given without a usable value is an error, never a silent fall back to the
// repo's own dist.
const packageRoot = (argv: ReadonlyArray<string>): string | undefined => {
  const inline = argv.find((a) => a.startsWith('--package-root='));
  if (inline !== undefined) return inline.slice('--package-root='.length) || undefined;
  const at = argv.indexOf('--package-root');
  if (at === -1) return '.';
  const value = argv[at + 1];
  return value === undefined || value.startsWith('--') ? undefined : value;
};
const rootArg = packageRoot(process.argv);
const ROOT = resolve(rootArg ?? '.');
if (rootArg === undefined || !existsSync(join(ROOT, 'dist/cli.js'))) {
  console.error(`bundle smoke: no dist/cli.js under ${rootArg === undefined ? '(missing --package-root value)' : ROOT}; build first, or check --package-root`);
  process.exit(1);
}
const runtimeVersion = (rt: string): string => spawnSync(rt, ['--version']).stdout?.toString().trim() || 'missing';
console.log(`bundle smoke: package root ${ROOT}, node ${runtimeVersion('node')}, bun ${runtimeVersion('bun')}`);

// The probes run on a scratch home, so a smoke run never reads or writes the user's.
const SMOKE_HOME = mkdtempSync(join(DIR, 'home-'));
const SMOKE_ENV = { ...process.env, HOME: SMOKE_HOME, USERPROFILE: SMOKE_HOME };

// A probe passes only when its process ended by itself with the expected status:
// a right answer from a process that then crashed, hung or exited non-zero is a failure.
const PROBE_TIMEOUT_MS = 20000;
const exited = (p: SpawnSyncReturns<Buffer>, status: number): boolean => p.error === undefined && p.signal === null && p.status === status;
const exitNote = (p: SpawnSyncReturns<Buffer>): string => p.error?.message ?? (p.signal === null ? `exit ${p.status}` : `signal ${p.signal}`);

// A flag the bundle must thread end to end: argv beyond `--path`, with what the answer must hold.
const probeArgs = (rt: string, args: ReadonlyArray<string>, expect: (d: { ok: boolean; data?: { text?: string; media?: ReadonlyArray<unknown> }; error?: string }) => boolean, status = 0): boolean => {
  const p = spawnSync(rt, ['dist/cli.js', ...args, '--output', 'json'], { cwd: ROOT, env: SMOKE_ENV, timeout: PROBE_TIMEOUT_MS, maxBuffer: 64 * 1024 * 1024 });
  try {
    return exited(p, status) && expect(JSON.parse(p.stdout?.toString() || ''));
  } catch {
    return false;
  }
};

// The `diff` package is bundled (not --external) and only Graph-backed commands reach
// it, so drive diff-drive-items through the library bundle against a stub client.
const DIFF_PROBE = `import { commands } from './dist/index.js';
const graph = { get: async () => ({ ok: true, value: { name: 'plan.md' } }), getBinary: async (p) => ({ ok: true, value: { contentType: 'text/plain', size: 3, text: p.includes('/i6/') ? 'one\\ntwo' : 'one\\n2' } }) };
const r = await commands['diff-drive-items'].execute(graph, { driveId: 'd1', itemId: 'i6', otherDriveId: 'd1', otherItemId: 'i7' });
process.stdout.write(JSON.stringify({ ok: r.ok, added: r.value?.added, removed: r.value?.removed }));`;
const probeDiff = (rt: string): boolean => {
  const p = spawnSync(rt, rt === 'node' ? ['--input-type=module', '-e', DIFF_PROBE] : ['-e', DIFF_PROBE], { cwd: ROOT, timeout: PROBE_TIMEOUT_MS });
  try {
    const d = JSON.parse(p.stdout?.toString() || '') as { ok: boolean; added?: number; removed?: number };
    return exited(p, 0) && d.ok && d.added === 1 && d.removed === 1;
  } catch {
    return false;
  }
};

const probe = (rt: string, cmd: string, path: string): { ok: boolean; note: string } => {
  const p = spawnSync(rt, ['dist/cli.js', cmd, '--path', path, '--output', 'json'], { cwd: ROOT, env: SMOKE_ENV, timeout: PROBE_TIMEOUT_MS, maxBuffer: 64 * 1024 * 1024 });
  if (!exited(p, 0)) return { ok: false, note: `${exitNote(p)}: ${(p.stdout?.toString() || p.stderr?.toString() || '').slice(0, 60)}` };
  try { const d = JSON.parse(p.stdout?.toString() || ''); return { ok: d.ok === true, note: d.ok ? (d.data?.media ? `media=${d.data.media.length}` : `${d.data?.contentType || ''}`) : `ERR:${d.errorCode || String(d.error).slice(0, 40)}` }; }
  catch { return { ok: false, note: 'CRASH/non-JSON: ' + (p.stdout?.toString() || p.stderr?.toString() || '').slice(0, 60) }; }
};

/*
 * `ask-marcel-office-read mcp` speaks JSON-RPC over stdout, so it needs a DIFFERENT probe.
 *
 * This asserts on RAW stdout, deliberately NOT via the SDK's Client. Two
 * distinct failure modes have to be caught and only one of them is visible to
 * a client:
 *
 *   1. Protocol breakage      -> a Client would catch this.
 *   2. stdout POLLUTION       -> a Client would NOT. Verified 2026-07-17: with
 *      `process.stdout.write('STRAY BANNER\n')` injected into the mcp path, the
 *      bundle emitted `STRAY BANNER\n{"result":...}` and a Client-based probe
 *      still reported a clean 5-tool handshake. The SDK's ReadBuffer skips
 *      lines it cannot parse, so a tolerant client hides the very bug this gate
 *      exists to find. The spec says the server MUST NOT write non-MCP output
 *      to stdout; a stricter client (or a future SDK) would drop the session.
 *
 * `mcp.test.ts` cannot cover this either — InMemoryTransport never touches
 * stdout. So this probe is the ONLY thing standing between a stray banner,
 * log line, or debug print and a broken release.
 */
const MCP_INIT = { jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'qa', version: '1' } } };
const MCP_LIST = { jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} };
// A tool call that needs no login: an unknown command answers a tool error naming the right one.
const MCP_CALL = { jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'run-command', arguments: { command: 'lsit-drives', params: {} } } };
// The read server is the read gateway only: `login` and `status` are tools of the auth server,
// and `run-write-command` is a tool of the write server.
const GATEWAY_TOOLS = ['get-command-docs', 'list-commands', 'run-command'];

const probeMcp = (rt: string): { ok: boolean; note: string } => {
  const p = spawnSync(rt, ['dist/cli.js', 'mcp'], {
    cwd: ROOT,
    env: SMOKE_ENV,
    input: `${JSON.stringify(MCP_INIT)}\n${JSON.stringify(MCP_LIST)}\n${JSON.stringify(MCP_CALL)}\n`,
    timeout: PROBE_TIMEOUT_MS,
    maxBuffer: 64 * 1024 * 1024,
  });
  if (!exited(p, 0)) return { ok: false, note: `the server did not exit cleanly when its input closed (${exitNote(p)})` };
  const stdout = p.stdout?.toString() ?? '';
  const lines = stdout.split('\n').filter((l) => l.trim() !== '');
  if (lines.length === 0) return { ok: false, note: `no stdout (stderr: ${(p.stderr?.toString() ?? '').slice(0, 60)})` };
  // EVERY line must be a JSON-RPC frame. One banner and this fails, which is
  // the whole point.
  const polluted = lines.filter((l) => {
    try {
      return (JSON.parse(l) as { jsonrpc?: string }).jsonrpc !== '2.0';
    } catch {
      return true;
    }
  });
  if (polluted.length > 0) return { ok: false, note: `STDOUT POLLUTED by ${polluted.length} non-JSON-RPC line(s): ${JSON.stringify(polluted[0]?.slice(0, 40))}` };
  const listReply = lines.map((l) => JSON.parse(l) as { id?: number; result?: { tools?: ReadonlyArray<{ name: string }> } }).find((m) => m.id === 2);
  const tools = listReply?.result?.tools ?? [];
  const names = tools.map((t) => t.name).toSorted((a, b) => a.localeCompare(b));
  if (names.join(',') !== GATEWAY_TOOLS.join(',')) return { ok: false, note: `expected the ${GATEWAY_TOOLS.length} gateway tools ${GATEWAY_TOOLS.join(', ')}, got ${names.join(', ')}` };
  const callReply = lines.map((l) => JSON.parse(l) as { id?: number; result?: { isError?: boolean; content?: ReadonlyArray<{ text?: string }> } }).find((m) => m.id === 3);
  if (callReply?.result?.isError !== true || !(callReply.result.content?.[0]?.text ?? '').includes('Did you mean')) return { ok: false, note: 'tools/call on a mistyped command did not answer a did-you-mean tool error' };
  return { ok: true, note: `${tools.length} tools, ${lines.length} clean JSON-RPC line(s)` };
};

let fails = 0;
for (const rt of ['node', 'bun']) {
  console.log(`\n=== convert-local-file-to-markdown @ ${rt} ===`);
  for (const [fmt, path] of Object.entries(F)) { const r = probe(rt, 'convert-local-file-to-markdown', path); if (!r.ok) fails++; console.log(`  ${r.ok ? '✓' : '✗'} ${fmt.padEnd(5)} -> ${r.note}`); }
  console.log(`=== extract-local-file-images @ ${rt} ===`);
  for (const [fmt, path] of [['docx', F.docx], ['xlsx', F.xlsx], ['pptx', F.pptx], ['pdf', pdfImg]]) { const r = probe(rt, 'extract-local-file-images', path); if (!r.ok) fails++; console.log(`  ${r.ok ? '✓' : '✗'} ${fmt.padEnd(5)} -> ${r.note}`); }
  const img = probe(rt, 'extract-local-file-images', imageDocx);
  const imgOk = img.ok && img.note !== 'media=0';
  if (!imgOk) fails++;
  console.log(`  ${imgOk ? '✓' : '✗'} image-docx -> ${img.note} (must be > 0)`);
  console.log(`=== flags and bundled libraries @ ${rt} ===`);
  const checks: ReadonlyArray<[string, boolean]> = [
    ['--max-cells caps a sheet', probeArgs(rt, ['convert-local-file-to-markdown', '--path', F.csv ?? '', '--max-cells', '1'], (d) => d.ok && (d.data?.text ?? '').includes('Table omitted'))],
    ['--sheet names the sheets', probeArgs(rt, ['convert-local-file-to-markdown', '--path', F.xlsx ?? '', '--sheet', 'NoSuchSheet'], (d) => !d.ok && (d.error ?? '').includes('no sheet named'), 1)],
    ['html: data image placeholder', probeArgs(rt, ['convert-local-file-to-markdown', '--path', F.html ?? ''], (d) => d.ok && (d.data?.text ?? '').includes('[image: c]'))],
    ['diff package (library bundle)', probeDiff(rt)],
  ];
  for (const [label, passed] of checks) {
    if (!passed) fails++;
    console.log(`  ${passed ? '✓' : '✗'} ${label}`);
  }
  console.log(`=== mcp stdio handshake @ ${rt} ===`);
  const m = probeMcp(rt); if (!m.ok) fails++; console.log(`  ${m.ok ? '✓' : '✗'} mcp   -> ${m.note}`);
}
/*
 * The read bin's own answers (package split, phase 6), as the write smoke checks
 * the write bin. Each probe gets a new HOME and no ASKMARCEL_* variable except
 * the bad tier token where it says so: never a helper search, because a stale
 * auth/dist can be on PATH (.claude/LESSONS.md, 2026-10-16 and 2026-10-18).
 */
const PACKAGE = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')) as { name?: string; version?: string };
const READ_HOMES = mkdtempSync(join(DIR, 'read-homes-'));
const CALLER_ENV = Object.fromEntries(Object.entries(process.env).filter(([key]) => !key.startsWith('ASKMARCEL_')));
const BAD_TOKEN = { ASKMARCEL_TOKEN_BASIC: 'bad' };
// A CLI run may leave the update-notifier configstore of this package; an mcp
// run never runs the notifier. Both may leave bun's own transpiler cache.
const BUN_CACHE = /^\.bun\/install\/cache\/.+$/;
const NOTIFIER_FILE = /^\.config\/configstore\/update-notifier-@ask-marcel\/office-read\.json$/;
// Write code in a shipped file (D9): the write methods in any quote, the upload
// session, the write graph and the browser sign-in.
const WRITE_MARKERS: ReadonlyArray<readonly [string, RegExp]> = [
  ['"PUT"', /(["'`])PUT\1/],
  ['"PATCH"', /(["'`])PATCH\1/],
  ['"DELETE"', /(["'`])DELETE\1/],
  ['createUploadSession', /createUploadSession/],
  ['createWriteGraph', /createWriteGraph/],
  ['playwright', /playwright/],
];
// What only the bin needs: the library entry loads none of them.
const BIN_MODULES = ['winston', 'commander', 'modelcontextprotocol', 'update-notifier'];

const filesUnder = (dir: string): ReadonlyArray<string> =>
  readdirSync(dir, { recursive: true, withFileTypes: true })
    .filter((entry) => entry.isFile())
    .map((entry) => relative(dir, join(entry.parentPath, entry.name)));

type ReadRun = { readonly p: SpawnSyncReturns<Buffer>; readonly out: string; readonly files: ReadonlyArray<string> };
const runRead = (rt: string, args: ReadonlyArray<string>, extraEnv: Readonly<Record<string, string>> = {}, input?: string): ReadRun => {
  const home = mkdtempSync(join(READ_HOMES, 'home-'));
  const p = spawnSync(rt, ['dist/cli.js', ...args], {
    cwd: ROOT,
    env: { ...CALLER_ENV, ...extraEnv, HOME: home, USERPROFILE: home },
    ...(input === undefined ? { stdio: ['ignore', 'pipe', 'pipe'] as const } : { input }),
    timeout: PROBE_TIMEOUT_MS,
    maxBuffer: 64 * 1024 * 1024,
  });
  return { p, out: p.stdout?.toString() ?? '', files: filesUnder(home) };
};
const envelope = (out: string): { errorCode?: string; error?: string } => {
  try {
    return JSON.parse(out) as { errorCode?: string; error?: string };
  } catch {
    return { error: `no JSON envelope: ${out.slice(0, 80)}` };
  }
};

type Check = readonly [string, boolean, string];

const probeShipped = (): ReadonlyArray<Check> =>
  ['dist/cli.js', 'dist/index.js'].map((file): Check => {
    const text = readFileSync(join(ROOT, file), 'utf8');
    const found = [
      ...WRITE_MARKERS.filter(([, marker]) => marker.test(text)).map(([name]) => name),
      ...(file === 'dist/index.js' ? BIN_MODULES.filter((name) => text.includes(name)) : []),
    ];
    return [`${file} holds no write code${file === 'dist/index.js' ? ' and no bin module' : ''}`, found.length === 0, found.length === 0 ? 'clean' : `holds ${found.join(', ')}`];
  });

const probeReadCli = (rt: string): ReadonlyArray<Check> => {
  const version = runRead(rt, ['--version']);
  const helpJson = runRead(rt, ['help-json']);
  const manifest = ((): { package?: string; commands?: ReadonlyArray<{ name: string }> } => {
    try {
      return JSON.parse(helpJson.out) as { package?: string; commands?: ReadonlyArray<{ name: string }> };
    } catch {
      return {};
    }
  })();
  const entries = (manifest.commands ?? []).map((c) => c.name);
  const write = runRead(rt, ['create-mail-draft', '--output', 'json']);
  const login = runRead(rt, ['login', '--output', 'json']);
  const bad = runRead(rt, ['list-drives', '--output', 'json'], BAD_TOKEN);
  const unknown = runRead(rt, ['list-drives', '--nope', 'x', '--output', 'json'], BAD_TOKEN);
  const stray = [...new Set([version, helpJson, write, login, bad, unknown].flatMap((r) => r.files.filter((path) => !NOTIFIER_FILE.test(path) && !BUN_CACHE.test(path))))];
  return [
    ['--version answers the manifest version', exited(version.p, 0) && version.out.trim() === PACKAGE.version, `${exitNote(version.p)}, ${version.out.trim()}`],
    [
      'help-json names the package with 207 entries (204 commands, docs, help-json, mcp)',
      exited(helpJson.p, 0) && manifest.package === PACKAGE.name && entries.length === 207 && ['docs', 'help-json', 'mcp'].every((name) => entries.includes(name)),
      `${manifest.package ?? 'no package'}, ${entries.length} entries`,
    ],
    [
      'create-mail-draft exits 1 and names @ask-marcel/office-write',
      exited(write.p, 1) && (envelope(write.out).error ?? '').includes('@ask-marcel/office-write'),
      exitNote(write.p),
    ],
    ['login exits 1 and names @ask-marcel/office-auth', exited(login.p, 1) && (envelope(login.out).error ?? '').includes('@ask-marcel/office-auth'), exitNote(login.p)],
    [
      'a bad tier token answers env_token_invalid',
      exited(bad.p, 1) && envelope(bad.out).errorCode === 'env_token_invalid',
      `${exitNote(bad.p)}, ${envelope(bad.out).errorCode ?? ''}`,
    ],
    [
      'an unknown flag exits 1',
      exited(unknown.p, 1) && envelope(unknown.out).errorCode === 'commander.unknownOption',
      `${exitNote(unknown.p)}, ${envelope(unknown.out).errorCode ?? ''}`,
    ],
    ['a CLI run leaves only the update-notifier configstore under HOME', stray.length === 0, stray.length === 0 ? 'clean' : stray.join(', ')],
  ];
};

const probeReadMcp = (rt: string): ReadonlyArray<Check> => {
  const served = runRead(rt, ['mcp'], {}, `${JSON.stringify(MCP_INIT)}\n`);
  const init = served.out
    .split('\n')
    .filter((line) => line.trim() !== '')
    .map((line): { id?: number; result?: { serverInfo?: { name?: string } } } => {
      try {
        return JSON.parse(line) as { id?: number; result?: { serverInfo?: { name?: string } } };
      } catch {
        return {};
      }
    })
    .find((frame) => frame.id === 1);
  const server = init?.result?.serverInfo?.name ?? '';
  const stray = served.files.filter((path) => !BUN_CACHE.test(path));
  return [
    ['the MCP server is ask-marcel-office-read', exited(served.p, 0) && server === 'ask-marcel-office-read', `${exitNote(served.p)}, ${server}`],
    ['an mcp run leaves nothing under HOME', stray.length === 0, stray.length === 0 ? 'clean' : stray.join(', ')],
  ];
};

/*
 * The library of a caller (F3), as the write smoke checks write's: the exports
 * by name, and a token from the auth package's token helper. The library picks
 * the file system and the process runner of the runtime that loads it, so only
 * a run under node proves the node pick. The helper is a stub that prints a
 * token it builds at run time, so no token-shaped text is committed.
 */
const HELPER_STUB = join(READ_HOMES, 'token-helper.mjs');
writeFileSync(
  HELPER_STUB,
  `#!/usr/bin/env node
const part = (value) => Buffer.from(JSON.stringify(value)).toString('base64url');
const expiresOn = Math.floor(Date.now() / 1000) + 3600;
const accessToken = [part({ alg: 'none', typ: 'JWT' }), part({ aud: 'https://graph.microsoft.com', exp: expiresOn, from: 'the stub helper' }), 'x'].join('.');
process.stdout.write(JSON.stringify({ accessToken, expiresOn }) + '\\n');
`
);
chmodSync(HELPER_STUB, 0o755);
const READ_EXPORTS = 'commands,createEnvThenHelperTokenSource,createReadClient,createReadGraph 204';
const LIBRARY_PROBE = `const m = await import('${PACKAGE.name}');
const r = await m.createEnvThenHelperTokenSource().graphToken('basic');
const signer = r.ok ? JSON.parse(Buffer.from(String(r.value).split('.')[1], 'base64url').toString()).from : 'error ' + r.error.code;
process.stdout.write(Object.keys(m).sort().join(',') + ' ' + Object.keys(m.commands).length + ' signed by ' + signer);`;

const probeReadLibrary = (rt: string): ReadonlyArray<Check> => {
  const home = mkdtempSync(join(READ_HOMES, 'home-'));
  const p = spawnSync(rt, rt === 'node' ? ['--input-type=module', '-e', LIBRARY_PROBE] : ['-e', LIBRARY_PROBE], {
    cwd: ROOT,
    env: { ...CALLER_ENV, ASKMARCEL_TOKEN_COMMAND: HELPER_STUB, HOME: home, USERPROFILE: home },
    stdio: ['ignore', 'pipe', 'pipe'],
    timeout: PROBE_TIMEOUT_MS,
  });
  const out = p.stdout?.toString() ?? '';
  return [
    [
      'the library loads by name with its exports and the 204 commands, and signs through the token helper',
      exited(p, 0) && out === `${READ_EXPORTS} signed by the stub helper`,
      out || (p.stderr?.toString().slice(0, 200) ?? exitNote(p)),
    ],
  ];
};

const report = (title: string, checks: ReadonlyArray<Check>): void => {
  console.log(`=== ${title} ===`);
  for (const [label, passed, note] of checks) {
    if (!passed) fails++;
    console.log(`  ${passed ? '✓' : '✗'} ${label} (${note})`);
  }
};
report('read: the shipped files', probeShipped());
for (const rt of ['node', 'bun']) {
  report(`read: the bin @ ${rt}`, probeReadCli(rt));
  report(`read: mcp @ ${rt}`, probeReadMcp(rt));
  report(`read: library @ ${rt}`, probeReadLibrary(rt));
}
rmSync(READ_HOMES, { recursive: true, force: true });

console.log(`\n${fails === 0 ? 'ALL CONVERTERS + MCP + READ BIN OK under node + bun ✓' : `!! ${fails} bundle FAILURES`}`);
process.exit(fails > 0 ? 1 : 0);
