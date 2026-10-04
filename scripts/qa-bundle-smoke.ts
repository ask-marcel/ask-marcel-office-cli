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
 * and `bun`, asserting `ok`. Run: `bun run build && bun scripts/qa-bundle-smoke.ts`.
 * `--package-root <dir>` runs the same probes against another copy of the
 * package (its `dist/` under <dir>), which is how qa-packed-install-smoke.ts
 * checks the tarball as installed outside the repo.
 * Exit non-zero on any bundler-interop failure.
 */
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import type { SpawnSyncReturns } from 'node:child_process';
import {
  buildRichDocx, buildRichXlsx, buildRichPptx, buildRichOdt, buildRichOds, buildRichOdp,
  buildPdfWithText, buildPdfWithImage, buildSampleDoc, buildLegacyXls, buildSampleMsg, buildSampleEml, buildSampleZipArchive,
} from '../src/test-helpers/office-fixtures.ts';

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

// Every run of the CLI records where its token helper lives under HOME, so the
// probes run on a scratch home: a smoke run never repoints the user's own.
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
 * `ask-marcel-office mcp` speaks JSON-RPC over stdout, so it needs a DIFFERENT probe.
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
  if (tools.length !== 6) return { ok: false, note: `expected 6 gateway tools, got ${tools.length}` };
  const callReply = lines.map((l) => JSON.parse(l) as { id?: number; result?: { isError?: boolean; content?: ReadonlyArray<{ text?: string }> } }).find((m) => m.id === 3);
  if (callReply?.result?.isError !== true || !(callReply.result.content?.[0]?.text ?? '').includes('Did you mean')) return { ok: false, note: 'tools/call on a mistyped command did not answer a did-you-mean tool error' };
  return { ok: true, note: `${tools.length} tools, ${lines.length} clean JSON-RPC line(s)` };
};

/*
 * The token helper (`dist/token.js`, and `dist/cli.js token`) prints exactly one
 * JSON line and touches only the home it is given. Each probe runs on a fresh
 * temporary HOME: one holding a cached Graph token (a cache hit, no network),
 * one holding nothing (stdin is not a terminal, so it must fail fast with
 * not_authenticated rather than open a browser). The bundle must not carry the
 * CLI's heavy modules, which is what keeps it near bare runtime startup.
 */
const TOKEN_EXP = Math.floor(Date.now() / 1000) + 3600;
const TOKEN_FIXTURE = `${btoa(JSON.stringify({ alg: 'none' }))}.${btoa(JSON.stringify({ exp: TOKEN_EXP, aud: 'https://graph.microsoft.com' }))}.sig`;
const tokenHome = (withToken: boolean): string => {
  const home = mkdtempSync(join(DIR, 'home-'));
  mkdirSync(join(home, '.ask-marcel'), { recursive: true });
  if (withToken) writeFileSync(join(home, '.ask-marcel', 'token-cache.json'), JSON.stringify({ access_token: TOKEN_FIXTURE, expires_on: TOKEN_EXP, refresh_token: '' }));
  return home;
};
const runToken = (rt: string, args: ReadonlyArray<string>, home: string): { p: SpawnSyncReturns<Buffer>; ms: number } => {
  const started = performance.now();
  const p = spawnSync(rt, args, { cwd: ROOT, env: { ...process.env, HOME: home, USERPROFILE: home }, stdio: ['ignore', 'pipe', 'pipe'], timeout: PROBE_TIMEOUT_MS });
  return { p, ms: Math.round(performance.now() - started) };
};
const oneLine = (p: SpawnSyncReturns<Buffer>): Record<string, unknown> | undefined => {
  const lines = (p.stdout?.toString() ?? '').split('\n').filter((l) => l !== '');
  try {
    return lines.length === 1 ? (JSON.parse(lines[0] ?? '') as Record<string, unknown>) : undefined;
  } catch {
    return undefined;
  }
};
const probeToken = (rt: string): ReadonlyArray<[string, boolean, string]> => {
  const bundle = readFileSync(join(ROOT, 'dist/token.js'), 'utf8');
  const heavy = ['commander', 'winston', 'update-notifier'].filter((name) => bundle.includes(name));
  const warm = tokenHome(true);
  const hit = runToken(rt, ['dist/token.js', '--tier', 'basic'], warm);
  const hitLine = oneLine(hit.p);
  const locator = JSON.parse(readFileSync(join(warm, '.ask-marcel', 'token-helper.json'), 'utf8')) as { entry?: string };
  const viaCli = runToken(rt, ['dist/cli.js', 'token', '--tier', 'basic'], warm);
  const cold = runToken(rt, ['dist/token.js', '--tier', 'basic'], tokenHome(false));
  return [
    ['dist/token.js is a node script without the CLI modules', bundle.startsWith('#!/usr/bin/env node') && heavy.length === 0, heavy.length === 0 ? 'clean' : `carries ${heavy.join(', ')}`],
    ['a cache hit prints the token line', exited(hit.p, 0) && hitLine?.accessToken === TOKEN_FIXTURE && hitLine.expiresOn === TOKEN_EXP, `${exitNote(hit.p)}, ${hit.ms} ms`],
    ['the locator names dist/token.js', locator.entry === join(ROOT, 'dist/token.js'), String(locator.entry)],
    ['cli.js token prints the same line', exited(viaCli.p, 0) && oneLine(viaCli.p)?.accessToken === TOKEN_FIXTURE, `${exitNote(viaCli.p)}, ${viaCli.ms} ms`],
    ['no session fails fast, no browser', exited(cold.p, 1) && oneLine(cold.p)?.errorCode === 'not_authenticated', `${exitNote(cold.p)}, ${cold.ms} ms`],
  ];
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
  console.log(`=== token helper @ ${rt} ===`);
  for (const [label, passed, note] of probeToken(rt)) {
    if (!passed) fails++;
    console.log(`  ${passed ? '✓' : '✗'} ${label} (${note})`);
  }
  console.log(`=== mcp stdio handshake @ ${rt} ===`);
  const m = probeMcp(rt); if (!m.ok) fails++; console.log(`  ${m.ok ? '✓' : '✗'} mcp   -> ${m.note}`);
}
console.log(`\n${fails === 0 ? 'ALL CONVERTERS + TOKEN HELPER + MCP OK under node + bun ✓' : `!! ${fails} bundler-interop FAILURES`}`);
process.exit(fails > 0 ? 1 : 0);
