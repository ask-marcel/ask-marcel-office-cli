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
 * Exit non-zero on any bundler-interop failure.
 */
import { mkdirSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import {
  buildRichDocx, buildRichXlsx, buildRichPptx, buildRichOdt, buildRichOds, buildRichOdp,
  buildPdfWithText, buildPdfWithImage, buildSampleDoc, buildLegacyXls, buildSampleMsg, buildSampleZipArchive,
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
F.zip = await write('f.zip', await buildSampleZipArchive());
F.csv = await write('f.csv', new TextEncoder().encode('a,b,c\n1,2,3\n'));
const pdfImg = await write('img.pdf', await buildPdfWithImage());

const probe = (rt: string, cmd: string, path: string): { ok: boolean; note: string } => {
  const p = spawnSync(rt, ['dist/cli.js', cmd, '--path', path, '--output', 'json'], { timeout: 60000, maxBuffer: 64 * 1024 * 1024 });
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

const probeMcp = (rt: string): { ok: boolean; note: string } => {
  const p = spawnSync(rt, ['dist/cli.js', 'mcp'], {
    input: `${JSON.stringify(MCP_INIT)}\n${JSON.stringify(MCP_LIST)}\n`,
    timeout: 60000,
    maxBuffer: 64 * 1024 * 1024,
  });
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
  if (tools.length !== 5) return { ok: false, note: `expected 5 gateway tools, got ${tools.length}` };
  return { ok: true, note: `${tools.length} tools, ${lines.length} clean JSON-RPC line(s)` };
};

let fails = 0;
for (const rt of ['node', 'bun']) {
  console.log(`\n=== convert-local-file-to-markdown @ ${rt} ===`);
  for (const [fmt, path] of Object.entries(F)) { const r = probe(rt, 'convert-local-file-to-markdown', path); if (!r.ok) fails++; console.log(`  ${r.ok ? '✓' : '✗'} ${fmt.padEnd(5)} -> ${r.note}`); }
  console.log(`=== extract-local-file-images @ ${rt} ===`);
  for (const [fmt, path] of [['docx', F.docx], ['xlsx', F.xlsx], ['pptx', F.pptx], ['pdf', pdfImg]]) { const r = probe(rt, 'extract-local-file-images', path); if (!r.ok) fails++; console.log(`  ${r.ok ? '✓' : '✗'} ${fmt.padEnd(5)} -> ${r.note}`); }
  console.log(`=== mcp stdio handshake @ ${rt} ===`);
  const m = probeMcp(rt); if (!m.ok) fails++; console.log(`  ${m.ok ? '✓' : '✗'} mcp   -> ${m.note}`);
}
console.log(`\n${fails === 0 ? 'ALL CONVERTERS + MCP OK under node + bun ✓' : `!! ${fails} bundler-interop FAILURES`}`);
process.exit(fails > 0 ? 1 : 0);
