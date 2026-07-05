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

let fails = 0;
for (const rt of ['node', 'bun']) {
  console.log(`\n=== convert-local-file @ ${rt} ===`);
  for (const [fmt, path] of Object.entries(F)) { const r = probe(rt, 'convert-local-file', path); if (!r.ok) fails++; console.log(`  ${r.ok ? '✓' : '✗'} ${fmt.padEnd(5)} -> ${r.note}`); }
  console.log(`=== extract-local-file-images @ ${rt} ===`);
  for (const [fmt, path] of [['docx', F.docx], ['xlsx', F.xlsx], ['pptx', F.pptx], ['pdf', pdfImg]]) { const r = probe(rt, 'extract-local-file-images', path); if (!r.ok) fails++; console.log(`  ${r.ok ? '✓' : '✗'} ${fmt.padEnd(5)} -> ${r.note}`); }
}
console.log(`\n${fails === 0 ? 'ALL CONVERTERS OK under node + bun ✓' : `!! ${fails} bundler-interop FAILURES`}`);
process.exit(fails > 0 ? 1 : 0);
