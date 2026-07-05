#!/usr/bin/env bun
/*
 * QA Phase C — offline parameter-validation matrix (docs/QA-PLAYBOOK.md §C).
 *
 * For EVERY command, exercises four validation cases that all fire BEFORE any
 * Graph call (strictly offline against the built `dist/cli.js` — the honest
 * bundle, not the source):
 *   - unknown-flag              → must reject (unknown option OR, on a command
 *                                 with required flags, the actionable
 *                                 missing-required message — both are correct)
 *   - missing each required     → error must name the omitted flag
 *   - empty "" each required    → clear validation error, never a silent no-op
 *   - duplicate each required   → rejected by the no-repeat parser
 *
 * Run: `bun scripts/qa-param-matrix.ts` (no login needed). Exit non-zero on any
 * real failure. Replaces .claude/qa-reports/phase-c-matrix.ts, whose unknown-flag
 * classifier false-flagged every command that also has a required option.
 */
const CLI = 'dist/cli.js';
const manifest = JSON.parse(await Bun.file('docs/commands.json').text());

type Case = { cmd: string; kind: string; argv: string[]; expectFlag?: string };
const cases: Case[] = [];
for (const c of manifest.commands) {
  const req = (c.options || []).filter((o: any) => o.required);
  // Omitting a required flag drops the WHOLE option (flag + value); a dangling
  // `--flag` with no value would confuse commander into a wrong error.
  const filled = (skip?: string, empty?: string): string[] =>
    req.filter((o: any) => o.name !== skip).flatMap((o: any) => [`--${o.name}`, o.name === empty ? '' : 'x']);
  cases.push({ cmd: c.name, kind: 'unknown', argv: [...filled(), '--definitely-not-a-real-flag', 'x'] });
  for (const o of req) {
    cases.push({ cmd: c.name, kind: `missing:${o.name}`, argv: filled(o.name), expectFlag: o.name });
    cases.push({ cmd: c.name, kind: `empty:${o.name}`, argv: filled(undefined, o.name) });
    cases.push({ cmd: c.name, kind: `dup:${o.name}`, argv: [...filled(), `--${o.name}`, 'y'] });
  }
}

// Async spawn so the worker pool genuinely parallelizes (spawnSync would block
// each worker and serialize all 800+ cases — minutes instead of seconds).
const run = async (argv: string[]): Promise<{ ok: boolean | null; msg: string }> => {
  const p = Bun.spawn(['node', CLI, ...argv, '--output', 'json'], { stdout: 'pipe', stderr: 'pipe' });
  const [out, errout] = await Promise.all([new Response(p.stdout).text(), new Response(p.stderr).text()]);
  await p.exited;
  try { const d = JSON.parse(out); return { ok: d.ok === true, msg: String(d.error || '') }; }
  catch { return { ok: null, msg: (out + errout).slice(0, 160) }; }
};

const POOL = 12;
const results: (Case & { ok: boolean | null; msg: string })[] = [];
let idx = 0;
const worker = async (): Promise<void> => { while (idx < cases.length) { const c = cases[idx++]; results.push({ ...c, ...(await run([c.cmd, ...c.argv])) }); } };
await Promise.all(Array.from({ length: POOL }, worker));

const fails: typeof results = [];
for (const r of results) {
  let pass = false;
  if (r.kind === 'unknown') pass = r.ok === false && /unknown|not a real flag|too many|unexpected|invalid|required option/i.test(r.msg);
  else if (r.kind.startsWith('missing:')) pass = r.ok === false && new RegExp((r.expectFlag || '').replace(/-/g, '[- ]'), 'i').test(r.msg);
  else pass = r.ok === false; // empty + dup: any clear rejection (never ok:true)
  if (!pass) fails.push(r);
}
const byKind: Record<string, { total: number; fail: number }> = {};
for (const r of results) { const k = r.kind.split(':')[0]; (byKind[k] ??= { total: 0, fail: 0 }).total++; }
for (const f of fails) byKind[f.kind.split(':')[0]].fail++;

console.log('=== Phase C offline param-validation matrix ===');
console.log('cases:', results.length, '| by kind (total/fail):', JSON.stringify(byKind));
console.log('\nFAILURES:', fails.length);
for (const f of fails.slice(0, 60)) console.log(`  [${f.cmd}] ${f.kind} ok=${f.ok} msg=${JSON.stringify(f.msg.slice(0, 120))}`);
process.exit(fails.length > 0 ? 1 : 0);
