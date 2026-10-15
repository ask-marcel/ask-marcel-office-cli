#!/usr/bin/env bun
/*
 * Proves that scripts/coverage-preload-tiers.ts can make the coverage gate
 * fail (rule 15.10: a gate only ever seen green is a hypothesis).
 *
 * In a temporary folder laid out like a workspace folder, a tested domain file,
 * an untested infra file and a type-only port, `bun test --coverage` with the
 * preload must list the untested file at 0% and must not load the port.
 * Without the preload, the untested file would be absent from the table.
 *
 * Exit codes: 0 the preload lists the untested file and skips the port; 1 it
 * does not.
 */

import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const FILES: Readonly<Record<string, string>> = {
  'src/domain/tested.ts': 'export const tested = (): number => 1;\n',
  'src/domain/tested.test.ts': "import { expect, it } from 'bun:test';\nimport { tested } from './tested.ts';\nit('runs', () => expect(tested()).toBe(1));\n",
  'src/infra/untested.ts': 'export const untested = (): number => 2;\n',
  'src/use-cases/ports/port.ts': 'export type Port = () => number;\nthrow new Error("a port was loaded");\n',
};

const folder = mkdtempSync(join(tmpdir(), 'coverage-preload-tiers-'));
for (const [path, text] of Object.entries(FILES)) {
  mkdirSync(join(folder, path, '..'), { recursive: true });
  writeFileSync(join(folder, path), text);
}

const proc = Bun.spawnSync(['bun', 'test', '--coverage', '--preload', join(import.meta.dir, 'coverage-preload-tiers.ts')], { cwd: folder, stdout: 'pipe', stderr: 'pipe' });
rmSync(folder, { recursive: true, force: true });
const report = `${proc.stdout.toString()}\n${proc.stderr.toString()}`;

const untestedRow = report.split('\n').find((line) => line.includes('src/infra/untested.ts'));
const failures = [
  ...(proc.exitCode === 0 ? [] : [`bun test exited ${proc.exitCode} (did the preload load a port?)`]),
  ...(untestedRow?.includes('0.00') ? [] : ['the untested infra file is not in the table at 0%']),
];

if (failures.length > 0) {
  for (const failure of failures) process.stderr.write(`coverage-preload-tiers: SELFTEST FAILED, ${failure}\n`);
  process.stderr.write(report);
  process.exit(1);
}
process.stdout.write('coverage-preload-tiers: selftest passed (an untested tier file shows at 0%, a port is not loaded).\n');
