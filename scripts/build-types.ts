#!/usr/bin/env bun
/*
 * Builds the declarations of one workspace package (auth/, read/, write/), so
 * that the published .d.ts never names the private core. Run it with the folder
 * as the working directory, after scripts/build-package.ts:
 *
 *   bun run ../scripts/build-types.ts
 *
 * Core is never published, and each bundle inlines it. A .d.ts that kept an
 * import of `@ask-marcel/office-core` (or a path into ../core) would fail in
 * every consumer, and `skipLibCheck` would hide it. So, per file (the approach
 * the package split's phase 2 chose over a declaration bundler):
 *
 *   1. tsc -p tsconfig.build.json emits the declarations of the public entry
 *      (src/index.ts) and of every module it reaches, core's included: the
 *      build tsconfig maps `@ask-marcel/office-core/*` onto ../core/src/* with
 *      `paths`, so core is compiled from source, into dist/.types/.
 *   2. The package's declarations move to dist/, core's to dist/shared/.
 *   3. Each module specifier is resolved against the emitted tree and written
 *      again as a relative `.js` path to its new place: a core import becomes
 *      `./shared/...`, and a `.ts` extension becomes `.js`. A specifier that
 *      reaches another package, or a file that was not emitted, fails the build.
 *
 * scripts/check-packed-package.ts then checks the packed result, and a clean
 * consumer typechecks against it.
 */

import { existsSync, mkdirSync, rmSync } from 'node:fs';
import { basename, dirname, join, relative, resolve } from 'node:path';
import ts from 'typescript';

const RAW = resolve('dist/.types');
const FOLDER = basename(process.cwd());
const CORE_PREFIX = '@ask-marcel/office-core/';

// Where an emitted declaration goes in dist/, or undefined when it is outside
// this package and core.
const relocate = (raw: string): string | undefined => {
  const own = join(RAW, FOLDER, 'src');
  const core = join(RAW, 'core', 'src');
  if (raw.startsWith(`${own}/`)) return resolve('dist', relative(own, raw));
  if (raw.startsWith(`${core}/`)) return resolve('dist', 'shared', relative(core, raw));
  return undefined;
};

const withoutExtension = (path: string): string => path.replace(/\.(d\.)?[cm]?[jt]s$/, '');

// The emitted declaration a specifier of `from` names, or a reason it has none.
const targetOf = (from: string, specifier: string): string | { readonly problem: string } | undefined => {
  if (specifier.startsWith(CORE_PREFIX)) return `${withoutExtension(join(RAW, 'core', 'src', specifier.slice(CORE_PREFIX.length)))}.d.ts`;
  if (specifier.startsWith('@ask-marcel/office-')) return { problem: `names another package (${specifier})` };
  if (!specifier.startsWith('.')) return undefined;
  return `${withoutExtension(resolve(dirname(from), specifier))}.d.ts`;
};

const respecified = (from: string, target: string): string => {
  const path = relative(dirname(from), withoutExtension(target));
  const dotted = path.startsWith('.') ? path : `./${path}`;
  return `${dotted}.js`;
};

const rewrite = async (raw: string, problems: string[]): Promise<void> => {
  const destination = relocate(raw);
  if (destination === undefined) {
    problems.push(`${relative(RAW, raw)}: the public entry reaches a module outside ${FOLDER}/src and core/src`);
    return;
  }
  let text = await Bun.file(raw).text();
  const imports = ts.preProcessFile(text, true, true).importedFiles.toSorted((a, b) => b.pos - a.pos);
  for (const { fileName, pos } of imports) {
    const target = targetOf(raw, fileName);
    if (target === undefined) continue;
    const finalTarget = typeof target === 'string' && existsSync(target) ? relocate(target) : undefined;
    if (typeof target !== 'string' || finalTarget === undefined) {
      problems.push(`${relative(RAW, raw)}: '${fileName}' ${typeof target === 'string' ? 'names no emitted declaration of this package or core' : target.problem}`);
      continue;
    }
    // `pos` is at the opening quote, or at the name itself in some forms.
    const start = text.startsWith(fileName, pos) ? pos : pos + 1;
    text = `${text.slice(0, start)}${respecified(destination, finalTarget)}${text.slice(start + fileName.length)}`;
  }
  mkdirSync(dirname(destination), { recursive: true });
  await Bun.write(destination, text);
};

const emit = Bun.spawnSync(['bun', '--bun', 'x', 'tsc', '-p', 'tsconfig.build.json'], { stdout: 'inherit', stderr: 'inherit' });
if (emit.exitCode !== 0) process.exit(emit.exitCode ?? 1);

const raws = [...new Bun.Glob('**/*.d.ts').scanSync({ cwd: RAW })].map((path) => join(RAW, path));
const problems: string[] = [];
for (const raw of raws) await rewrite(raw, problems);
rmSync(RAW, { recursive: true, force: true });

if (problems.length > 0) {
  for (const problem of problems) process.stderr.write(`build-types: ${problem}\n`);
  process.exit(1);
}
process.stderr.write(`build-types: ${raws.length} declaration file(s) in dist/, core's under dist/shared/\n`);
