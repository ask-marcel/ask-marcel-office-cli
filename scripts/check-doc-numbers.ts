#!/usr/bin/env bun
/*
 * Prose-number gate: each command count in the docs of a package must match
 * what the registry of that package holds.
 *
 * Four numeric claims drifted during the 2.4.0 cycle and a fifth was wrong on
 * main for two releases (`docs/USAGE.md` promised "the 180 READ commands"
 * while the registry held 188, and `README.md` said 188 three sections
 * earlier). `mcp-shell.ts` already derives its counts and carries a comment saying
 * the prose must be derived too; nothing enforced it until this gate.
 *
 * Since phase 7 of the package split, each package folder runs this gate on
 * its own README.md, USAGE.md and COMMANDS.md, after `gen-docs.ts --check`, in
 * its `check:docs` script. A doc states only the counts of its own package.
 * The counts come from the registry and the rendered manifest, never from a
 * second literal.
 *
 * Each claim is an anchored sentence with a role. The read claims run when the
 * registry has read commands, the write claims when it has write commands, the
 * registry claims (the two count sentences of COMMANDS.md) when it has
 * commands, and the "lists N entries" claim for every package. A claim that
 * runs and whose anchor matches nothing is a FAILURE, not a pass: a reworded
 * sentence must not disable its own check. A space in an anchor matches any
 * run of white space, so a line break in a sentence does not hide it.
 *
 * The count sweep: a digit count before "commands" or "entries" (with at most
 * two words between) in README.md or USAGE.md that no claim checks is a
 * failure. So a new count sentence needs an anchored form or a new claim.
 *
 * The generated command reference (the `--manifest` JSON and the
 * `--commands-doc` Markdown) must also carry the bins filled in: a `{bin}`,
 * `{auth-bin}`, `{read-bin}` or `{write-bin}` left in either is a failure,
 * because gen-docs fills them and a reader cannot run a placeholder.
 *
 * Exit codes:
 *   0  every claim matches, every anchor was found, and no placeholder is left
 *   1  a claim disagrees with the registry, an anchor matched nothing, a
 *      count in README.md or USAGE.md has no claim, or a generated document
 *      holds a bin placeholder
 *
 *   2  an argument is missing, a --package exports no package, or two
 *      packages hold one command name with different entries
 *
 *   --selftest  proves the matcher rejects a wrong number and a lost anchor
 *               and finds a sentence with a line break, that a claim of an
 *               absent role does not run, that the sweep refuses a count with
 *               no claim and passes a checked one, that a placeholder is
 *               refused and a filled bin passes, and that the join of the
 *               packages keeps every command and refuses a conflict (see rule
 *               15.10: a gate only ever seen green is a hypothesis). It needs
 *               `--package` only.
 *
 * The package and the three documents are arguments, not imports, so one
 * script serves every package of the split (plan "Scripts ownership"). Each
 * package folder runs it from that folder, for example read/:
 *
 *   bun run ../scripts/check-doc-numbers.ts --package src/composition/single-package.ts \
 *     --readme README.md --usage USAGE.md --commands-doc COMMANDS.md --manifest dist/commands.json
 *
 * `--package` names a module that exports `commandPackage` (a CommandPackage),
 * or the export that `--export` names (auth: `authPackage`). `--manifest` is
 * needed only when the registry has commands. Paths are relative to the
 * working directory.
 */

import { parseArgs } from 'node:util';
import { buildManifest } from '@ask-marcel/office-core/use-cases/commands/docs-builders.ts';
import { loadDocumentedPackage, mergePackages } from './command-packages.ts';

const USAGE =
  'usage: bun run ../scripts/check-doc-numbers.ts --package <module exporting commandPackage> [--export <name>] --readme <README.md> --usage <USAGE.md> --commands-doc <COMMANDS.md> [--manifest <commands.json>] | --package <module> --selftest';

const { values: args } = parseArgs({
  options: {
    package: { type: 'string', multiple: true },
    export: { type: 'string' },
    readme: { type: 'string' },
    'commands-doc': { type: 'string' },
    usage: { type: 'string' },
    manifest: { type: 'string' },
    selftest: { type: 'boolean' },
  },
});
const stop = (message: string): never => {
  console.error(`check-doc-numbers: ${message}`);
  process.exit(2);
};

const loaded = await loadDocumentedPackage(args.package ?? stop(USAGE), args.export);
const commandPackage = loaded.ok ? loaded.value : stop(`${loaded.error} ${USAGE}`);
const entries = Object.entries(commandPackage.registry);
const missing = args.readme === undefined || args['commands-doc'] === undefined || args.usage === undefined || (entries.length > 0 && args.manifest === undefined);
if (args.selftest !== true && missing) stop(USAGE);

/** The documents the claims sit in, by role: their paths are arguments. */
const DOCS = { readme: args.readme ?? '', commandsDoc: args['commands-doc'] ?? '', usage: args.usage ?? '', manifest: args.manifest } as const;

/** The generated documents, which must hold no bin placeholder. */
const GENERATED = DOCS.manifest === undefined ? [DOCS.commandsDoc] : [DOCS.manifest, DOCS.commandsDoc];

/** A bin placeholder that gen-docs fills: `{bin}` or a role form. */
const PLACEHOLDER = /\{(?:(?:auth|read|write)-)?bin\}/g;

/** When a claim runs: `read` and `write` when the registry has such commands, `registry` when it has any, `every` always. */
type Role = 'read' | 'write' | 'registry' | 'every';

type Claim = {
  readonly role: Role;
  /** Path of the document holding the claim, relative to the working directory. */
  readonly file: string;
  /** What the sentence is about, for the failure message. */
  readonly label: string;
  /** Anchor text with one capture group per number, in `expected` order. A space matches any white space. */
  readonly pattern: RegExp;
  readonly expected: ReadonlyArray<number>;
};

const manifest = buildManifest(commandPackage, '0.0.0');

// The package's lifecycle stubs (in read `docs`, `help-json`, `mcp`) each carry
// `graphMethod: 'GET'` although none calls Graph, so counting verbs or
// categories over the whole manifest inflates GET by their number and invents
// a category. Every verb and category claim in the docs is about the Graph
// surface: the manifest minus these.
const graphCommands = manifest.commands.filter((entry) => entry.category !== 'lifecycle');

const total = entries.length;
const write = entries.filter(([, command]) => command.meta.effect !== 'read').length;
const read = total - write;
const draft = entries.filter(([, command]) => command.meta.effect === 'draft').length;
const transientUpload = entries.filter(([, command]) => command.meta.effect === 'transient-upload').length;
const categories = new Set(graphCommands.map((entry) => entry.category)).size;

// The manifest carries the lifecycle stubs the registry does not; the
// `help-json` sentences state that total.
const manifestTotal = manifest.commands.length;

// The read README splits the reads by verb: GET or a read-only POST (the
// searches and the free/busy lookup). The write README splits the writes by
// effect, whatever their verb (three drafts are POST, and the PDF conversions
// are declared GET although they upload a temporary file).
const readEntries = graphCommands.filter((entry) => entry.effect === 'read');
const readGet = readEntries.filter((entry) => entry.graphMethod === 'GET').length;
const readOnlyPost = readEntries.filter((entry) => entry.graphMethod === 'POST').length;

const CLAIMS: ReadonlyArray<Claim> = [
  { role: 'every', file: DOCS.readme, label: 'help-json entries', pattern: /`help-json` lists (\d+) entries/, expected: [manifestTotal] },
  { role: 'read', file: DOCS.readme, label: 'read command count', pattern: /\b[Tt]he (\d+) read commands\b/, expected: [read] },
  {
    role: 'read',
    file: DOCS.readme,
    label: 'read commands by HTTP verb',
    pattern: /The (\d+) read commands are (\d+) GET reads and (\d+) read-only POST/,
    expected: [read, readGet, readOnlyPost],
  },
  { role: 'read', file: DOCS.usage, label: 'read command count', pattern: /\b[Tt]he (\d+) read commands\b/, expected: [read] },
  { role: 'read', file: DOCS.usage, label: 'run-command payload in the MCP section', pattern: /`run-command` runs the (\d+) read commands/, expected: [read] },
  { role: 'write', file: DOCS.readme, label: 'write command count', pattern: /\b[Tt]he (\d+) write commands\b/, expected: [write] },
  {
    role: 'write',
    file: DOCS.readme,
    label: 'write commands by effect',
    pattern: /the (\d+) write commands: (\d+) mail-draft operations and (\d+) PDF conversions/,
    expected: [write, draft, transientUpload],
  },
  { role: 'write', file: DOCS.usage, label: 'write command count', pattern: /\b[Tt]he (\d+) write commands\b/, expected: [write] },
  { role: 'write', file: DOCS.usage, label: 'run-write-command payload in the MCP section', pattern: /`run-write-command` runs the (\d+) write commands/, expected: [write] },
  { role: 'registry', file: DOCS.commandsDoc, label: 'command and category count', pattern: /The (\d+) commands are in (\d+) categor(?:y|ies)\./, expected: [total, categories] },
  {
    role: 'registry',
    file: DOCS.commandsDoc,
    label: 'help-json entries, commands and lifecycle commands',
    pattern: /`help-json` lists (\d+) entries: the (\d+) commands and the (\d+) lifecycle commands below\./,
    expected: [manifestTotal, total, commandPackage.lifecycle.length],
  },
];

/** The registry counts that decide which claims run. */
type RoleCounts = { readonly read: number; readonly write: number; readonly total: number };

const runs = (claim: Claim, counts: RoleCounts): boolean =>
  claim.role === 'every' || (claim.role === 'read' && counts.read > 0) || (claim.role === 'write' && counts.write > 0) || (claim.role === 'registry' && counts.total > 0);

/** The claims that run for this registry. */
const ACTIVE = CLAIMS.filter((claim) => runs(claim, { read, write, total }));

/** A digit count before "commands" or "entries", with at most two words between. */
const COUNT = /\b\d+\s+(?:[\w-]+\s+){0,2}(?:commands|entries)\b/g;

/** The anchor of a claim, global and with the capture indices: a space matches any white space. */
const anchorOf = (claim: Claim): RegExp => new RegExp(claim.pattern.source.replaceAll(' ', String.raw`\s+`), 'gd');

/** Line number of a match, so a failure is clickable rather than a hunt. */
const lineOf = (text: string, index: number): number => text.slice(0, index).split('\n').length;

const checkClaim = (text: string, claim: Claim): ReadonlyArray<string> => {
  const failures: Array<string> = [];
  const pattern = anchorOf(claim);
  let found = 0;

  for (const match of text.matchAll(pattern)) {
    found += 1;
    const line = lineOf(text, match.index ?? 0);
    claim.expected.forEach((expected, position) => {
      const actual = Number(match[position + 1]);
      if (actual !== expected) failures.push(`${claim.file}:${line}  ${claim.label}: expected ${expected}, found ${actual}`);
    });
  }

  // An anchor that matches nothing means the sentence was reworded and its
  // check quietly stopped running. That is a failure, not a pass.
  if (found === 0) failures.push(`${claim.file}  ${claim.label}: anchor /${claim.pattern.source}/ matched nothing (reworded? update the claim in this gate)`);
  return failures;
};

const checkPlaceholders = (file: string, text: string): ReadonlyArray<string> =>
  [...text.matchAll(PLACEHOLDER)].map((match) => `${file}:${lineOf(text, match.index)}  bin placeholder ${match[0]} left in a generated document (run docs:gen)`);

// The count sweep: each digit count before "commands" or "entries" in a file
// must be a number that a claim of that file checks. Else a new count sentence
// has no check, and the count can drift while the gate passes.
const unmatchedCounts = (file: string, text: string, claims: ReadonlyArray<Claim>): ReadonlyArray<string> => {
  const checked = new Set<number>();
  for (const claim of claims.filter((each) => each.file === file)) {
    for (const match of text.matchAll(anchorOf(claim))) for (const [start] of (match.indices ?? []).slice(1)) checked.add(start);
  }
  return [...text.matchAll(COUNT)]
    .filter((match) => !checked.has(match.index))
    .map((match) => `${file}:${lineOf(text, match.index)}  "${match[0].replaceAll(/\s+/g, ' ')}" is a count that no claim checks (use an anchored sentence, or add a claim)`);
};

// The join: the registry split in two packages (the second with other bins) must
// come back whole with the names of the first, and a name that the second holds
// with another summary must be refused.
const selftestJoin = (): string | undefined => {
  const names = Object.keys(commandPackage.registry);
  const half = Math.ceil(names.length / 2);
  const part = (from: number, to: number): typeof commandPackage.registry => Object.fromEntries(Object.entries(commandPackage.registry).slice(from, to));
  const other = { ...commandPackage, binName: 'other-bin', registry: part(half - 1, names.length) };
  const joined = mergePackages([{ ...commandPackage, registry: part(0, half) }, other]);
  if (!joined.ok || Object.keys(joined.value.registry).length !== names.length || joined.value.binName !== commandPackage.binName) return 'the join lost a command or took the names of the second package';
  const [name, command] = Object.entries(commandPackage.registry)[0] ?? [];
  if (name === undefined || command === undefined) return 'the join cannot be checked on an empty registry';
  const conflict = { ...other, registry: { [name]: { ...command, meta: { ...command.meta, summary: `${command.meta.summary} (changed)` } } } };
  if (mergePackages([commandPackage, conflict]).ok) return 'a command with two different entries was accepted';
  return undefined;
};

const runSelftest = (): number => {
  const claim = CLAIMS[0];
  if (claim === undefined) {
    console.error('check-doc-numbers: selftest cannot run, the claim table is empty');
    return 1;
  }
  const absent = CLAIMS.filter((each) => runs(each, { read: 0, write: 0, total: 0 }));
  const cases = [
    { name: 'a wrong total is refused', passed: checkClaim(`\`help-json\` lists ${manifestTotal + 1} entries.`, claim).length > 0 },
    { name: 'a line break does not hide a right sentence', passed: checkClaim(`\`help-json\` lists\n${manifestTotal} entries.`, claim).length === 0 },
    { name: 'a lost anchor is refused', passed: checkClaim('a document that never mentions the counts at all', claim).length > 0 },
    { name: 'only the every-package claim runs for an empty registry', passed: absent.length === 1 && absent[0] === claim },
    { name: 'each claim runs when every role is present', passed: CLAIMS.every((each) => runs(each, { read: 1, write: 1, total: 2 })) },
    { name: 'the sweep passes a checked count', passed: unmatchedCounts(claim.file, `\`help-json\` lists ${manifestTotal} entries.`, absent).length === 0 },
    { name: 'the sweep refuses a count with no claim', passed: unmatchedCounts(claim.file, 'It has the 12\nmail commands.', absent).length === 1 },
    { name: 'a bin placeholder is refused', passed: checkPlaceholders('selftest.md', 'run `{read-bin} list-drives` first').length > 0 },
    { name: 'a filled bin passes', passed: checkPlaceholders('selftest.md', 'run `ask-marcel-office-read list-drives` first').length === 0 },
    { name: 'the join keeps every command and refuses a conflict', passed: selftestJoin() === undefined },
  ];
  const failed = cases.filter((each) => !each.passed);
  for (const each of failed) console.error(`check-doc-numbers: SELFTEST FAILED — ${each.name}`);
  if (failed.length > 0) return 1;
  console.log(`check-doc-numbers: selftest passed (${cases.map((each) => each.name).join('; ')}).`);
  return 0;
};

const run = async (): Promise<number> => {
  if (args.selftest === true) return runSelftest();

  const failures: Array<string> = [];
  for (const claim of ACTIVE) {
    const text = await Bun.file(claim.file).text();
    failures.push(...checkClaim(text, claim));
  }
  for (const file of [DOCS.readme, DOCS.usage]) failures.push(...unmatchedCounts(file, await Bun.file(file).text(), ACTIVE));
  for (const file of GENERATED) failures.push(...checkPlaceholders(file, await Bun.file(file).text()));

  if (failures.length > 0) {
    console.error(`check-doc-numbers: the docs of ${commandPackage.packageName} disagree with the registry or hold a bin placeholder.\n`);
    for (const failure of failures) console.error(`  ${failure}`);
    console.error(`\nRegistry now holds: ${total} commands (${read} read, ${write} write: ${draft} draft, ${transientUpload} transient-upload), ${readGet} GET reads, ${readOnlyPost} read-only POST, ${categories} categories, ${manifestTotal} in the help-json manifest.`);
    return 1;
  }

  console.log(
    `check-doc-numbers: ${ACTIVE.length} claims match the registry of ${commandPackage.packageName} (${total} commands, ${read} read, ${write} write, ${readGet} GET reads, ${categories} categories), each count in ${DOCS.readme} and ${DOCS.usage} has a claim, and the ${GENERATED.length} generated documents hold no bin placeholder.`
  );
  return 0;
};

process.exit(await run());
