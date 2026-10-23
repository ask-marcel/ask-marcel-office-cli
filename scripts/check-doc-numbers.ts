#!/usr/bin/env bun
/*
 * Prose-number gate: every hardcoded command count in the docs must match
 * what the registry actually holds.
 *
 * Four numeric claims drifted during the 2.4.0 cycle and a fifth was wrong on
 * main for two releases (`docs/USAGE.md` promised "the 180 READ commands"
 * while the registry held 188, and `README.md` said 188 three sections
 * earlier). `mcp-shell.ts` already derives its counts and carries a comment saying
 * the prose must be derived too; nothing enforced it until this gate.
 *
 * The counts come from the registry and the rendered manifest, never from a
 * second literal. A claim whose anchor text no longer matches anything is a
 * FAILURE, not a pass: a reworded sentence would otherwise disable its own
 * check silently, which is the failure mode this gate exists to prevent.
 *
 * The generated command reference (the `--manifest` JSON and the
 * `--commands-doc` Markdown) must also carry the bins filled in: a `{bin}`,
 * `{auth-bin}`, `{read-bin}` or `{write-bin}` left in either is a failure,
 * because gen-docs fills them and a reader cannot run a placeholder.
 *
 * Exit codes:
 *   0  every claim matches, every anchor was found, and no placeholder is left
 *   1  a claim disagrees with the registry, an anchor matched nothing, or a
 *      generated document holds a bin placeholder
 *
 *   2  an argument is missing, a --package exports no `commandPackage`, or
 *      two packages hold one command name with different entries
 *
 *   --selftest  proves the matcher rejects a wrong number, that a placeholder
 *               is refused and a filled bin passes, and that the join of the
 *               packages keeps every command and refuses a conflict
 *               (see rule 15.10: a gate only ever seen green is a hypothesis)
 *
 * The package and the three documents are arguments, not imports, so one
 * script serves every package of the split (plan "Scripts ownership");
 * `bun run check:docs` passes the root package and the write package, which
 * the root documents together until the docs split (see command-packages.ts):
 *
 *   bun run scripts/check-doc-numbers.ts --package src/composition/single-package.ts \
 *     --package write/src/composition/write-package.ts \
 *     --readme README.md --commands-doc docs/COMMANDS.md --usage docs/USAGE.md \
 *     --manifest docs/commands.json
 *
 * `--package` names a module that exports `commandPackage` (a CommandPackage).
 * It can come more than once: the registries are joined. Paths are relative to
 * the working directory.
 */

import { parseArgs } from 'node:util';
import { buildManifest } from '@ask-marcel/office-core/use-cases/commands/docs-builders.ts';
import { loadPackages, mergePackages } from './command-packages.ts';

const USAGE =
  'usage: bun run scripts/check-doc-numbers.ts --package <module exporting commandPackage> [--package <another>] --readme <README.md> --commands-doc <COMMANDS.md> --usage <USAGE.md> --manifest <commands.json> [--selftest]';

const { values: args } = parseArgs({
  options: {
    package: { type: 'string', multiple: true },
    readme: { type: 'string' },
    'commands-doc': { type: 'string' },
    usage: { type: 'string' },
    manifest: { type: 'string' },
    selftest: { type: 'boolean' },
  },
});
if (args.package === undefined || args.readme === undefined || args['commands-doc'] === undefined || args.usage === undefined || args.manifest === undefined) {
  console.error(`check-doc-numbers: ${USAGE}`);
  process.exit(2);
}

const loaded = await loadPackages(args.package);
if (!loaded.ok) {
  console.error(`check-doc-numbers: ${loaded.error} ${USAGE}`);
  process.exit(2);
}
const merged = mergePackages(loaded.value);
if (!merged.ok) {
  console.error(`check-doc-numbers: ${merged.error}`);
  process.exit(2);
}
const commandPackage = merged.value;

/** The documents the claims sit in, by role: their paths are arguments. */
const DOCS = { readme: args.readme, commandsDoc: args['commands-doc'], usage: args.usage, manifest: args.manifest } as const;

/** The generated documents, which must hold no bin placeholder. */
const GENERATED = [DOCS.manifest, DOCS.commandsDoc] as const;

/** A bin placeholder that gen-docs fills: `{bin}` or a role form. */
const PLACEHOLDER = /\{(?:(?:auth|read|write)-)?bin\}/g;

type Claim = {
  /** Path of the document holding the claim, relative to the working directory. */
  readonly file: string;
  /** What the sentence is about, for the failure message. */
  readonly label: string;
  /** Anchor text with one capture group per number, in `expected` order. */
  readonly pattern: RegExp;
  readonly expected: ReadonlyArray<number>;
};

const entries = Object.entries(commandPackage.registry);
const manifest = buildManifest(commandPackage, '0.0.0');

// The package's lifecycle stubs (in the single package `login`, `status`,
// `logout`, `docs`, `help-json`, `mcp`) each carry `graphMethod: 'GET'` although
// none calls Graph, so counting verbs or categories over the whole manifest
// inflates GET by their number and invents a twelfth category. Every verb and category claim in the docs is about the Graph
// surface, which is what `docs/commands.json` holds: the manifest minus these.
const graphCommands = manifest.commands.filter((entry) => entry.category !== 'lifecycle');

const total = entries.length;
const write = entries.filter(([, command]) => command.meta.effect !== 'read').length;
const read = total - write;
const draft = entries.filter(([, command]) => command.meta.effect === 'draft').length;
const transientUpload = entries.filter(([, command]) => command.meta.effect === 'transient-upload').length;
const categories = new Set(graphCommands.map((entry) => entry.category)).size;

// The manifest carries the lifecycle stubs the registry does not; the COMMANDS.md
// gap sentence explains that difference and must track both sides of it.
const manifestTotal = manifest.commands.length;

// The README splits the surface by effect first and verb second: the reads are
// GET or a read-only POST (the searches and the free/busy lookup), and the writes
// are counted apart whatever their verb (three drafts are POST, and the PDF
// conversions are declared GET although they upload a temporary file).
const readEntries = graphCommands.filter((entry) => entry.effect === 'read');
const readGet = readEntries.filter((entry) => entry.graphMethod === 'GET').length;
const readOnlyPost = readEntries.filter((entry) => entry.graphMethod === 'POST').length;

const CLAIMS: ReadonlyArray<Claim> = [
  { file: DOCS.readme, label: 'command total in the nav and deep-docs links', pattern: /All (\d+) commands/g, expected: [total] },
  {
    file: DOCS.readme,
    label: 'safety breakdown by HTTP verb',
    pattern: /The (\d+) commands break down as (\d+) GET reads, (\d+) read-only POST[^.]*, and (\d+) writes: (\d+) mail-draft operations and (\d+) PDF conversions/,
    expected: [total, readGet, readOnlyPost, write, draft, transientUpload],
  },
  { file: DOCS.readme, label: 'run-command payload in the MCP tool table', pattern: /The (\d+) \*\*read\*\* commands/, expected: [read] },
  {
    file: DOCS.readme,
    label: 'run-write-command payload in the MCP tool table',
    pattern: /The (\d+) \*\*write\*\* commands: (\d+) mail-draft operations and (\d+) PDF conversions/,
    expected: [write, draft, transientUpload],
  },
  { file: DOCS.commandsDoc, label: 'header count and category count', pattern: /All (\d+) commands across (\d+) categories/, expected: [total, categories] },
  { file: DOCS.commandsDoc, label: 'help-json manifest total', pattern: /manifest total \((\d+)\)/, expected: [manifestTotal] },
  { file: DOCS.commandsDoc, label: 'the sentence explaining the lifecycle gap', pattern: /a (\d+)-vs-(\d+) gap/, expected: [total, manifestTotal] },
  { file: DOCS.usage, label: 'run-command payload in the gateway listing', pattern: /the (\d+) READ commands/, expected: [read] },
  { file: DOCS.usage, label: 'run-write-command payload in the gateway listing', pattern: /the (\d+) WRITE commands \((\d+) drafts, (\d+) PDF\)/, expected: [write, draft, transientUpload] },
  {
    file: DOCS.usage,
    label: 'write-command counts in the MCP notes',
    pattern: /carries the (\d+) write commands\s+\((\d+) mail-draft operations and (\d+) PDF conversions\)/,
    expected: [write, draft, transientUpload],
  },
];

/** Line number of a match, so a failure is clickable rather than a hunt. */
const lineOf = (text: string, index: number): number => text.slice(0, index).split('\n').length;

const checkClaim = (text: string, claim: Claim): ReadonlyArray<string> => {
  const failures: Array<string> = [];
  const pattern = new RegExp(claim.pattern.source, claim.pattern.flags.includes('g') ? claim.pattern.flags : `${claim.pattern.flags}g`);
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
  const claim = CLAIMS[1];
  if (claim === undefined) {
    console.error('check-doc-numbers: selftest cannot run, the claim table is empty');
    return 1;
  }
  const wrong = `The ${total + 1} commands break down as ${readGet} GET reads, ${readOnlyPost} read-only POST (three searches and a free/busy lookup), and ${write} writes: ${draft} mail-draft operations and ${transientUpload} PDF conversions.`;
  const rejected = checkClaim(wrong, claim);
  const missing = checkClaim('a document that never mentions the counts at all', claim);

  if (rejected.length === 0) {
    console.error('check-doc-numbers: SELFTEST FAILED — a wrong total was accepted');
    return 1;
  }
  if (missing.length === 0) {
    console.error('check-doc-numbers: SELFTEST FAILED — a missing anchor was accepted');
    return 1;
  }
  if (checkPlaceholders('selftest.md', 'run `{read-bin} list-drives` first').length === 0) {
    console.error('check-doc-numbers: SELFTEST FAILED — a bin placeholder was accepted');
    return 1;
  }
  if (checkPlaceholders('selftest.md', 'run `ask-marcel-office list-drives` first').length > 0) {
    console.error('check-doc-numbers: SELFTEST FAILED — a filled bin was refused');
    return 1;
  }
  const joinFailure = selftestJoin();
  if (joinFailure !== undefined) {
    console.error(`check-doc-numbers: SELFTEST FAILED — ${joinFailure}`);
    return 1;
  }
  console.log(
    `check-doc-numbers: selftest passed (a wrong total, a lost anchor and a bin placeholder are rejected, a filled bin passes; the join keeps every command and refuses a conflict).`
  );
  return 0;
};

const run = async (): Promise<number> => {
  if (args.selftest === true) return runSelftest();

  const failures: Array<string> = [];
  for (const claim of CLAIMS) {
    const text = await Bun.file(claim.file).text();
    failures.push(...checkClaim(text, claim));
  }
  for (const file of GENERATED) failures.push(...checkPlaceholders(file, await Bun.file(file).text()));

  if (failures.length > 0) {
    console.error('check-doc-numbers: the docs disagree with the registry or hold a bin placeholder.\n');
    for (const failure of failures) console.error(`  ${failure}`);
    console.error(`\nRegistry now holds: ${total} commands (${read} read, ${write} write: ${draft} draft, ${transientUpload} transient-upload), ${readGet} GET reads, ${readOnlyPost} read-only POST, ${categories} categories, ${manifestTotal} in the help-json manifest.`);
    return 1;
  }

  console.log(
    `check-doc-numbers: ${CLAIMS.length} claims match the registry (${total} commands, ${read} read, ${write} write, ${readGet} GET reads, ${categories} categories), and the ${GENERATED.length} generated documents hold no bin placeholder.`
  );
  return 0;
};

process.exit(await run());
