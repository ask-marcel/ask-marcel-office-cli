#!/usr/bin/env bun
/*
 * Prose-number gate: every hardcoded command count in the docs must match
 * what the registry actually holds.
 *
 * Four numeric claims drifted during the 2.4.0 cycle and a fifth was wrong on
 * main for two releases (`docs/USAGE.md` promised "the 180 READ commands"
 * while the registry held 188, and `README.md` said 188 three sections
 * earlier). `mcp.ts` already derives its counts and carries a comment saying
 * the prose must be derived too; nothing enforced it until this gate.
 *
 * The counts come from the registry and the rendered manifest, never from a
 * second literal. A claim whose anchor text no longer matches anything is a
 * FAILURE, not a pass: a reworded sentence would otherwise disable its own
 * check silently, which is the failure mode this gate exists to prevent.
 *
 * Exit codes:
 *   0  every claim matches, and every anchor was found
 *   1  a claim disagrees with the registry, or an anchor matched nothing
 *
 *   --selftest  proves the matcher rejects a wrong number (see rule 15.10:
 *               a gate only ever seen green is a hypothesis)
 */

import { commands } from '../src/use-cases/commands/index.ts';
import { buildManifest } from '../src/use-cases/commands/docs.ts';

type Claim = {
  /** Repo-relative path of the document holding the claim. */
  readonly file: string;
  /** What the sentence is about, for the failure message. */
  readonly label: string;
  /** Anchor text with one capture group per number, in `expected` order. */
  readonly pattern: RegExp;
  readonly expected: ReadonlyArray<number>;
};

const entries = Object.entries(commands);
const manifest = buildManifest(commands, 'ask-marcel-office-cli', '0.0.0');

// The six lifecycle stubs (`login`, `logout`, `update`, `docs`, `help-json`,
// `mcp`) each carry `graphMethod: 'GET'` although none calls Graph, so counting
// verbs or categories over the whole manifest inflates GET by six and invents a
// twelfth category. Every verb and category claim in the docs is about the Graph
// surface, which is what `docs/commands.json` holds: the manifest minus these.
const graphCommands = manifest.commands.filter((entry) => entry.category !== 'lifecycle');

const total = entries.length;
const write = entries.filter(([, command]) => command.meta.mutates === true).length;
const read = total - write;
const get = graphCommands.filter((entry) => entry.graphMethod === 'GET').length;
const post = graphCommands.filter((entry) => entry.graphMethod === 'POST').length;
const categories = new Set(graphCommands.map((entry) => entry.category)).size;

// The manifest carries the lifecycle stubs the registry does not; the COMMANDS.md
// gap sentence explains that difference and must track both sides of it.
const manifestTotal = manifest.commands.length;

// The draft commands are the writes, and three of the four are POST. The README
// counts POST twice over: once as the read-only POSTs (the searches and the
// free/busy lookup) and once as the mail-draft operations, so the read-only
// figure is the POSTs that are not drafts.
const draftPosts = graphCommands.filter((entry) => entry.graphMethod === 'POST' && commands[entry.name]?.meta.mutates === true).length;
const readOnlyPost = post - draftPosts;

const CLAIMS: ReadonlyArray<Claim> = [
  { file: 'README.md', label: 'command total in the nav and deep-docs links', pattern: /All (\d+) commands/g, expected: [total] },
  {
    file: 'README.md',
    label: 'safety breakdown by HTTP verb',
    pattern: /The (\d+) commands break down as (\d+) GET, (\d+) read-only POST[^.]*, and (\d+) mail-draft operations/,
    expected: [total, get, readOnlyPost, write],
  },
  { file: 'README.md', label: 'run-command payload in the MCP tool table', pattern: /The (\d+) \*\*read\*\* commands/, expected: [read] },
  { file: 'README.md', label: 'run-write-command payload in the MCP tool table', pattern: /The (\d+) mail-draft \*\*write\*\* commands/, expected: [write] },
  { file: 'docs/COMMANDS.md', label: 'header count and category count', pattern: /All (\d+) commands across (\d+) categories/, expected: [total, categories] },
  { file: 'docs/COMMANDS.md', label: 'help-json manifest total', pattern: /manifest total \((\d+)\)/, expected: [manifestTotal] },
  { file: 'docs/COMMANDS.md', label: 'the sentence explaining the lifecycle gap', pattern: /a (\d+)-vs-(\d+) gap/, expected: [total, manifestTotal] },
  { file: 'docs/USAGE.md', label: 'run-command payload in the gateway listing', pattern: /the (\d+) READ commands/, expected: [read] },
  { file: 'docs/USAGE.md', label: 'run-write-command payload in the gateway listing', pattern: /the (\d+) mail-draft WRITE commands/, expected: [write] },
  { file: 'docs/USAGE.md', label: 'draft-command count in the MCP notes', pattern: /carries the (\d+) draft commands/, expected: [write] },
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

const runSelftest = (): number => {
  const claim = CLAIMS[1];
  if (claim === undefined) {
    console.error('check-doc-numbers: selftest cannot run, the claim table is empty');
    return 1;
  }
  const wrong = `The ${total + 1} commands break down as ${get} GET, ${readOnlyPost} read-only POST (three searches and a free/busy lookup), and ${write} mail-draft operations.`;
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
  console.log(`check-doc-numbers: selftest passed (a wrong total and a lost anchor are both rejected).`);
  return 0;
};

const run = async (): Promise<number> => {
  if (process.argv.includes('--selftest')) return runSelftest();

  const failures: Array<string> = [];
  for (const claim of CLAIMS) {
    const text = await Bun.file(claim.file).text();
    failures.push(...checkClaim(text, claim));
  }

  if (failures.length > 0) {
    console.error('check-doc-numbers: the docs disagree with the registry.\n');
    for (const failure of failures) console.error(`  ${failure}`);
    console.error(`\nRegistry now holds: ${total} commands (${read} read, ${write} write), ${get} GET, ${readOnlyPost} read-only POST, ${categories} categories, ${manifestTotal} in the help-json manifest.`);
    return 1;
  }

  console.log(`check-doc-numbers: ${CLAIMS.length} claims match the registry (${total} commands, ${read} read, ${write} write, ${get} GET, ${categories} categories).`);
  return 0;
};

process.exit(await run());
