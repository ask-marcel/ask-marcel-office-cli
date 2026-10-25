#!/usr/bin/env bun
/*
 * The 2.8.0 docs stay as npm shows them (package split, phase 7, G2).
 *
 * npm installs `ask-marcel-office-cli` 2.8.0 until phase 8. Its README on npm
 * links `docs/COMMANDS.md`, `docs/USAGE.md`, `docs/commands.json` and
 * `docs/demo.gif` on main. So these four files keep the v2.8.0 bytes. Each
 * Markdown file starts with a banner between the FROZEN-2.8.0 markers, which
 * points at the docs of the three packages. The banner text can change. The
 * text under the banner (after the END marker line and one blank line) must
 * have the sha256 of the v2.8.0 file. The other files must have the sha256 of
 * their v2.8.0 bytes.
 *
 * The check hashes bytes, not text: a UTF-8 decode changes each byte that is
 * not valid UTF-8 to U+FFFD, so a changed byte of the gif could hide.
 *
 *   --selftest  proves that a right file passes, that a changed byte, a
 *               missing banner, a banner with no BEGIN marker, a changed
 *               byte that the UTF-8 decode hides and a missing file are
 *               refused, and that a changed banner text passes (rule 15.10).
 *
 * Exit codes: 0 the four files keep the 2.8.0 bytes; 1 a file is changed or
 * missing, or the selftest failed.
 *
 * Run it from the repository root: bun scripts/check-frozen-docs.ts [--selftest]
 */

type Rule = { readonly sha256: string; readonly banner: boolean };

/** The sha256 of each file at v2.8.0 (`git show v2.8.0:<path> | sha256sum`). */
const FROZEN: Readonly<Record<string, Rule>> = {
  'docs/COMMANDS.md': { sha256: 'a921b787d9b2a7da7dd63bcb31a7688cf6aadba3b4952b294cd57c173ad253a8', banner: true },
  'docs/USAGE.md': { sha256: '3c7947f6d3ca6d68d5b7862582146f45ecd0a112978c6c50742870aa042618ad', banner: true },
  'docs/commands.json': { sha256: 'e24e86f425f307b37e36984bd06575475d10537f1debac19f3972bc416b5153b', banner: false },
  'docs/demo.gif': { sha256: 'b0674ec14f82f5f54d4da9e0565377d5fb4f85202b4d597864a8a7a18066ce27', banner: false },
};

const BEGIN = '<!-- FROZEN-2.8.0:BEGIN -->\n';
const END = '<!-- FROZEN-2.8.0:END -->\n\n';

const sha = (bytes: Uint8Array): string => new Bun.CryptoHasher('sha256').update(bytes).digest('hex');

// The problems of one file: its bytes, or the bytes under its banner, must have the 2.8.0 hash.
const problemsOf = (path: string, bytes: Uint8Array, rule: Rule): ReadonlyArray<string> => {
  if (!rule.banner) return sha(bytes) === rule.sha256 ? [] : [`${path}: the bytes are not the 2.8.0 bytes`];
  const buffer = Buffer.from(bytes);
  const end = buffer.indexOf(END);
  if (buffer.indexOf(BEGIN) !== 0 || end === -1) return [`${path}: the file does not start with a FROZEN-2.8.0 banner and a blank line`];
  return sha(buffer.subarray(end + Buffer.byteLength(END))) === rule.sha256 ? [] : [`${path}: the text under the banner is not the 2.8.0 text`];
};

// The problems of one file as the check reads it: `bytes` is undefined when the file does not exist.
const fileProblems = (path: string, bytes: Uint8Array | undefined, rule: Rule): ReadonlyArray<string> =>
  bytes === undefined ? [`${path}: the file does not exist`] : problemsOf(path, bytes, rule);

const selftest = (): number => {
  const text = (value: string): Uint8Array => new TextEncoder().encode(value);
  const page = (banner: string, body: string): Uint8Array => text(`${BEGIN}${banner}\n${END}${body}`);
  const markdown = { sha256: sha(text('# Usage\n')), banner: true };
  // 0xff and 0xfe are not valid UTF-8: a text decode gives U+FFFD for each, so 0xfd looks the same.
  const gif = { sha256: sha(Uint8Array.of(0x47, 0x49, 0x46, 0xff, 0xfe)), banner: false };
  const cases = [
    { name: 'a right page passes', passed: problemsOf('t.md', page('> 2.8.0', '# Usage\n'), markdown).length === 0 },
    { name: 'a changed banner text passes', passed: problemsOf('t.md', page('> another banner', '# Usage\n'), markdown).length === 0 },
    { name: 'a changed byte under the banner is refused', passed: problemsOf('t.md', page('> 2.8.0', '# Usagf\n'), markdown).length > 0 },
    { name: 'a missing banner is refused', passed: problemsOf('t.md', text('# Usage\n'), markdown).length > 0 },
    { name: 'a banner with no BEGIN marker is refused', passed: problemsOf('t.md', text(`> 2.8.0\n${END}# Usage\n`), markdown).length > 0 },
    { name: 'the right gif bytes pass', passed: problemsOf('t.gif', Uint8Array.of(0x47, 0x49, 0x46, 0xff, 0xfe), gif).length === 0 },
    { name: 'a changed gif byte that a UTF-8 decode hides is refused', passed: problemsOf('t.gif', Uint8Array.of(0x47, 0x49, 0x46, 0xff, 0xfd), gif).length > 0 },
    { name: 'a missing file is refused', passed: fileProblems('t.md', undefined, markdown).join() === 't.md: the file does not exist' },
  ];
  const failed = cases.filter((c) => !c.passed);
  for (const c of failed) process.stderr.write(`check-frozen-docs: SELFTEST FAILED: ${c.name}\n`);
  if (failed.length === 0) process.stdout.write(`check-frozen-docs: selftest passed (${cases.map((c) => c.name).join('; ')}).\n`);
  return failed.length === 0 ? 0 : 1;
};

const check = async (): Promise<number> => {
  const problems: string[] = [];
  for (const [path, rule] of Object.entries(FROZEN)) {
    const file = Bun.file(path);
    problems.push(...fileProblems(path, (await file.exists()) ? new Uint8Array(await file.arrayBuffer()) : undefined, rule));
  }
  for (const problem of problems) process.stderr.write(`check-frozen-docs: ${problem}\n`);
  if (problems.length > 0) {
    process.stderr.write('check-frozen-docs: the npm page of ask-marcel-office-cli 2.8.0 links these files. Restore them from the v2.8.0 tag; only a banner text can change.\n');
    return 1;
  }
  process.stdout.write(`check-frozen-docs: the ${Object.keys(FROZEN).length} 2.8.0 docs keep the v2.8.0 bytes (${Object.keys(FROZEN).join(', ')}).\n`);
  return 0;
};

process.exit(process.argv.includes('--selftest') ? selftest() : await check());
