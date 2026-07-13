// OPT-IN, MUTATING live smoke of the 4 mail write commands. Closes the
// P1-DRAFT-01 gap: unit fakes cannot see that Graph's createReply/createForward
// body is the comment PLUS the quoted original, so only a live draft proves it.
//
// It creates 4 drafts in the signed-in mailbox, reads each body back, asserts it,
// and prints the draft IDs for YOU to delete (the CLI has no delete command, by
// design). It NEVER sends. Run against a warm login:
//
//   bun scripts/qa-write-smoke.ts --yes-create-drafts
//
// Overrides: SMOKE_TO=<addr> (default: signed-in user), SMOKE_BIN=<path to cli>.
// Output is status-only: lengths + PASS/FAIL + draft IDs, never the body text
// (which would carry real mailbox content).

const BIN = process.env.SMOKE_BIN ?? 'ask-marcel-office';
const TO = process.env.SMOKE_TO ?? '';

type Envelope = { ok: boolean; data?: unknown; error?: string };

const run = (args: ReadonlyArray<string>): Envelope => {
  const p = Bun.spawnSync([BIN, ...args, '--output', 'json'], { stdout: 'pipe', stderr: 'pipe' });
  const out = p.stdout.toString();
  const parsed = ((): unknown => {
    try {
      return JSON.parse(out);
    } catch {
      return undefined;
    }
  })();
  if (parsed !== null && typeof parsed === 'object') {
    const rec = parsed as Record<string, unknown>;
    return { ok: rec.ok === true, data: rec.data, error: typeof rec.error === 'string' ? rec.error : undefined };
  }
  return { ok: false, error: p.stderr.toString().slice(0, 200) || 'no JSON envelope' };
};

const idOf = (data: unknown): string | undefined => {
  if (data !== null && typeof data === 'object' && typeof (data as Record<string, unknown>).id === 'string') return (data as Record<string, unknown>).id as string;
  return undefined;
};

const firstMessageId = (data: unknown): string | undefined => {
  const rec = data as Record<string, unknown> | undefined;
  const value = rec?.value;
  const list = Array.isArray(value) ? value : Array.isArray(data) ? (data as ReadonlyArray<unknown>) : [];
  const first = list[0];
  return idOf(first);
};

const bodyContentOf = (messageId: string): string => {
  const r = run(['get-mail-message', '--message-id', messageId, '--select', 'body,subject']);
  const body = (r.data as Record<string, unknown> | undefined)?.body;
  const content = (body as Record<string, unknown> | undefined)?.content;
  return r.ok && typeof content === 'string' ? content : '';
};

const stamp = String(Bun.nanoseconds()); // unique per run without Date.now()
const results: Array<{ name: string; pass: boolean; detail: string }> = [];
const createdDraftIds: string[] = [];
const note = (name: string, pass: boolean, detail: string): void => {
  results.push({ name, pass, detail });
};

const main = async (): Promise<void> => {
  if (!process.argv.includes('--yes-create-drafts')) {
    process.stdout.write('Refusing to run: this creates real drafts in your mailbox.\nRe-run with --yes-create-drafts once you are ready (you delete the drafts after).\n');
    process.exit(2);
  }

  // Resolve the signed-in address (for reply/forward self-addressing) unless overridden.
  const me = TO || ((): string => {
    const r = run(['get-current-user', '--select', 'mail,userPrincipalName']);
    const rec = r.data as Record<string, unknown> | undefined;
    return (typeof rec?.mail === 'string' && rec.mail) || (typeof rec?.userPrincipalName === 'string' && rec.userPrincipalName) || '';
  })();
  if (!me) {
    process.stdout.write('Could not resolve a recipient address. Set SMOKE_TO=<you@tenant> and retry.\n');
    process.exit(1);
  }

  // Harvest a real source message to reply to / forward.
  const src = run(['list-mail-messages', '--top', '1', '--select', 'id,subject']);
  const srcId = firstMessageId(src.data);
  if (!src.ok || !srcId) {
    process.stdout.write(`Could not harvest a source message (list-mail-messages ok=${src.ok}). Need at least one message in the mailbox.\n`);
    process.exit(1);
  }

  // 1) create-reply-draft: body MUST be comment + quoted original (P1-DRAFT-01).
  const replyComment = `SMOKE reply ${stamp}`;
  const reply = run(['create-reply-draft', '--reply-to-message-id', srcId, '--body-content', replyComment]);
  const replyId = idOf(reply.data);
  if (reply.ok && replyId) {
    createdDraftIds.push(replyId);
    const body = bodyContentOf(replyId);
    const carriesQuote = body.length > replyComment.length + 100 && /From:|Sent:|wrote:|-{5,} ?Original/i.test(body);
    note('create-reply-draft', reply.ok && body.includes(replyComment.split(' ')[0] ?? 'SMOKE') && carriesQuote, `bodyLen=${body.length} vs commentLen=${replyComment.length}; carriesQuote=${carriesQuote}`);
  } else {
    note('create-reply-draft', false, `create ok=${reply.ok}, no draft id (err=${reply.error ?? '-'})`);
  }

  // 2) create-forward-draft: body MUST be comment + quoted original, recipients set.
  const fwdComment = `SMOKE forward ${stamp}`;
  const fwd = run(['create-forward-draft', '--forward-message-id', srcId, '--to-recipients', me, '--body-content', fwdComment]);
  const fwdId = idOf(fwd.data);
  if (fwd.ok && fwdId) {
    createdDraftIds.push(fwdId);
    const body = bodyContentOf(fwdId);
    const carriesQuote = body.length > fwdComment.length + 100 && /From:|Subject:|Sent:|-{5,} ?Forward/i.test(body);
    note('create-forward-draft', fwd.ok && carriesQuote, `bodyLen=${body.length} vs commentLen=${fwdComment.length}; carriesQuote=${carriesQuote}`);
  } else {
    note('create-forward-draft', false, `create ok=${fwd.ok}, no draft id (err=${fwd.error ?? '-'})`);
  }

  // 3) create-mail-draft: body is authored directly (no quote to preserve).
  const createBody = `SMOKE create body ${stamp}`;
  const created = run(['create-mail-draft', '--subject', `SMOKE create ${stamp}`, '--body-content', createBody, '--to-recipients', me]);
  const createdId = idOf(created.data);
  if (created.ok && createdId) {
    createdDraftIds.push(createdId);
    const body = bodyContentOf(createdId);
    note('create-mail-draft', created.ok && body.includes(createBody), `bodyLen=${body.length}; containsAuthoredBody=${body.includes(createBody)}`);

    // 4) update-mail-draft: replaces the body of the draft just created.
    const updatedBody = `SMOKE updated body ${stamp}`;
    const updated = run(['update-mail-draft', '--message-id', createdId, '--body-content', updatedBody]);
    const afterBody = bodyContentOf(createdId);
    note('update-mail-draft', updated.ok && afterBody.includes(updatedBody), `updated ok=${updated.ok}; containsNewBody=${afterBody.includes(updatedBody)}`);
  } else {
    note('create-mail-draft', false, `create ok=${created.ok}, no draft id (err=${created.error ?? '-'})`);
    note('update-mail-draft', false, 'skipped (no draft from create-mail-draft)');
  }

  // Report.
  process.stdout.write(`\n=== write-command live smoke (recipient: ${me}) ===\n`);
  for (const r of results) process.stdout.write(`${r.pass ? 'PASS' : 'FAIL'}  ${r.name.padEnd(22)} ${r.detail}\n`);
  const allPass = results.every((r) => r.pass);
  process.stdout.write(`\n${allPass ? 'ALL WRITE COMMANDS OK' : 'SOME CHECKS FAILED — inspect above'}\n`);
  process.stdout.write(`\nDELETE these ${createdDraftIds.length} drafts in Outlook (Drafts folder) — the CLI cannot delete:\n`);
  for (const id of createdDraftIds) process.stdout.write(`  ${id}\n`);
  process.stdout.write('\nSeparately, verify login --force interactively: run `ask-marcel-office login --force`, confirm the browser opens, then `ask-marcel-office login` should show elevated/ic3 as available.\n');
  process.exit(allPass ? 0 : 1);
};

await main();
