// OPT-IN, MUTATING live smoke of the drafts-dedup finding (eval item C). Proves
// what unit fakes cannot: that a conversationId $filter on the Drafts folder
// misses reply drafts a thread has split across several conversationIds, while
// find-mail-drafts (subject + recipient match, client-side) finds them all.
//
// It creates 3 reply drafts on one real thread, then compares the two lookup
// paths, and prints the draft IDs for YOU to delete (the CLI has no delete
// command, by design). It NEVER sends. Run against a warm login:
//
//   bun scripts/qa-dedup-smoke.ts --yes-create-drafts
//
// Overrides: SMOKE_BIN=<path to cli>. Output is status-only (ids + counts),
// never body text.

const BIN = process.env.SMOKE_BIN ?? 'ask-marcel-office';

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

const stringField = (data: unknown, field: string): string | undefined => {
  if (data === null || typeof data !== 'object') return undefined;
  const value = (data as Record<string, unknown>)[field];
  return typeof value === 'string' ? value : undefined;
};

const arrayUnder = (data: unknown, field: 'value'): ReadonlyArray<unknown> => {
  const list = (data as Record<string, unknown> | undefined)?.[field];
  return Array.isArray(list) ? list : [];
};

const idsUnder = (data: unknown, field: 'value' | 'matches'): ReadonlyArray<string> => {
  const list = (data as Record<string, unknown> | undefined)?.[field];
  if (!Array.isArray(list)) return [];
  return list.flatMap((item) => {
    const id = stringField(item, 'id');
    return id === undefined ? [] : [id];
  });
};

const createdDraftIds: string[] = [];

const main = async (): Promise<void> => {
  if (!process.argv.includes('--yes-create-drafts')) {
    process.stdout.write('Refusing to run: this creates real drafts in your mailbox.\nRe-run with --yes-create-drafts once you are ready (you delete the drafts after).\n');
    process.exit(2);
  }

  // Harvest replyable Inbox candidates. Drafts, delivery reports, and meeting
  // items reject createReply (ErrorInvalidReferenceItem), so try candidates
  // until one accepts a reply draft; that message's thread becomes the target.
  const listed = run(['list-mail-folder-messages', '--mail-folder-id', 'inbox', '--top', '10', '--select', 'id,subject,conversationId']);
  const candidates = arrayUnder(listed.data, 'value');
  if (!listed.ok || candidates.length === 0) {
    process.stdout.write(`Could not list Inbox messages to harvest a thread (ok=${listed.ok}). Need at least one replyable message in the Inbox.\n`);
    process.exit(1);
  }

  const stamp = String(Bun.nanoseconds());
  const draftConversationIds: string[] = [];
  let replyTargetId = '';
  let subject = '';
  let c0 = '';
  for (const candidate of candidates) {
    const id = stringField(candidate, 'id');
    const candidateSubject = stringField(candidate, 'subject') ?? '';
    const candidateC0 = stringField(candidate, 'conversationId');
    if (id === undefined || candidateC0 === undefined || candidateSubject.length === 0) continue;
    const reply = run(['create-reply-draft', '--reply-to-message-id', id, '--body-content', `SMOKE dedup ${stamp} #0`]);
    const draftId = stringField(reply.data, 'id');
    if (reply.ok && draftId !== undefined) {
      createdDraftIds.push(draftId);
      const conv = stringField(reply.data, 'conversationId');
      if (conv !== undefined) draftConversationIds.push(conv);
      replyTargetId = id;
      subject = candidateSubject;
      c0 = candidateC0;
      break;
    }
  }
  if (createdDraftIds.length === 0) {
    process.stdout.write('No Inbox message in the top 10 accepted a reply draft (all drafts/reports/meeting items?). Try a mailbox with ordinary received mail.\n');
    process.exit(1);
  }

  // Create 2 more reply drafts on the SAME inbound message; record each returned
  // conversationId (they may split off the inbound C0 — that is the finding).
  for (let i = 1; i < 3; i += 1) {
    const reply = run(['create-reply-draft', '--reply-to-message-id', replyTargetId, '--body-content', `SMOKE dedup ${stamp} #${i}`]);
    const draftId = stringField(reply.data, 'id');
    if (!reply.ok || draftId === undefined) {
      process.stdout.write(`follow-up create-reply-draft #${i} failed (ok=${reply.ok}, err=${reply.error ?? '-'}). Continuing with ${createdDraftIds.length} draft(s).\n`);
      break;
    }
    createdDraftIds.push(draftId);
    const conv = stringField(reply.data, 'conversationId');
    if (conv !== undefined) draftConversationIds.push(conv);
  }

  const distinctConversations = [...new Set(draftConversationIds)];
  const splitObserved = distinctConversations.some((conv) => conv !== c0) || distinctConversations.length > 1;

  // OLD path: filter Drafts by the inbound conversationId. Count how many of OUR
  // drafts it returns. Fewer than we created is the finding (split and/or lag).
  const filtered = run(['list-mail-folder-messages', '--mail-folder-id', 'drafts', '--filter', `conversationId eq '${c0}'`, '--top', '100']);
  const filteredIds = new Set(idsUnder(filtered.data, 'value'));
  const oldPathHits = createdDraftIds.filter((id) => filteredIds.has(id)).length;

  // NEW path: find-mail-drafts on the subject. Must return ALL our drafts.
  const found = run(['find-mail-drafts', '--subject', subject]);
  const foundIds = new Set(idsUnder(found.data, 'matches'));
  const newPathHits = createdDraftIds.filter((id) => foundIds.has(id)).length;

  process.stdout.write('\n=== drafts-dedup live smoke ===\n');
  process.stdout.write(`created drafts:            ${createdDraftIds.length}\n`);
  process.stdout.write(`inbound conversationId C0: ${c0.slice(0, 24)}...\n`);
  process.stdout.write(`distinct draft convIds:    ${distinctConversations.length} (conversationId split observed: ${splitObserved})\n`);
  process.stdout.write(`OLD path ($filter eq C0):  ${oldPathHits}/${createdDraftIds.length} of our drafts (ok=${filtered.ok})\n`);
  process.stdout.write(`NEW path (find-mail-drafts): ${newPathHits}/${createdDraftIds.length} of our drafts (ok=${found.ok})\n`);

  const newPathWins = createdDraftIds.length > 0 && newPathHits === createdDraftIds.length && newPathHits >= oldPathHits;
  process.stdout.write(`\n${newPathWins ? 'NEW PATH FOUND ALL DRAFTS' : 'CHECK FAILED — find-mail-drafts did not cover every created draft'}\n`);
  if (oldPathHits < createdDraftIds.length) process.stdout.write('(as expected, the conversationId $filter under-returned — that is the finding.)\n');

  process.stdout.write(`\nDELETE these ${createdDraftIds.length} drafts in Outlook (Drafts folder) — the CLI cannot delete:\n`);
  for (const id of createdDraftIds) process.stdout.write(`  ${id}\n`);
  process.exit(newPathWins ? 0 : 1);
};

await main();
