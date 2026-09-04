#!/usr/bin/env bun
/*
 * QA Phase E1 — live command sweep (docs/QA-PLAYBOOK.md §E1).
 *
 * Manifest-driven and STATUS-ONLY BY CONSTRUCTION: for every command it records
 * only the command name + `ok`/`errorCode` — it NEVER logs tenant content
 * (names, subjects, ids). Read-only tenant access (mutating draft commands are
 * skipped). Harvests an id pool live, runs one happy-path per command, and
 * reports harvest completeness so untestable-this-run commands are explicit,
 * never silently counted as healthy.
 *
 * Prereqs: clean `main`, `bun run build`, `npm i -g .`, a warm `ask-marcel-office
 * login`. Run: `bun scripts/qa-live-sweep.ts`. Ledger → /tmp/qa/live-ledger.json.
 *
 * Replaces the ad-hoc drivers that used to rot in .claude/qa-reports/ (they
 * invoked the pre-rename binary and depended on stale temp files).
 */
import { spawnSync } from 'node:child_process';

const BIN = 'ask-marcel-office';
const manifest = JSON.parse(await Bun.file('docs/commands.json').text());

let browserOpens = 0;
const run = (args: string[], timeoutMs = 90000): { ok: boolean; data?: any; code?: string } => {
  const p = spawnSync(BIN, [...args, '--output', 'json'], { timeout: timeoutMs, maxBuffer: 128 * 1024 * 1024, env: { ...process.env, ASKMARCEL_LOG_LEVEL: 'info' } });
  const errout = p.stderr?.toString() || '';
  if (/system_browser|"rung":"browser"|ask_marcel_port|fallback_to_playwright/.test(errout)) browserOpens++;
  if (p.signal === 'SIGTERM') return { ok: false, code: 'TIMEOUT' };
  try {
    const d = JSON.parse(p.stdout?.toString() || '');
    return d.ok ? { ok: true, data: d.data } : { ok: false, code: d.errorCode || 'err' };
  } catch {
    return { ok: false, code: 'UNPARSEABLE' };
  }
};
const first = (r: any): any => r?.data?.value?.[0];
const val = (r: any): any[] => r?.data?.value || [];

// ---------------- HARVEST ----------------
const pool: Record<string, string | undefined> = {};
const P = (k: string, v: any): void => { if (v && !pool[k]) pool[k] = String(v); };

const qc = run(['my-quick-context']).data || {};
P('user-id', run(['get-current-user']).data?.id || qc.user?.id);
P('drive-id', qc.primaryDriveId);
P('mail-folder-id', qc.inboxId);
P('calendar-id', qc.primaryCalendarId);
const rootItemId = pool['drive-id'] ? run(['get-drive-root-item', '--drive-id', pool['drive-id']]).data?.id : undefined;

// drive items: any file (+ a real .xlsx) from several vectors
const items: any[] = [];
for (const src of [
  rootItemId ? ['list-folder-files', '--drive-id', pool['drive-id']!, '--item-id', rootItemId, '--top', '50'] : null,
  ['list-recent-files', '--top', '50'],
  ['list-shared-with-me', '--top', '30'],
  ['search-my-documents', '--query', 'report', '--top', '25'],
].filter(Boolean) as string[][]) {
  for (const it of val(run(src))) {
    const di = it.remoteItem?.parentReference?.driveId || it.parentReference?.driveId || pool['drive-id'];
    const id = it.remoteItem?.id || it.id;
    if (id && di) items.push({ driveId: di, itemId: id, name: it.name || '', isFile: !!(it.file || it.remoteItem?.file) });
  }
}
const anyFile = items.find((x) => x.isFile) || items[0];
const xlsx = items.find((x) => /\.xlsx$/i.test(x.name)) || anyFile;
const pairs = { file: anyFile, xlsx, folder: rootItemId ? { driveId: pool['drive-id'], itemId: rootItemId } : undefined };

// mail: a message, and one WITH an attachment
const msgs = run(['list-mail-messages', '--top', '15']);
let attMsg: any, attId: any;
for (const mm of val(msgs)) { if (mm.hasAttachments) { const a = first(run(['list-mail-attachments', '--message-id', mm.id])); if (a) { attMsg = mm.id; attId = a.id; break; } } }
P('message-id', attMsg || first(msgs)?.id);
P('attachment-id', attId);
P('message-rule-id', first(run(['list-mail-rules']))?.id);
// calendar
P('event-id', first(run(['list-calendar-events', '--top', '5']))?.id);
P('calendar-id', first(run(['list-calendars']))?.id);
P('calendar-group-id', first(run(['list-calendar-groups']))?.id);
// sharepoint
const site = first(run(['search-sharepoint-sites-by-name', '--query', 'team', '--top', '5'])) || first(run(['search-all-accessible-sites', '--query', 'team']));
P('site-id', site?.id);
const siteUrl = site?.webUrl ? new URL(site.webUrl) : undefined;
P('hostname', siteUrl?.hostname);
if (pool['site-id']) {
  const list = first(run(['list-sharepoint-site-lists', '--site-id', pool['site-id']]));
  P('list-id', list?.id);
  if (pool['list-id']) {
    P('list-item-id', first(run(['list-sharepoint-site-list-items', '--site-id', pool['site-id'], '--list-id', pool['list-id'], '--top', '2']))?.id);
    P('column-id', first(run(['list-sharepoint-list-columns', '--site-id', pool['site-id'], '--list-id', pool['list-id']]))?.id);
  }
}
// groups / teams / chats
// A UNIFIED group the signed-in user BELONGS to. `list-groups` is a tenant-wide
// DIRECTORY read, so its first hit is typically a security or distribution group
// with no mailbox that the user is not a member of: the group-mailbox commands then
// answer MailboxNotEnabledForRESTAPI, which reads as a product failure and is not,
// and no thread id is harvested at all — which left the whole six-command group-post
// family unaudited on the run that shipped it (QA-DRIVER-02, 2026-09-04).
const unifiedGroups = val(run(['list-my-memberships', '--top', '60'])).filter(
  (g: any) => g['@odata.type'] === '#microsoft.graph.group' && (g.groupTypes ?? []).includes('Unified')
);
P('group-id', (unifiedGroups[0] ?? first(run(['list-groups', '--top', '5'])))?.id);
// thread -> post -> post-attachment, from the first unified group that has a post.
let postAttId: string | undefined;
for (const g of unifiedGroups) {
  for (const t of val(run(['list-group-threads', '--group-id', g.id, '--top', '10']))) {
    const posts = val(run(['list-group-thread-posts', '--group-id', g.id, '--thread-id', t.id]));
    if (posts.length === 0) continue;
    pool['group-id'] = String(g.id);
    P('thread-id', t.id);
    P('post-id', posts[0].id);
    postAttId ??= first(run(['list-group-post-attachments', '--group-id', g.id, '--thread-id', t.id, '--post-id', posts[0].id]))?.id;
    break;
  }
  if (pool['thread-id'] && postAttId) break;
}
P('team-id', first(run(['list-joined-teams']))?.id);
if (pool['team-id']) for (const ch of val(run(['list-team-channels', '--team-id', pool['team-id']]))) { if (run(['get-channel-files-folder', '--team-id', pool['team-id'], '--channel-id', ch.id]).ok) { P('channel-id', ch.id); break; } P('channel-id', ch.id); }
const chatId = first(run(['list-chats', '--top', '5']))?.id;
P('chat-id', chatId);
// a chat MESSAGE id (get-teams-chat-message needs one) — try a few chats for a non-empty one
let chatMsgId: any;
for (const c of val(run(['list-chats', '--top', '8']))) { const mid = first(run(['list-teams-chat-messages', '--chat-id', c.id]))?.id; if (mid) { pool['chat-id'] = c.id; chatMsgId = mid; break; } }
// tasks
P('todo-task-list-id', first(run(['list-todo-task-lists']))?.id);
if (pool['todo-task-list-id']) P('todo-task-id', first(run(['list-todo-tasks', '--todo-task-list-id', pool['todo-task-list-id'], '--top', '3']))?.id);
for (const plan of val(run(['list-planner-plans'])).slice(0, 8)) {
  P('planner-plan-id', plan.id);
  const t = first(run(['list-plan-tasks', '--planner-plan-id', plan.id]));
  const b = first(run(['list-plan-buckets', '--planner-plan-id', plan.id]));
  if (t) P('planner-task-id', t.id);
  if (b) P('planner-bucket-id', b.id);
  if (pool['planner-task-id'] && pool['planner-bucket-id']) break;
}
// onenote (personal)
P('notebook-id', first(run(['list-onenote-notebooks']))?.id);
if (pool['notebook-id']) { P('onenote-section-id', first(run(['list-onenote-notebook-sections', '--notebook-id', pool['notebook-id']]))?.id);
  if (pool['onenote-section-id']) P('onenote-page-id', first(run(['list-onenote-section-pages', '--onenote-section-id', pool['onenote-section-id'], '--top', '2']))?.id); }
// site onenote (separate ids — a site WITH a notebook)
const siteNb: any = { siteId: undefined, notebookId: undefined, sectionId: undefined, pageId: undefined };
for (const s of [site, ...val(run(['search-all-accessible-sites', '--query', 'team'])).slice(0, 6)].filter(Boolean)) {
  const nb = first(run(['list-sharepoint-site-onenote-notebooks', '--site-id', s.id]));
  if (nb) { siteNb.siteId = s.id; siteNb.notebookId = nb.id;
    const sec = first(run(['list-sharepoint-site-onenote-notebook-sections', '--site-id', s.id, '--notebook-id', nb.id]));
    if (sec) { siteNb.sectionId = sec.id; siteNb.pageId = first(run(['list-sharepoint-site-onenote-section-pages', '--site-id', s.id, '--onenote-section-id', sec.id]))?.id; }
    break; }
}
// excel worksheet/table/chart
if (xlsx) { P('worksheet-id', first(run(['list-excel-worksheets', '--drive-id', xlsx.driveId, '--item-id', xlsx.itemId]))?.name);
  P('table-id', first(run(['list-excel-tables', '--drive-id', xlsx.driveId, '--item-id', xlsx.itemId]))?.name);
  if (pool['worksheet-id']) P('chart-id', first(run(['list-excel-worksheet-charts', '--drive-id', xlsx.driveId, '--item-id', xlsx.itemId, '--worksheet-id', pool['worksheet-id']]))?.id); }
// a NON-current drive-item version (current is refused by design)
if (anyFile) { const vers = val(run(['list-drive-item-versions', '--drive-id', anyFile.driveId, '--item-id', anyFile.itemId])); P('version-id', (vers[1] || vers[0])?.id); }
// a group conversation (boundary-dependent)
let convGroup: any, convId: any;
for (const g of [...unifiedGroups, ...val(run(['list-groups', '--top', '15']))]) { const conv = first(run(['list-group-conversations', '--group-id', g.id])); if (conv) { convGroup = g.id; convId = conv.id; break; } }
// a real nextLink for next-page
const nlProbe = run(['list-mail-messages', '--top', '2']);
const nextLink = (nlProbe as any).data?.nextLink ?? undefined;

if (browserOpens > 0) { console.log(`!! browserOpens=${browserOpens} during harvest — ABORT (would storm)`); process.exit(1); }

// ---------------- ARG RESOLUTION (command-aware) ----------------
const DS = '2026-06-01T00:00:00Z', DE = '2026-07-05T23:59:59Z';
const LOCAL = new Set(['convert-local-file-to-markdown', 'extract-local-file-images']);
const EXCEL = /excel|used-range|worksheet|(^get-excel)/;
const pickPair = (cmd: string): any => (/zip/.test(cmd) ? undefined : EXCEL.test(cmd) ? pairs.xlsx : cmd === 'list-folder-files' ? pairs.folder : pairs.file);

const argFor = (cmd: string, opt: string): string[] | null => {
  const pr = pickPair(cmd);
  // command-aware special cases first
  // Both local commands take --path, but not the same kind of file: a .msg has
  // no embedded images, so feeding it to extract-local-file-images produced an
  // `unsupported_document` ERR on every run. That false failure was written off
  // as a harvest artifact twice (2026-07-20, 2026-08-31), which left the command
  // effectively unaudited — a real regression in it would have looked identical.
  if (opt === 'path')
    return LOCAL.has(cmd)
      ? ['--path', cmd === 'extract-local-file-images' ? 'src/test-helpers/assets/image-sample.docx' : 'src/test-helpers/assets/sample.msg']
      : siteUrl
        ? ['--path', siteUrl.pathname]
        : null;
  if (opt === 'url') return cmd === 'next-page' ? (nextLink ? ['--url', nextLink] : null)
    : cmd === 'resolve-drive-share-link' ? ['--url', site?.webUrl || 'https://example.sharepoint.com/x']
    : cmd === 'resolve-mail-link' ? ['--url', 'https://outlook.office.com/mail/inbox/id/AAQkAGnope']
    : cmd === 'resolve-teams-link' ? ['--url', 'https://teams.microsoft.com/l/message/19:x@thread.v2/1700000000000']
    : ['--url', 'https://outlook.office365.com/calendar/item/AAMknope'];
  if (opt === 'message-id' && cmd === 'get-teams-chat-message') return chatMsgId ? ['--message-id', chatMsgId] : null;
  if (opt === 'notebook-id' && /sharepoint-site/.test(cmd)) return siteNb.notebookId ? ['--notebook-id', siteNb.notebookId] : null;
  if (opt === 'onenote-section-id' && /sharepoint-site/.test(cmd)) return siteNb.sectionId ? ['--onenote-section-id', siteNb.sectionId] : null;
  if (opt === 'onenote-page-id' && /sharepoint-site/.test(cmd)) return siteNb.pageId ? ['--onenote-page-id', siteNb.pageId] : null;
  if (opt === 'site-id' && /sharepoint-site-onenote/.test(cmd)) return siteNb.siteId ? ['--site-id', siteNb.siteId] : (pool['site-id'] ? ['--site-id', pool['site-id']] : null);
  if (opt === 'attachment-id' && /group-post/.test(cmd)) return postAttId ? ['--attachment-id', postAttId] : null;
  if (opt === 'conversation-id') return convId ? ['--conversation-id', convId] : null;
  if (opt === 'group-id' && /conversation/.test(cmd)) return convGroup ? ['--group-id', convGroup] : null;
  if (opt === 'drive-id') return pr?.driveId ? ['--drive-id', pr.driveId] : null;
  if (opt === 'item-id') return pr?.itemId ? ['--item-id', pr.itemId] : null;
  // generic map (verified against the manifest's required-option names)
  const map: Record<string, string | undefined> = {
    ...pool,
    'start-date-time': DS, 'end-date-time': DE,
    query: 'report', 'title-substring': 'meeting', name: pool['user-id'] ? 'user' : 'user', 'folder-name': 'documents',
    address: 'A1:C5', range: 'A1:C5', schedules: qc.user?.userPrincipalName || 'me',
  };
  const v = map[opt];
  return v ? [`--${opt}`, v] : null;
};

// ---------------- SWEEP ----------------
const SKIP = new Set(manifest.commands.filter((c: any) => c.mutates).map((c: any) => c.name));
const ledger: Record<string, string> = {};
const nodataOpts = new Set<string>();
for (const c of manifest.commands) {
  if (SKIP.has(c.name)) { ledger[c.name] = 'SKIP-mutating'; continue; }
  let argv = [c.name];
  let unresolved: string | null = null;
  for (const o of (c.options || []).filter((x: any) => x.required)) { const a = argFor(c.name, o.name); if (!a) { unresolved = o.name; break; } argv.push(...a); }
  if ((c.options || []).some((o: any) => o.name === 'top') && !argv.includes('--top')) argv.push('--top', '3');
  if (unresolved) { ledger[c.name] = `NODATA-${unresolved}`; nodataOpts.add(unresolved); continue; }
  const r = run(argv);
  ledger[c.name] = r.ok ? 'ok' : `ERR:${r.code}`;
  if (browserOpens > 0) { console.log(`!! CIRCUIT BREAKER at ${c.name} (browserOpens=${browserOpens}) — ABORT`); break; }
}
try { await Bun.write('/tmp/qa/live-ledger.json', JSON.stringify(ledger, null, 1)); } catch { /* ok */ }

// ---------------- REPORT (status only) ----------------
const counts: Record<string, number> = {};
for (const v of Object.values(ledger)) { const k = v.startsWith('ok') ? 'ok' : v.startsWith('ERR') ? 'ERR' : v.startsWith('NODATA') ? 'NODATA' : 'SKIP'; counts[k] = (counts[k] || 0) + 1; }
console.log('\n=== HARVEST COMPLETENESS ===');
console.log('  resolved ids:', Object.keys(pool).filter((k) => pool[k]).sort().join(', '));
console.log('  UNRESOLVED id types (→ NODATA):', [...nodataOpts].sort().join(', ') || '(none — full coverage)');
console.log('\n=== SWEEP', JSON.stringify(counts), `browserOpens=${browserOpens} (MUST be 0) total=${Object.keys(ledger).length} ===`);
console.log('\n=== ERR (command -> errorCode; no tenant content) ==='); for (const [k, v] of Object.entries(ledger)) if (v.startsWith('ERR')) console.log(`  ${k} -> ${v}`);
console.log('\n=== NODATA (harvest could not supply a required id this run) ==='); for (const [k, v] of Object.entries(ledger)) if (v.startsWith('NODATA')) console.log(`  ${k} -> ${v}`);
if (browserOpens > 0) process.exit(1);
