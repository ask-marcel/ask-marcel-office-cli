#!/usr/bin/env bun
/*
 * One-shot probe — for the Teams-substrate capture plan (Phase A).
 *
 * Launches Playwright (visible) against the persistent profile dir used by
 * the production CLI, navigates to teams.microsoft.com (where the Teams
 * web client lives), and harvests EVERY outgoing `Authorization: Bearer …`
 * header it sees while the user is signed in and clicking around.
 *
 * Each bearer's JWT is decoded; we record distinct (appid, aud) pairs
 * with their `scp` claim. The report highlights matches for the Teams
 * client identity (`appid = 1fec8e78-bce4-4aaf-ab1b-5451cc387264`) with
 * substrate audiences (`https://chatsvcagg.teams.microsoft.com` or
 * `https://api.spaces.skype.com`) — the target for Phase B.
 *
 * NOT shipped in the npm package — lives in scripts/ alongside other
 * one-shot tools. No tests; can't be unit-tested without launching a
 * real browser, same exclusion as playwright-loader.ts.
 *
 * Run:
 *   ask-marcel login                          # ensure profile cookies are fresh
 *   bun run scripts/probe-elevated-scopes.ts  # opens visible browser
 *   cat reports/probe-elevated-scopes.md      # inspect the result table
 *
 * The browser opens visibly. Sign in interactively if redirected. Click
 * into the Chat pane and open an actual conversation so the page makes
 * its chatsvcagg / substrate calls. Press ENTER in the terminal when
 * ready to advance.
 */

import { homedir } from 'node:os';
import { join } from 'node:path';
import { chromium, type Request } from 'playwright';
import { decodeJwtPayload } from '../src/domain/jwt-utils.ts';

// Reduced from the prior 6-URL set. teams.microsoft.com/v2/ is the new
// Teams web client; the Teams client appid emits its chatsvcagg / Skype
// substrate bearers during normal use here.
const CANDIDATES: ReadonlyArray<string> = ['https://teams.microsoft.com/v2/'];

const MIN_SETTLE_MS = 5_000;
const NAVIGATION_TIMEOUT_MS = 60_000;

const TEAMS_CLIENT_APP_ID = '1fec8e78-bce4-4aaf-ab1b-5451cc387264';
const SUBSTRATE_AUDS = ['https://chatsvcagg.teams.microsoft.com', 'https://api.spaces.skype.com'] as const;

type TokenObservation = {
  readonly url: string;
  readonly appid: string;
  readonly aud: string;
  readonly scp: string;
};

const profileDir = join(homedir(), '.ask-marcel', 'browser-profile');
const reportPath = join(process.cwd(), 'reports', 'probe-elevated-scopes.md');

const claimAsString = (v: unknown): string => {
  if (typeof v === 'string') return v;
  if (Array.isArray(v)) return v.filter((x): x is string => typeof x === 'string').join(' ');
  return '';
};

const observationKey = (o: TokenObservation): string => `${o.appid}|${o.aud}`;

const isTeamsSubstrate = (o: TokenObservation): boolean => o.appid === TEAMS_CLIENT_APP_ID && SUBSTRATE_AUDS.some((a) => o.aud === a || o.aud.includes(a));

const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

const waitForEnter = (message: string): Promise<void> =>
  new Promise<void>((resolve) => {
    process.stdout.write(message);
    const onData = (): void => {
      process.stdin.removeListener('data', onData);
      process.stdin.pause();
      resolve();
    };
    process.stdin.resume();
    process.stdin.on('data', onData);
  });

const formatRow = (o: TokenObservation): string => {
  const winner = isTeamsSubstrate(o) ? ' **🎯**' : '';
  return `| \`${o.appid}\`${winner} | \`${o.aud}\` | ${o.url} | ${o.scp || '_(no scp)_'} |`;
};

const writeReport = async (observations: ReadonlyArray<TokenObservation>): Promise<string> => {
  const winners = observations.filter(isTeamsSubstrate);
  const header = [
    `# Teams-substrate probe report — ${new Date().toISOString()}`,
    '',
    `Profile dir: \`${profileDir}\``,
    `Min settle per URL: ${MIN_SETTLE_MS / 1000}s, then user-driven (ENTER to advance)`,
    `Target: appid = \`${TEAMS_CLIENT_APP_ID}\` (Microsoft Teams client) AND aud ∈ { ${SUBSTRATE_AUDS.map((a) => `\`${a}\``).join(', ')} }`,
    '',
    '## Observed bearers (🎯 = Teams-substrate target)',
    '',
    '| appid | aud | seen-at-host | scp |',
    '|---|---|---|---|',
  ];
  const rows = observations.length === 0 ? ['| _(no bearers captured)_ | — | — | — |'] : observations.map(formatRow);
  const recommendation =
    winners.length === 0
      ? '\n## Recommendation\n\n**No Teams-substrate bearer was captured.** Possibilities:\n\n1. The page never emitted a `chatsvcagg` / `skype` audience bearer during the probe window. Re-run, click into the Chat pane, open an actual conversation, then ENTER.\n2. Cookies are stale — re-run `ask-marcel login` and try again.\n3. The tenant routes Teams traffic differently. Inspect the full `appid` / `aud` rows above for clues.'
      : `\n## Recommendation\n\n**${winners.length} Teams-substrate bearer(s) captured.** Phase B is unblocked.\n\nLead candidate:\n\n- appid: \`${winners[0].appid}\`\n- aud: \`${winners[0].aud}\`\n- seen at: ${winners[0].url}\n\nUse these in Phase B as \`TEAMS_CLIENT_APP_ID\` + \`TEAMS_SUBSTRATE_AUDS\`.`;
  const body = [...header, ...rows, recommendation].join('\n');
  await Bun.write(reportPath, body);
  return body;
};

const probeUrl = async (
  contextLike: { newPage: () => Promise<{ goto: (url: string, opts: { waitUntil: 'domcontentloaded'; timeout: number }) => Promise<unknown>; close: () => Promise<void> }> },
  url: string,
  index: number,
  total: number,
  observed: Map<string, TokenObservation>
): Promise<number> => {
  process.stdout.write(`\n>>> [${index + 1}/${total}] Probing ${url}\n`);
  const before = observed.size;
  const page = await contextLike.newPage();
  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: NAVIGATION_TIMEOUT_MS });
  } catch (e: unknown) {
    process.stdout.write(`    navigation failed: ${e instanceof Error ? e.message : String(e)}\n`);
  }
  await sleep(MIN_SETTLE_MS);
  await waitForEnter(`    Sign in, open the Chat pane, click into a conversation — then press ENTER to capture and close… `);
  await page.close();
  const added = observed.size - before;
  process.stdout.write(`    captured ${added} new (appid, aud) pair(s) for this URL\n`);
  return added;
};

const main = async (): Promise<void> => {
  process.stdout.write(`probe-elevated-scopes: launching visible Chrome against profile ${profileDir}\n`);
  const context = await chromium.launchPersistentContext(profileDir, {
    headless: false,
    channel: 'msedge',
    args: ['--disable-blink-features=AutomationControlled'],
  });

  const observed = new Map<string, TokenObservation>();

  let totalRequests = 0;
  let totalBearers = 0;
  const hostCounts = new Map<string, number>();
  const onRequest = (req: Request): void => {
    totalRequests += 1;
    try {
      const host = new URL(req.url()).host;
      hostCounts.set(host, (hostCounts.get(host) ?? 0) + 1);
    } catch {
      /* ignore unparseable URLs */
    }
    const auth = req.headers()['authorization'];
    if (typeof auth !== 'string' || !auth.toLowerCase().startsWith('bearer ')) return;
    totalBearers += 1;
    const token = auth.slice('bearer '.length).trim();
    if (token.split('.').length !== 3) return;
    const claims = decodeJwtPayload(token);
    const appid = claimAsString(claims['appid']);
    const aud = claimAsString(claims['aud']);
    const scp = claimAsString(claims['scp']);
    // No audience filter — record EVERY bearer we see. The report flags
    // Teams-substrate matches with 🎯. Useful to see what audiences a
    // Teams session uses even when none of them is the substrate target.
    if (!appid) return;
    let host = '';
    try {
      host = new URL(req.url()).host;
    } catch {
      /* ignore */
    }
    const obs: TokenObservation = { url: host, appid, aud, scp };
    observed.set(observationKey(obs), obs);
  };
  context.on('request', onRequest);

  process.stdout.write(`\nThe browser window is now open at teams.microsoft.com/v2/.\n`);
  process.stdout.write(`Sign in if redirected, then open the Chat pane and click into a real\n`);
  process.stdout.write(`conversation so the page makes its chatsvcagg / substrate calls.\n`);
  process.stdout.write(`Press ENTER in this terminal when ready to capture the report.\n`);

  try {
    for (let i = 0; i < CANDIDATES.length; i += 1) {
      await probeUrl(context, CANDIDATES[i] ?? '', i, CANDIDATES.length, observed);
    }
  } finally {
    await context.close();
  }

  const all = [...observed.values()].toSorted((a, b) => {
    // Teams-substrate matches sort first; then by appid; then aud.
    const aw = isTeamsSubstrate(a) ? 0 : 1;
    const bw = isTeamsSubstrate(b) ? 0 : 1;
    if (aw !== bw) return aw - bw;
    return (a.appid + a.aud).localeCompare(b.appid + b.aud);
  });
  process.stdout.write(`\nprobe-elevated-scopes diagnostics:\n`);
  process.stdout.write(`  total requests seen across all pages: ${totalRequests}\n`);
  process.stdout.write(`  total bearer-bearing requests:        ${totalBearers}\n`);
  process.stdout.write(`  distinct (appid, aud) pairs:          ${all.length}\n`);
  const winners = all.filter(isTeamsSubstrate);
  process.stdout.write(`  Teams-substrate matches:              ${winners.length}\n`);
  const topHosts = [...hostCounts.entries()].toSorted((a, b) => b[1] - a[1]).slice(0, 15);
  process.stdout.write(`  top 15 request hosts:\n`);
  for (const [host, count] of topHosts) process.stdout.write(`    ${host}: ${count}\n`);
  const body = await writeReport(all);
  process.stdout.write(`\n${body}\n\nReport written to ${reportPath}\n`);
};

await main();
