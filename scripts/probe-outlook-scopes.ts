#!/usr/bin/env bun
/*
 * Probe: does any first-party app Outlook Web loads emit a Graph-audience
 * bearer carrying `Mail.Read.Shared`?
 *
 * Why this exists. The five `list-shared-mailbox-*` commands need the delegated
 * `Mail.Read.Shared` scope and answer `ErrorAccessDenied` without it. The CLI
 * owns no app registration, so a scope cannot be requested: it captures bearers
 * that Microsoft's own web apps emit, and a scope is reachable exactly when some
 * harvestable first-party app already carries it. That ceiling has moved once
 * before (OfficeHome carries `Contacts.Read` where the Teams client does not,
 * see the comment above ELEVATED_APP_IDS in src/infra/browser-auth.ts), so the
 * question is empirical, not architectural.
 *
 * What it does. Opens the same persistent browser profile `login` uses, loads
 * Outlook Web, and decodes every outgoing `Authorization: Bearer` header the way
 * browser-auth.ts does. It reports, per appid, the audiences seen and the scopes
 * granted.
 *
 * What it never does. Print, log or persist a token. Only decoded `appid`,
 * `aud` and `scp` claims are shown. It navigates read-only and sends nothing.
 *
 * Run:  bun run scripts/probe-outlook-scopes.ts
 *       bun run scripts/probe-outlook-scopes.ts --headless
 *
 * A first run may need an interactive sign-in; leave the window open until the
 * mailbox has rendered.
 *
 * RESULT, 2026-09-06, one tenant. Outlook Web emits a Graph-audience bearer
 * under appid 9199bf20-a13f-4107-85dc-02114787ef48 carrying `Mail.ReadWrite.Shared`
 * and `Mail.Send.Shared` (not `Mail.Read.Shared`), alongside roughly a hundred
 * others including `Mail.Send`, `Files.ReadWrite.All` and `Group.ReadWrite.All`.
 * So the scope IS reachable. See `probe-outlook-shared-read.ts` for why that
 * does not settle the question, and why the answer is still no.
 *
 * Capture is opportunistic, and this is the hard part. The appid emits bearers
 * on nearly every load but almost always for a NON-Graph audience; across
 * 2026-09-06 the Graph-audience mint happened only on the day's first cold
 * start. Waiting longer does not help, nor does reloading, nor does clearing the
 * origin's web storage to force MSAL to re-acquire. Run this first thing,
 * against a profile that has not opened Outlook yet that day. A run that reports
 * the appid with "(no Graph audience)" is this, not a broken probe.
 */

import { join } from 'node:path';
import { decodeJwtPayload } from '@ask-marcel/office-core/domain/jwt-utils.ts';

const GRAPH_AUD = 'https://graph.microsoft.com';

/** The scope the shared-mailbox commands are blocked on, plus its neighbours. */
const SCOPES_OF_INTEREST = ['Mail.Read.Shared', 'Mail.ReadWrite.Shared', 'Mail.Send.Shared', 'Mail.Read', 'Mail.ReadWrite', 'Mail.ReadBasic'];

/**
 * Outlook Web has moved hosts more than once and a tenant may land on either.
 * Both are visited so a redirect to the other does not read as "no tokens".
 */
const OUTLOOK_URLS = ['https://outlook.cloud.microsoft/mail/', 'https://outlook.office.com/mail/'];

const profileDir = process.env['ASKMARCEL_BROWSER_PROFILE'] ?? join(process.env['HOME'] ?? '', '.ask-marcel', 'browser-profile');
const headless = process.argv.includes('--headless');
const settleMs = Number(process.env['PROBE_SETTLE_MS'] ?? 45_000);

type Seen = { readonly audiences: Set<string>; readonly scopes: Set<string>; readonly urls: Set<string> };
const byAppId = new Map<string, Seen>();

const record = (appid: string, aud: string, scp: string, url: string): void => {
  const entry = byAppId.get(appid) ?? { audiences: new Set<string>(), scopes: new Set<string>(), urls: new Set<string>() };
  entry.audiences.add(aud);
  for (const scope of scp.split(' ').filter((s) => s !== '')) entry.scopes.add(scope);
  // Host only: a full URL can carry mailbox ids and message ids in its path.
  entry.urls.add(new URL(url).host);
  byAppId.set(appid, entry);
};

const report = (): void => {
  if (byAppId.size === 0) {
    console.log('\nNo bearer tokens were seen. Either the sign-in did not complete, or the page never called an API in the settle window.');
    console.log('Retry with a longer window: PROBE_SETTLE_MS=120000 bun run scripts/probe-outlook-scopes.ts');
    return;
  }

  console.log(`\n${'='.repeat(78)}\nAppids seen, with the audiences and scopes their bearers carried\n${'='.repeat(78)}`);
  for (const [appid, seen] of [...byAppId].toSorted((a, b) => a[0].localeCompare(b[0]))) {
    const graph = seen.audiences.has(GRAPH_AUD);
    console.log(`\nappid ${appid}   ${graph ? '<- GRAPH AUDIENCE (usable by this CLI)' : '(no Graph audience)'}`);
    console.log(`  hosts     ${[...seen.urls].toSorted((a, b) => a.localeCompare(b)).join(', ')}`);
    console.log(`  audiences ${[...seen.audiences].toSorted((a, b) => a.localeCompare(b)).join(', ')}`);
    console.log(`  scopes    ${[...seen.scopes].toSorted((a, b) => a.localeCompare(b)).join(' ') || '(none in the token)'}`);
  }

  console.log(`\n${'='.repeat(78)}\nVERDICT\n${'='.repeat(78)}`);
  const graphAppIds = [...byAppId].filter(([, seen]) => seen.audiences.has(GRAPH_AUD));
  if (graphAppIds.length === 0) {
    console.log('No Graph-audience bearer was captured at all, so nothing here is harvestable by this CLI.');
    return;
  }
  for (const scope of SCOPES_OF_INTEREST) {
    const carriers = graphAppIds.filter(([, seen]) => seen.scopes.has(scope)).map(([appid]) => appid);
    console.log(`  ${scope.padEnd(22)} ${carriers.length > 0 ? `CARRIED by ${carriers.join(', ')}` : 'not carried by any Graph-audience token seen'}`);
  }
  const sharedCarriers = graphAppIds.filter(([, seen]) => seen.scopes.has('Mail.Read.Shared') || seen.scopes.has('Mail.ReadWrite.Shared'));
  console.log(
    sharedCarriers.length > 0
      ? `\nA *.Shared mail scope IS carried, by appid ${sharedCarriers.map(([a]) => a).join(' or ')}. That is NOT the same as access:
on 2026-09-06 this token was refused 403 ErrorAccessDenied on /users/{own-upn}/messages, the very
endpoint the shared-mailbox commands call. Run probe-outlook-shared-read.ts before concluding
anything, and do not wire a third identity on the strength of a scope list alone.`
      : '\nNo harvestable Outlook Web token carries a *.Shared mail scope on this tenant.'
  );
};

const run = async (): Promise<number> => {
  const { chromium } = await import('playwright');
  console.log(`probe: profile ${profileDir}`);
  console.log(`probe: headless=${headless}, settle=${settleMs}ms\n`);

  let context;
  for (const channel of ['msedge', 'chrome'] as const) {
    try {
      context = await chromium.launchPersistentContext(profileDir, { headless, channel, args: ['--disable-blink-features=AutomationControlled'] });
      console.log(`probe: launched ${channel}`);
      break;
    } catch {
      // try the next channel
    }
  }
  if (context === undefined) {
    console.error('probe: could not launch msedge or chrome.');
    return 1;
  }

  const page = await context.newPage();
  page.on('request', (req) => {
    const auth = req.headers()['authorization'];
    if (typeof auth !== 'string' || !auth.startsWith('Bearer ')) return;
    const claims = decodeJwtPayload(auth.slice('Bearer '.length));
    const appid = typeof claims['appid'] === 'string' ? claims['appid'] : undefined;
    const aud = typeof claims['aud'] === 'string' ? claims['aud'] : undefined;
    const scp = typeof claims['scp'] === 'string' ? claims['scp'] : '';
    if (appid === undefined || aud === undefined) return;
    record(appid, aud, scp, req.url());
  });

  for (const url of OUTLOOK_URLS) {
    console.log(`probe: navigating to ${url}`);
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60_000 });
    } catch (e) {
      console.log(`probe: navigation to ${url} failed (${e instanceof Error ? e.message.split('\n')[0] : 'unknown'}), continuing`);
      continue;
    }
    console.log(`probe: settling ${settleMs}ms while the mailbox loads and calls its APIs...`);
    await page.waitForTimeout(settleMs);
    console.log(`probe: ${byAppId.size} appid(s) seen so far`);
    if ([...byAppId.values()].some((s) => s.audiences.has(GRAPH_AUD))) break;
  }

  await context.close();
  report();
  return 0;
};

process.exit(await run());
