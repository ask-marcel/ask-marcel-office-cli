#!/usr/bin/env bun
/*
 * Probe step 2: does the Outlook Web token actually READ a mailbox that is not
 * the signed-in user's own?
 *
 * Step 1 (`probe-outlook-scopes.ts`) established that appid
 * 9199bf20-a13f-4107-85dc-02114787ef48 emits a Graph-audience bearer carrying
 * `Mail.ReadWrite.Shared`. A granted scope is not an access grant: Exchange
 * still has to allow the delegation. This calls the exact endpoint the five
 * `list-shared-mailbox-*` commands call and reports what Graph answers.
 *
 * What it never does. Print a token, or print any mail content. Only the HTTP
 * status, Graph's error code, and a count of returned items are shown.
 *
 * Run:  bun run scripts/probe-outlook-shared-read.ts --mailbox <upn> [--mailbox <upn> ...]
 *
 * Pair each result with what the shipped CLI answers for the same mailbox:
 *   ask-marcel-office list-shared-mailbox-messages --user-id <upn> --top 1
 *
 * RESULT, 2026-09-06, one tenant, and it is the answer:
 *
 *   GET /users/{signed-in user's OWN upn}/messages
 *     Outlook Web token, Mail.ReadWrite.Shared present   403 ErrorAccessDenied
 *     the Teams token this CLI already carries           200, returns the message
 *
 * The Outlook token was refused on the caller's own mailbox, through the exact
 * endpoint the five `list-shared-mailbox-*` commands call, while the weaker
 * token the CLI already holds answers it. A scope in the JWT is therefore NOT
 * sufficient here: Exchange applies its own resource-side policy to that appid,
 * the same class of gate the ODSP allow-list applies on the file side.
 *
 * The conclusion for anyone tempted to retry this: harvesting another appid is
 * not the missing piece. `Mail.Read.Shared` being absent from the Teams token
 * was never the whole reason the shared-mailbox commands 403, so finding a
 * token that has it would not fix them. The commands' 2.5.0 wording, which
 * points at the group route instead, stands on this evidence.
 *
 * Not established: whether this token can read anything on Graph at all. The
 * `/me` ladder below answers that and needs one more capture, which five
 * attempts across 2026-09-06 did not get.
 *
 * On CAPTURING the Graph-audience token, since that is the hard part. The appid
 * emits bearers on nearly every load, but almost always for a NON-Graph
 * audience; the Graph-audience mint happened only on the first cold start of
 * the day. Waiting longer does not help, nor does reloading, nor does clearing
 * the origin's localStorage / sessionStorage / IndexedDB to force MSAL to
 * re-acquire (tried, and removed again for being dead weight). Treat the
 * Graph mint as opportunistic: run this first thing, against a profile that has
 * not opened Outlook yet that day.
 */

import { join } from 'node:path';
import { decodeJwtPayload } from '@ask-marcel/office-core/domain/jwt-utils.ts';

const OUTLOOK_APP_ID = '9199bf20-a13f-4107-85dc-02114787ef48';
const GRAPH_AUD = 'https://graph.microsoft.com';
const OUTLOOK_URL = 'https://outlook.office.com/mail/';

const mailboxes = process.argv.reduce<Array<string>>((acc, arg, i) => (arg === '--mailbox' && process.argv[i + 1] !== undefined ? [...acc, process.argv[i + 1] as string] : acc), []);
const profileDir = process.env['ASKMARCEL_BROWSER_PROFILE'] ?? join(process.env['HOME'] ?? '', '.ask-marcel', 'browser-profile');
const settleMs = Number(process.env['PROBE_SETTLE_MS'] ?? 45_000);

type Probe = { readonly mailbox: string; readonly status: number; readonly errorCode: string; readonly items: number };

const readMailbox = async (token: string, mailbox: string): Promise<Probe> => {
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/messages?$top=1&$select=id`;
  const response = await fetch(url, { headers: { authorization: `Bearer ${token}` } });
  const body: unknown = await response.json().catch(() => ({}));
  const error = (body as { error?: { code?: string; message?: string } }).error;
  const value = (body as { value?: ReadonlyArray<unknown> }).value;
  return { mailbox, status: response.status, errorCode: error?.code ?? '', items: Array.isArray(value) ? value.length : 0 };
};

const run = async (): Promise<number> => {
  if (mailboxes.length === 0) {
    console.error('probe: pass at least one --mailbox <upn>');
    return 1;
  }

  const { chromium } = await import('playwright');
  let context;
  for (const channel of ['msedge', 'chrome'] as const) {
    try {
      context = await chromium.launchPersistentContext(profileDir, { headless: false, channel, args: ['--disable-blink-features=AutomationControlled'] });
      break;
    } catch {
      // try the next channel
    }
  }
  if (context === undefined) {
    console.error('probe: could not launch msedge or chrome.');
    return 1;
  }

  let captured: string | undefined;
  const page = await context.newPage();
  context.on('request', (req) => {
    if (captured !== undefined) return;
    const auth = req.headers()['authorization'];
    if (typeof auth !== 'string' || !auth.startsWith('Bearer ')) return;
    const raw = auth.slice('Bearer '.length);
    const claims = decodeJwtPayload(raw);
    if (claims['appid'] !== OUTLOOK_APP_ID || claims['aud'] !== GRAPH_AUD) return;
    captured = raw;
    console.log(`probe: captured a Graph bearer for appid ${OUTLOOK_APP_ID} (token not shown)`);
  });

  for (const url of ['https://outlook.cloud.microsoft/mail/', OUTLOOK_URL]) {
    if (captured !== undefined) break;
    console.log(`probe: navigating to ${url} (window ${settleMs}ms)`);
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60_000 }).catch(() => undefined);
    const deadline = Date.now() + settleMs;
    while (captured === undefined && Date.now() < deadline) {
      await page.waitForTimeout(5000);
      if (captured === undefined && Date.now() < deadline) await page.reload({ waitUntil: 'domcontentloaded', timeout: 60_000 }).catch(() => undefined);
    }
    console.log(`probe: window closed for ${url}, captured=${captured !== undefined}`);
  }

  await context.close();

  if (captured === undefined) {
    console.error('probe: no Graph-audience bearer captured for the Outlook appid.');
    return 1;
  }

  console.log(`\n${'='.repeat(70)}\nWhat the Outlook Web token can reach on Graph\n${'='.repeat(70)}`);
  // Sanity ladder: if /me itself is refused the token is not a usable Graph
  // mail bearer at all, and a 403 on /users/{upn} says nothing about delegation.
  for (const path of ['/me', '/me/messages?$top=1&$select=id', '/me/mailFolders?$top=1&$select=id']) {
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, { headers: { authorization: `Bearer ${captured}` } });
    const body: unknown = await response.json().catch(() => ({}));
    const code = (body as { error?: { code?: string } }).error?.code ?? '';
    console.log(`  ${path.padEnd(44)} HTTP ${response.status}  ${response.status === 200 ? 'OK' : `REFUSED (${code || 'no code'})`}`);
  }
  for (const mailbox of mailboxes) {
    const result = await readMailbox(captured, mailbox);
    const verdict = result.status === 200 ? 'READ OK' : `REFUSED (${result.errorCode || 'no code'})`;
    console.log(`  ${`/users/${mailbox}/messages`.padEnd(44)} HTTP ${result.status}  ${verdict}`);
  }
  console.log('\nCompare each line with: ask-marcel-office list-shared-mailbox-messages --user-id <upn> --top 1');
  return 0;
};

process.exit(await run());
