// DECISION-GATE spike for eval item A (headless elevated-token re-capture).
//
// LESSONS 2026-05-06 records that `m365.cloud.microsoft` refused headless
// Playwright (anti-automation interstitial), which is why acquireElevatedToken
// launches HEADED today. The substrate path proves teams.microsoft.com tolerates
// headless. This spike answers, empirically and TODAY: does the elevated
// (M365ChatClient / OfficeHome, Graph-audience) bearer capture in HEADLESS mode,
// and on which host?
//
// It is READ-ONLY: it launches the real warm browser profile headless, navigates,
// and reads the outgoing Authorization bearer from network traffic. It writes
// nothing, sends nothing. Run on a machine with a warm login:
//
//   bun scripts/spike-headless-elevated.ts
//
// Verdict gates the feature: if either host captures -> build login --headless
// against that host. If NEITHER captures (and the profile is warm), the gotcha
// still holds -> do not ship the flag; fall back to the honesty-docs remedy.

import { rmSync } from 'node:fs';
import { join } from 'node:path';
import { chromium } from 'playwright';

const ELEVATED_APP_IDS = ['c0ab8ce9-e9a0-42e7-b064-33d422df41f1', '4765445b-32c6-49b0-83e6-1d93765276ca'];
const GRAPH_AUD = 'https://graph.microsoft.com';
const HOSTS = [
  { name: 'm365.cloud.microsoft', url: 'https://m365.cloud.microsoft/search' },
  { name: 'teams.microsoft.com', url: 'https://teams.microsoft.com/v2/' },
];
const BUDGET_MS = 30_000;
const POLL_MS = 1000;
const NAV_TIMEOUT_MS = 30_000;

const profileDir = process.env.ASKMARCEL_BROWSER_PROFILE ?? join(process.env.HOME ?? process.env.USERPROFILE ?? '', '.ask-marcel', 'browser-profile');
// Default headless (the mode under test). SPIKE_HEADED=1 runs the SAME probe
// headed, to tell an anti-automation headless block apart from cold cookies:
// if headed captures where headless did not, the block is headless-specific.
const HEADLESS = process.env.SPIKE_HEADED !== '1';

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

const decodeJwtPayload = (raw: string): Record<string, unknown> => {
  try {
    const segment = raw.split('.')[1] ?? '';
    const json = Buffer.from(segment.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8');
    return JSON.parse(json) as Record<string, unknown>;
  } catch {
    return {};
  }
};

const cleanupSingletonLocks = (): void => {
  for (const name of ['SingletonLock', 'SingletonCookie', 'SingletonSocket']) {
    try {
      rmSync(join(profileDir, name), { force: true });
    } catch {
      // ignore
    }
  }
};

const launchHeadless = async (): Promise<{ context: Awaited<ReturnType<typeof chromium.launchPersistentContext>>; channel: string }> => {
  for (const channel of ['msedge', 'chrome']) {
    try {
      const context = await chromium.launchPersistentContext(profileDir, { headless: HEADLESS, channel, args: ['--disable-blink-features=AutomationControlled'] });
      return { context, channel };
    } catch {
      // channel not installed — fall through
    }
  }
  const context = await chromium.launchPersistentContext(profileDir, { headless: HEADLESS, args: ['--disable-blink-features=AutomationControlled'] });
  return { context, channel: 'bundled' };
};

type Verdict = { host: string; channel: string; captured: boolean; appid?: string; note?: string };

const probe = async (host: { name: string; url: string }): Promise<Verdict> => {
  cleanupSingletonLocks();
  const { context, channel } = await launchHeadless();
  let captured: string | undefined;
  const page = await context.newPage();
  page.on('request', (req) => {
    if (captured !== undefined) return;
    const auth = req.headers()['authorization'];
    if (typeof auth !== 'string' || !auth.startsWith('Bearer ')) return;
    const claims = decodeJwtPayload(auth.slice('Bearer '.length));
    const appid = typeof claims['appid'] === 'string' ? claims['appid'] : undefined;
    const aud = typeof claims['aud'] === 'string' ? claims['aud'] : undefined;
    if (appid !== undefined && ELEVATED_APP_IDS.includes(appid) && aud === GRAPH_AUD) captured = appid;
  });

  let note: string | undefined;
  try {
    await page.goto(host.url, { waitUntil: 'domcontentloaded', timeout: NAV_TIMEOUT_MS });
  } catch (e) {
    note = `navigation error: ${e instanceof Error ? e.message : String(e)}`;
  }
  const deadline = Date.now() + BUDGET_MS;
  while (captured === undefined && Date.now() < deadline) {
    await sleep(POLL_MS);
  }
  const landedUrl = page.url();
  await context.close();
  const onLoginWall = landedUrl.includes('login.microsoftonline.com') || landedUrl.includes('login.live.com');
  return {
    host: host.name,
    channel,
    captured: captured !== undefined,
    ...(captured === undefined ? {} : { appid: captured }),
    ...(note !== undefined ? { note } : onLoginWall ? { note: `landed on a sign-in wall (${landedUrl.slice(0, 60)}) — profile cookies look COLD` } : {}),
  };
};

const main = async (): Promise<void> => {
  process.stdout.write(`=== headless elevated-capture spike ===\nprofile: ${profileDir}\nbudget:  ${BUDGET_MS / 1000}s per host\n\n`);
  const verdicts: Verdict[] = [];
  for (const host of HOSTS) {
    process.stdout.write(`probing ${host.name} (headless) ...\n`);
    const verdict = await probe(host);
    verdicts.push(verdict);
    const line = verdict.captured ? `CAPTURED (appid ${verdict.appid?.slice(0, 8)}, channel ${verdict.channel})` : `no capture (channel ${verdict.channel})`;
    process.stdout.write(`  ${verdict.host}: ${line}${verdict.note ? ` — ${verdict.note}` : ''}\n`);
  }

  const winner = verdicts.find((v) => v.captured);
  const cold = verdicts.some((v) => v.note?.includes('COLD'));
  process.stdout.write('\n=== verdict ===\n');
  if (winner !== undefined) {
    process.stdout.write(`HEADLESS ELEVATED CAPTURE WORKS on ${winner.host} — build login --headless against that host.\n`);
    process.exit(0);
  }
  if (cold) {
    process.stdout.write('No capture, but the profile looks COLD (landed on a sign-in wall). Re-run `ask-marcel-office login`, then this spike, before concluding.\n');
    process.exit(2);
  }
  process.stdout.write('No headless capture on either host with a warm profile — the 2026-05-06 gotcha still holds. Do NOT ship login --headless; use the honesty-docs fallback.\n');
  process.exit(1);
};

await main();
