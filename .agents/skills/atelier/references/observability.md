# Observability (you cannot fix what you cannot see)

A running system should explain itself. In production the question is never just "is something wrong" but "what, where, and since when". The `Logger` port (hard rule 4) is the floor; this reference covers the rest: reliability targets as numbers, correlated traces and metrics on an open standard, behaviour instrumentation, and alerts that only fire when a human must act.

## Reliability targets are numbers (SLOs)

"Reliable enough" is a feeling until it is a number someone agreed to. Write down the objective and alert against its budget:

```yaml
# thresholds.yaml (the same file governance.md points stakeholders at)
slo:
  availability:   { target: 0.999, window: 30d }   # error budget: 0.1%
  latency_p95_ms: { target: 300, route: /v1/invoices }
  latency_p99_ms: { target: 800, route: /v1/invoices }
  error_rate:     { target: 0.005, window: 30d }
```

Every target has three parts: a metric, a number, a window. Load tests prove the latency budget pre-ship (`references/reliability.md`, Performance is a budget); alerts watch it post-ship (below).

## Instrument on an open standard, correlated

Logs, metrics, and traces join up only when they share a trace id. Use OpenTelemetry, not a vendor SDK: the exporter is per-environment config, so no tool lock-in (`references/delivery.md`, Rent open standards).

Where it lives in the layout: the OTel SDK is infrastructure. Domain and use-cases stay clean; spans are added at the edges the same way logging is, with a higher-order wrapper owned by composition:

```ts
// src/infra/telemetry.ts: the only file importing @opentelemetry/*
import { trace } from '@opentelemetry/api';

export type WithSpan = <T>(name: string, attrs: Record<string, string>, run: () => Promise<T>) => Promise<T>;

export const createWithSpan = (tracerName: string): WithSpan => {
  const tracer = trace.getTracer(tracerName);
  return (name, attrs, run) =>
    tracer.startActiveSpan(name, async (span) => {
      for (const [k, v] of Object.entries(attrs)) span.setAttribute(k, v);
      try {
        return await run();
      } finally {
        span.end();
      }
    });
};
```

```ts
// composition: wrap each use-case once; the use-case itself never imports telemetry
const placeOrder = (input: OrderInput): Promise<Result<Summary, StepError>> =>
  withSpan('place-order', { 'org.id': orgId }, () => placeOrderUseCase(deps)(input));
```

Attribute discipline: opaque internal ids only, never natural identifiers (hard rule 27; `references/privacy.md`). A `user.id` UUID correlates; an email in a span is a leak with retention.

The Winston adapter joins the same story by emitting the active trace id on every line (an OTel format or a small custom format), so a log line found by text search links to its trace.

Java: Quarkus ships OTel; `@WithSpan` on the application service or the adapter, Micrometer for metrics (`references/java-quarkus.md`).

## Watch behaviour, not just health

CPU and heap graphs stay green while the funnel collapses. Instrument what the product does, split by outcome, so dashboards answer "are users succeeding":

```ts
// src/infra/metrics.ts: counters + histogram behind a Metrics port
const started = meter.createCounter('checkout.started');
const completed = meter.createCounter('checkout.completed');
const latency = meter.createHistogram('checkout.latency.ms');

export const recordCheckout = (ok: boolean, ms: number): void => {
  started.add(1);
  if (ok) completed.add(1);
  latency.record(ms, { outcome: ok ? 'ok' : 'abandoned' });
};
```

The same data serves reliability and product decisions: call volumes, latencies, drop-off. That is also what makes "keep or kill on adoption" measurable (`references/product.md`, Keep validating after launch).

Frontend (Next.js): a render error that blanks the page must reach the backend, not `console`. An error boundary reports to telemetry with the session id, and web vitals (LCP, CLS, INP) go out tagged the same way, so a slow page correlates to its backend traces.

```tsx
<ErrorBoundary onError={(e, info) => telemetry.captureException(e, { info, sessionId })}>
  <App />
</ErrorBoundary>
```

(The boundary and the telemetry client live in `src/lib/`, wired by page shells; the design system stays logic-free, rule 21.)

## Alert on what matters

- **Symptom-based, budget-tied.** Alert on error rate and latency percentiles (p95/p99: averages hide the worst experiences) crossing the SLO budget, sustained over a window (`for: 10m`), not on raw CPU.
- **Page only when a human must act.** Everything else is a dashboard or a ticket.
- **Alert hygiene is a deletion rule.** An alert that pages twice without prompting a fix gets automated away or deleted; a pager everyone ignores protects nothing.
- **After an incident slips through, fix what broke**, not only the alert that missed it. The postmortem owns that (`references/delivery.md`, Blameless postmortems).

```yaml
- alert: CheckoutErrorBudgetBurn
  expr: sum(rate(http_requests_total{route="/checkout",status=~"5.."}[5m]))
        / sum(rate(http_requests_total{route="/checkout"}[5m])) > 0.02
  for: 10m            # sustained, not a blip
  labels: { severity: page }
```

Anomaly-style rules (week-over-week growth, sustained trend windows) beat static thresholds for metrics that drift, cost above all (`references/metrics.md`, Cost is a first-class metric).

## What good looks like

When something breaks you find the cause in minutes from a dashboard: the alert names the symptom, the trace shows the failing hop, the correlated logs show why, and the metric shows since when. If diagnosing an incident required adding new instrumentation first, that gap is the postmortem's first action item.

## Review checklist (changes touching production paths)

1. Can one request be followed end to end (trace id present, logs correlated)?
2. New user-facing flow: outcome counters and a latency histogram, split by success/abandon?
3. Any new attribute or log field: opaque id, not a natural identifier? (rule 27)
4. New SLO-relevant route: threshold recorded as a number with a window; alert tied to the budget?
5. Would this page a human for something no human needs to act on?
