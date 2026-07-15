# Metrics (measure whether you are improving)

Opinions about whether delivery is getting better are cheap; a small set of agreed metrics settles the argument. Measure the performance of the delivery system as a whole, on a recognized framework, so you compare against the industry instead of grading your own homework. Everything here is a **system metric**: it describes the pipeline and the process, and it improves when you fix the system, never when you push individuals harder.

## The four DORA metrics (the baseline)

- **Deployment frequency**: how often you ship.
- **Lead time for changes**: commit timestamp to production timestamp.
- **Change failure rate**: how often a release causes a problem.
- **Time to restore**: how fast you recover when it does.

Together they balance speed against stability, so you cannot game one by sacrificing the other. Derive them from real pipeline events, never from a hand-updated slide: the deploy job emits a machine-readable deployment event (`references/delivery.md`, The pipeline is the only deployer), and incidents emit an opened/resolved pair. Metrics you cannot recompute from raw events are anecdotes with decimals.

```yaml
# the deploy job's last step: one event per deployment, the metric source of truth
- name: record deployment event
  if: success()
  run: |
    curl -sf -X POST "$METRICS_EVENT_URL" -H 'content-type: application/json' \
      -d "{\"event_type\":\"deployment\",\"service\":\"app-api\",\"sha\":\"${GITHUB_SHA}\",\"deployed_at\":\"$(date -u +%FT%TZ)\"}"
```

## Flow metrics (where work stalls)

Pair the delivery metrics with flow: **cycle time** (clock time from work started to shipped), **throughput** (items shipped), and **work in progress**. Not story points: points are a local estimation currency that inflates, does not compare across teams, and hides queue time. Measure the clock.

```sql
SELECT date_trunc('week', done_at) AS week,
  percentile_cont(0.5) WITHIN GROUP (ORDER BY done_at - started_at) AS median_cycle_time,
  count(*) AS throughput
FROM issues WHERE done_at IS NOT NULL GROUP BY 1 ORDER BY 1;
```

## System metrics, never a stick

Group every metric by service or team, never by person. A per-developer leaderboard invites padded estimates, split PRs to juke the count, and silence about real problems: it measures fear, not throughput. The question is "is our delivery getting healthier", not "who is fastest this week".

## The trend, not the snapshot

A single reading means little; normal variation looks like a crisis or a win. Read the direction over weeks and alert only on a sustained shift (a 7-day average holding above threshold for a day), the same alert hygiene as `references/observability.md`. The trend is also the scoreboard for the standard itself: if the gates and disciplines are paying off, these four lines say so.

## Cost is a first-class metric

Cloud spend is a latency-shaped metric: measure it per service, alert on unexplained growth the way you alert on error rate, and design components to cost near nothing when idle, so scale brings revenue faster than it brings a bill. Discovering cost in the monthly invoice means the waste already ran for weeks.

```yaml
- alert: DailyCostAnomaly
  expr: avg_over_time(cloud_cost_usd_total[7d]) > 1.5 * avg_over_time(cloud_cost_usd_total[7d] offset 7d)
  for: 24h   # spend up >50% week-over-week, sustained, before anyone is paged
```

On AI endpoints the pre-call spend gate (`references/ai.md`, Cap what a caller can spend) prevents what this alert confirms; per-caller metering is what lets the dashboard read per-tenant truth instead of a blended bill.

## Review checklist

1. New deployable service: does its pipeline emit the deployment event, and do incidents emit opened/resolved pairs?
2. Is any metric being reported per person instead of per system? Regroup it.
3. Is anyone reacting to a single reading where the 7-day trend says otherwise?
4. Does the service have a cost line someone watches, and does it idle near zero?
5. Velocity in story points presented as a delivery metric? Replace with cycle time and throughput.
