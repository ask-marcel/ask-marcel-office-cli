#!/usr/bin/env python3
"""The per-class mutation gate for atelier's PIT runs.

PIT's mutationThreshold judges the whole run's score, so a weak class passes beside strong ones
(found 2026-10-23, the same gap as Stryker's break). This reads the XML report the run just wrote
and holds each class to the same threshold, read from pom.xml. A nested class counts toward its
outer class, the unit pit-changed.sh targets.

  python3 scripts/check-pit-score.py             every class in target/pit-reports/mutations.xml
  python3 scripts/check-pit-score.py --selftest  planted reports, no PIT

The pom must ask PIT for XML (outputFormats XML and HTML, references/java-quarkus.md). Exit 0
passes; 1 is a class below the threshold, or a report or threshold that is missing, unreadable or
empty, never a pass.
"""
from __future__ import annotations

import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

REPORT = Path("target/pit-reports/mutations.xml")
POM = Path("pom.xml")


def threshold_of(pom_text: str) -> int | None:
    match = re.search(r"<mutationThreshold>\s*(\d+)\s*</mutationThreshold>", pom_text)
    return int(match.group(1)) if match else None


def scores_of(report_xml: str) -> dict[str, tuple[int, int]]:
    """(detected, total) per outer class: PIT's own score, detected over every mutation."""
    scores: dict[str, tuple[int, int]] = {}
    for mutation in ET.fromstring(report_xml).iter("mutation"):
        cls = (mutation.findtext("mutatedClass") or "").split("$", 1)[0]
        detected, total = scores.get(cls, (0, 0))
        scores[cls] = (detected + (mutation.get("detected") == "true"), total + 1)
    return scores


def judge(report_xml: str, pom_text: str) -> tuple[list[str], list[str]]:
    threshold = threshold_of(pom_text)
    if threshold is None:
        return [], ["pom.xml sets no <mutationThreshold>; atelier's is 90"]
    try:
        scores = scores_of(report_xml)
    except ET.ParseError as error:
        return [], [f"{REPORT} is not XML ({error})"]
    if not scores:
        return [], [f"{REPORT} holds no mutation: the run mutated nothing"]
    lines, failures = [], []
    for cls, (detected, total) in sorted(scores.items()):
        score = 100 * detected / total
        lines.append(f"  {cls}  {score:.2f}")
        if score < threshold:
            failures.append(f"{cls}  score={score:.2f}  required={threshold}")
    return lines, failures


def main() -> int:
    if not REPORT.is_file() or not POM.is_file():
        print(f"pit-score: needs {REPORT} and {POM}; the pom must ask PIT for XML and the run must have finished", file=sys.stderr)
        return 1
    lines, failures = judge(REPORT.read_text(encoding="utf-8"), POM.read_text(encoding="utf-8"))
    print("pit-score: per-class scores")
    print("\n".join(lines))
    if not failures:
        print("pit-score: every class meets the mutation threshold.")
        return 0
    print("\npit-score: per-class gate violations:", file=sys.stderr)
    for failure in failures:
        print(f"  {failure}", file=sys.stderr)
    print("\npit-score: strengthen the tests that pin these classes, never lower the threshold.", file=sys.stderr)
    return 1


def _mutation(cls: str, detected: bool) -> str:
    return f"<mutation detected='{str(detected).lower()}'><mutatedClass>{cls}</mutatedClass></mutation>"


def selftest() -> int:
    pom = "<project><mutationThreshold>90</mutationThreshold></project>"
    strong = [_mutation("app.domain.Strong", True)] * 18 + [_mutation("app.domain.Strong", False)] * 2
    weak = [_mutation("app.domain.Weak$Rules", True)] + [_mutation("app.domain.Weak", False)] * 2
    report = "<mutations>" + "".join(strong + weak) + "</mutations>"
    checks = [
        (judge(report, pom)[1] == ["app.domain.Weak  score=33.33  required=90"],
         f"a weak class beside a strong one (the run totals 82.61) gave {judge(report, pom)[1]}, not Weak alone, nested folded in"),
        (judge("<mutations/>", pom)[1][0].endswith("mutated nothing"), "an empty report passed"),
        (judge(report, "<project/>")[1][0].startswith("pom.xml sets no"), "a pom with no threshold passed"),
        (judge("<mutations>", pom)[1][0].endswith(")") and "not XML" in judge("<mutations>", pom)[1][0], "a truncated report passed"),
    ]
    failed = [message for ok, message in checks if not ok]
    for message in failed:
        print(f"pit-score: selftest FAIL: {message}", file=sys.stderr)
    if failed:
        return 1
    print("pit-score: selftest OK: a weak class fails beside a strong one at the line, a nested class counts "
          "toward its outer one, an empty report, a pom with no threshold and a truncated report fail")
    return 0


if __name__ == "__main__":
    sys.exit(selftest() if "--selftest" in sys.argv[1:] else main())
