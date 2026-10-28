#!/usr/bin/env python3
"""Hard rule 38's host half: the settings git alone cannot hold, checked on GitHub.

Rule 38 lands a branch by rebase, deletes it the moment it lands, and lets nobody delete main,
force-push it or put a merge commit on it. Git enforces none of that; the host does once its owner
sets it, and a setting set by hand drifts unseen. This check reads the repository on GitHub and
fails on each finding:

  DEFAULT  the default branch is not main;
  RULE     main lacks an active rule of a type the ruleset file lists (host-ruleset.json:
           deletion, non_fast_forward, required_linear_history);
  MERGE    merge commits or squash merges are allowed, rebase merges are not, or a merged head
           branch is not deleted automatically.

It needs read access only. REST hides the merge settings from a read-only token, so they come from
GraphQL (probed 2026-10-15); a setting GitHub still does not show is an error, never a pass.

Adapted from the atelier asset, without --apply (the owner sets the host by hand, CONTRIBUTING.md,
Landing a branch) and without the HOST_RULESET opt-out (the repository is public).

  python3 scripts/check-host.py             check (branches.yml runs it daily)
  python3 scripts/check-host.py --selftest  prove it fails on each drift

Env: GITHUB_REPOSITORY (default: gh repo view), GH_TOKEN, RULESET (default: host-ruleset.json beside
this script). Exit 0 passes, 1 is a finding, 2 is a host or ruleset it could not read. Needs gh.
"""
from __future__ import annotations

import contextlib
import io
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

RULE = "Hard rule 38 (atelier references/workflow.md, Branch lifecycle)."
# Each merge setting: its REST name (what the settings page sets), its GraphQL name (what the check
# reads) and the value rule 38 wants.
MERGE = {
    "allow_merge_commit": ("mergeCommitAllowed", False),
    "allow_squash_merge": ("squashMergeAllowed", False),
    "allow_rebase_merge": ("rebaseMergeAllowed", True),
    "delete_branch_on_merge": ("deleteBranchOnMerge", True),
}


class HostError(Exception):
    """The host or the ruleset file could not be read: exit 2, never a pass."""


class GitHub:
    """The repository as gh sees it. The selftest swaps in a fake with the same methods."""

    def __init__(self, repo: str) -> None:
        self.repo = repo

    def _api(self, *args: str) -> object:
        try:
            done = subprocess.run(["gh", "api", *args], capture_output=True, text=True, encoding="utf-8")
        except FileNotFoundError:
            raise HostError("gh is not on PATH (https://cli.github.com)") from None
        if done.returncode != 0:
            raise HostError(f"gh api {args[0]}: {(done.stderr or done.stdout).strip()}")
        return json.loads(done.stdout) if done.stdout.strip() else None

    def settings(self) -> dict:
        owner, name = self.repo.split("/", 1)
        fields = " ".join(graphql for graphql, _ in MERGE.values())
        query = ("query($owner: String!, $name: String!) { repository(owner: $owner, name: $name) "
                 f"{{ defaultBranchRef {{ name }} {fields} }} }}")
        data = self._api("graphql", "-f", f"query={query}", "-f", f"owner={owner}", "-f", f"name={name}")
        repo = ((data or {}).get("data") or {}).get("repository") if isinstance(data, dict) else None
        if not repo:
            raise HostError(f"GitHub returned no repository {self.repo}: {json.dumps(data)[:200]}")
        return {"default": (repo.get("defaultBranchRef") or {}).get("name"),
                **{rest: repo.get(graphql) for rest, (graphql, _) in MERGE.items()}}

    def rules_on_main(self) -> set:
        return {rule.get("type") for rule in self._api(f"repos/{self.repo}/rules/branches/main") or []}


def load_ruleset(path: Path) -> dict:
    """The ruleset main is held to. A file naming no rule would pass every repository: an error."""
    try:
        body = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise HostError(f"cannot read the ruleset {path}: {error}") from None
    rules = body.get("rules") if isinstance(body, dict) else None
    kinds = [rule.get("type") for rule in rules or [] if isinstance(rule, dict)]
    if not kinds or not body.get("name") or not all(isinstance(kind, str) and kind for kind in kinds):
        raise HostError(f"the ruleset {path} names no rule to hold main to")
    return body


def check(host: GitHub, ruleset: dict) -> list:
    """The findings, each failing the run."""
    seen = host.settings()
    findings = []
    if seen["default"] != "main":
        findings.append(f"DEFAULT  the default branch is {seen['default']}, not main")
    active = host.rules_on_main()
    findings += [f"RULE  main has no active {rule['type']} rule" for rule in ruleset["rules"]
                 if rule["type"] not in active]
    for rest, (graphql, wanted) in MERGE.items():
        if not isinstance(seen[rest], bool):
            raise HostError(f"GitHub did not show {graphql} to this token, so {rest} cannot be checked")
        if seen[rest] is not wanted:
            findings.append(f"MERGE  {rest} is {json.dumps(seen[rest])}, rule 38 wants {json.dumps(wanted)}")
    return findings


def run(host: GitHub, ruleset: dict) -> int:
    findings = check(host, ruleset)
    if findings:
        for finding in findings:
            print(f"  ╳ {finding}", file=sys.stderr)
        print(f"\n  {len(findings)} finding(s) on {host.repo}: the owner sets them in the repository settings "
              f"(CONTRIBUTING.md, Landing a branch).\n  {RULE}", file=sys.stderr)
        return 1
    print(f"check-host: {host.repo}: main is the default branch, protected by the ruleset; rebase merges only, "
          "head branches deleted on merge")
    return 0


def current_repo() -> str:
    try:
        done = subprocess.run(["gh", "repo", "view", "--json", "nameWithOwner", "--jq", ".nameWithOwner"],
                              capture_output=True, text=True, encoding="utf-8")
    except FileNotFoundError:
        done = None
    if done is None or done.returncode != 0 or "/" not in done.stdout:
        raise HostError("cannot tell which repository to check: set GITHUB_REPOSITORY=owner/name")
    return done.stdout.strip()


def main(argv: list[str]) -> int:
    # The findings carry a box-drawing cross, which a cp1252 console cannot print (2026-10-09).
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")
    if argv == ["--selftest"]:
        return selftest()
    if argv:
        print("usage: check-host.py [--selftest]", file=sys.stderr)
        return 2
    try:
        ruleset = load_ruleset(Path(os.environ.get("RULESET") or Path(__file__).resolve().parent / "host-ruleset.json"))
        host = GitHub(os.environ.get("GITHUB_REPOSITORY") or current_repo())
        return run(host, ruleset)
    except HostError as error:
        print(f"check-host: {error}", file=sys.stderr)
        return 2


class FakeGitHub:
    """A repository in memory. Its rules sit on the default branch, as the ruleset targets it."""

    def __init__(self, default: str = "main", rules: tuple = (), merge: dict | None = None) -> None:
        self.repo, self.default, self.rules = "acme/widget", default, set(rules)
        self.merge = dict(merge if merge is not None else {rest: want for rest, (_, want) in MERGE.items()})

    def settings(self) -> dict:
        return {"default": self.default, **self.merge}

    def rules_on_main(self) -> set:
        return set(self.rules) if self.default == "main" else set()


def selftest() -> int:
    failures: list[str] = []
    ruleset = load_ruleset(Path(__file__).resolve().parent / "host-ruleset.json")
    kinds = [rule["type"] for rule in ruleset["rules"]]
    good = {rest: want for rest, (_, want) in MERGE.items()}

    def outcome(host: FakeGitHub) -> tuple[int, str]:
        said = io.StringIO()
        with contextlib.redirect_stdout(said), contextlib.redirect_stderr(said):
            try:
                code = run(host, ruleset)
            except HostError as error:
                print(f"check-host: {error}")
                code = 2
        return code, said.getvalue()

    def expect(what: str, got: tuple[int, str], code: int, *held: str, absent: tuple = ()) -> None:
        if got[0] != code or any(text not in got[1] for text in held) or any(text in got[1] for text in absent):
            failures.append(f"{what}: exit {got[0]}, wanted {code}:\n{got[1]}")

    # Each drift fails for its own reason, and a setting GitHub does not show is an error, not a pass.
    expect("a held repository", outcome(FakeGitHub(rules=kinds)), 0, "protected by the ruleset")
    expect("master as the default", outcome(FakeGitHub("master", rules=kinds)), 1,
           "DEFAULT  the default branch is master")
    expect("no ruleset", outcome(FakeGitHub()), 1, *(f"no active {kind} rule" for kind in kinds))
    for kind in kinds:
        expect(f"main without {kind}", outcome(FakeGitHub(rules=[k for k in kinds if k != kind])), 1,
               f"no active {kind} rule", absent=tuple(f"no active {k} rule" for k in kinds if k != kind))
    for rest, (graphql, want) in MERGE.items():
        expect(f"{rest} flipped", outcome(FakeGitHub(rules=kinds, merge={**good, rest: not want})), 1,
               f"MERGE  {rest} is {json.dumps(not want)}")
        expect(f"{rest} not shown", outcome(FakeGitHub(rules=kinds, merge={**good, rest: None})), 2,
               f"did not show {graphql}")
    # The input: ruleset files that hold main to nothing.
    with tempfile.TemporaryDirectory() as tmp:
        for text in ('{"name": "main", "rules": []}', '{"rules": [{"type": "deletion"}]}', "[]", "not json"):
            (Path(tmp) / "ruleset.json").write_text(text, encoding="utf-8")
            try:
                load_ruleset(Path(tmp) / "ruleset.json")
                failures.append(f"the ruleset {text!r} was accepted")
            except HostError:
                pass

    # The real class: the commands gh receives, with subprocess.run stubbed. The GraphQL answer is
    # the atelier probe's (2026-10-15): GitHub's defaults, five findings with one rule set.
    calls: list = []
    answers = {
        "graphql": '{"data": {"repository": {"defaultBranchRef": {"name": "main"}, "isPrivate": false, '
                   '"rebaseMergeAllowed": true, "squashMergeAllowed": true, "mergeCommitAllowed": true, '
                   '"deleteBranchOnMerge": false}}}',
        "repos/acme/widget/rules/branches/main": '[{"type": "deletion"}]',
    }

    def stub(cmd: list, **_: object) -> subprocess.CompletedProcess:
        calls.append(cmd[2:])
        return subprocess.CompletedProcess(cmd, 0, answers.get(cmd[2], ""), "")

    real, subprocess.run = subprocess.run, stub
    try:
        host = GitHub("acme/widget")
        seen = host.settings()
        if seen != {"default": "main", **{rest: rest != "delete_branch_on_merge" for rest in MERGE}}:
            failures.append(f"the probe's GraphQL answer read as {seen}")
        findings = check(host, ruleset)
        if len(findings) != 5:
            failures.append(f"the probe's answer gave {len(findings)} findings, wanted 5: {findings}")
    finally:
        subprocess.run = real
    if [args[:1] for args in calls] != [["graphql"], ["graphql"], ["repos/acme/widget/rules/branches/main"]]:
        failures.append("gh received:\n" + "\n".join(map(str, calls)))
    # GitHub's names, spelled once more: the probe answers true for three, so a swap would read the same.
    spelled = {"allow_merge_commit": "mergeCommitAllowed", "allow_squash_merge": "squashMergeAllowed",
               "allow_rebase_merge": "rebaseMergeAllowed", "delete_branch_on_merge": "deleteBranchOnMerge"}
    if {rest: graphql for rest, (graphql, _) in MERGE.items()} != spelled:
        failures.append(f"the merge settings map to {MERGE}, not GitHub's names")
    query = calls[0]
    if not all(f"{graphql} " in query[2] for graphql, _ in MERGE.values()) or query[3:] != ["-f", "owner=acme", "-f", "name=widget"]:
        failures.append(f"the GraphQL call was {query}")

    for failure in failures:
        print(f"check-host.py --selftest: {failure}")
    if failures:
        return 1
    print("check-host.py --selftest: each drift fails for its own reason (the default branch, each rule, each "
          "merge setting), a setting GitHub does not show and a ruleset naming no rule are errors, and gh "
          "receives the documented calls")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
