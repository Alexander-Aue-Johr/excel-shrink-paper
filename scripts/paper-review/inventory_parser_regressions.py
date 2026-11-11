"""Inventory commit messages and Rust test declarations at a pinned checkout."""

import argparse
import collections
import hashlib
import json
import pathlib
import re
import subprocess


LEXICAL_MODULES = {
    "excel-shrink-core/src/shared_strings_tests.rs",
    "excel-shrink-core/src/sheet/ranges_regression_tests.rs",
    "excel-shrink-core/src/sheet/streaming.rs",
    "excel-shrink-core/src/workbook.rs",
    "excel-shrink-core/src/xml.rs",
    "excel-shrink-core/src/xml/namespaces.rs",
    "excel-shrink-core/src/xml/namespaces/header.rs",
    "excel-shrink-cli/src/regression/streaming_xml_tests.rs",
}
PRESERVATION_SUITES = {
    f"excel-shrink-core/src/sheet/{name}_regression_tests.rs"
    for name in ("cells", "rows", "tail_fill")
}
TEST_DECLARATION = re.compile(
    r"#\[(?:test|(?:tokio|async_std)::test)(?:\([^\]]*\))?\]"
    r"\s*(?:#\[[^\]]*\]\s*)*(?:async\s+)?fn\s+(\w+)"
)


def git(repository, *arguments):
    """Read Git output without changing the checkout or its refs."""
    return subprocess.check_output(
        ["git", *arguments], cwd=repository
    ).decode("utf-8", errors="replace")


def suite_for(path):
    """Assign a primary source-module group, independent of test outcomes."""
    if path in LEXICAL_MODULES:
        return "lexical modules"
    if path in PRESERVATION_SUITES:
        return "cell, row and tail suites"
    if path.startswith("excel-shrink-core/src/"):
        return "other core"
    return "other Rust tests"


def inventory_tests(repository):
    """Count tracked Rust test functions once, including ignored declarations."""
    tests = []
    for path in git(repository, "ls-files", "*.rs").splitlines():
        contents = (repository / path).read_bytes()
        source = contents.decode("utf-8")
        digest = hashlib.sha256(contents).hexdigest()
        for match in TEST_DECLARATION.finditer(source):
            tests.append({
                "file": path,
                "line": source.count("\n", 0, match.start()) + 1,
                "name": match.group(1),
                "source_sha256": digest,
                "primary_suite": suite_for(path),
            })
    return tests


def inventory_commits(repository):
    """Retain complete messages for every commit reachable from local refs."""
    records = git(repository, "log", "--all", "--format=%H%x1f%aI%x1f%B%x1e")
    commits = []
    for record in records.split("\x1e"):
        fields = record.strip().split("\x1f", 2)
        if len(fields) == 3:
            commits.append(dict(zip(("commit", "date", "message"), fields)))
    return commits


def write_json(path, value):
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def main():
    arguments = argparse.ArgumentParser(description=__doc__)
    arguments.add_argument("repository", type=pathlib.Path)
    arguments.add_argument("output", type=pathlib.Path)
    options = arguments.parse_args()
    tests = inventory_tests(options.repository)
    commits = inventory_commits(options.repository)
    options.output.mkdir(parents=True, exist_ok=True)
    write_json(options.output / "rust-tests.json", tests)
    write_json(options.output / "commit-messages.json", commits)
    write_json(options.output / "inventory.json", {
        "source_revision": git(options.repository, "rev-parse", "HEAD").strip(),
        "commit_scope": "All commits reachable from locally available refs (git log --all).",
        "commit_messages": len(commits),
        "rust_test_declarations": len(tests),
        "counting_rule": "One Rust test function per declaration, including ignored tests. Fixture iterations are counted once. Counts describe source, not executed outcomes.",
        "suite_counts": dict(collections.Counter(test["primary_suite"] for test in tests)),
        "files": dict(collections.Counter(test["file"] for test in tests)),
        "refs": git(options.repository, "show-ref").splitlines(),
    })


if __name__ == "__main__":
    main()