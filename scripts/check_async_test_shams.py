#!/usr/bin/env python3
"""Fail CI when a [Fact]/[Theory]-decorated `async Task` test contains no `await`.

Catches a specific failure mode that's hit this repo twice: a test method
is declared `async Task` but its body never awaits anything. The intended
SUT call was never written; the test runs to completion green and verifies
nothing about the system. See round-3 cleanup (commit b36df3305) for
historical examples — five StripeWebhookServiceTests and four
CosmosSharePointSettingsServiceTests all fit this shape before deletion.

The structural signal is unambiguous: if you declared the method `async
Task`, you intended to await something. If you don't, either drop `async`
or finish the test.

Run:
    python3 scripts/check_async_test_shams.py

Exit codes:
    0 — clean
    1 — one or more violations (prints file:line per violation)
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

def _default_test_root() -> Path:
    """Find the test root relative to the script location.

    Checks `<repo>/tests`, `<repo>/test`, `<repo>/dotnet/tests` in that order.
    Override by passing the path as the first CLI argument.
    """
    repo_root = Path(__file__).resolve().parent.parent
    for candidate in ("tests", "test", "dotnet/tests"):
        p = repo_root / candidate
        if p.is_dir():
            return p
    return repo_root / "tests"


TEST_ROOT = Path(sys.argv[1]) if len(sys.argv) > 1 else _default_test_root()

# Match `public async Task MethodName(...) {`. We deliberately require
# `async Task` — sync `void` / `Task`-returning methods aren't the pattern
# we're guarding against. We confirm it's an xUnit test by walking
# backwards over attributes / comments / blank lines for a [Fact] or
# [Theory] marker (see is_xunit_test below).
METHOD_RE = re.compile(
    r'public\s+async\s+Task\s+(\w+)\s*\([^\)]*\)\s*\{',
    re.MULTILINE,
)

# [Fact] or [Theory] (with optional generic / argument list).
TEST_ATTR_RE = re.compile(r'\[(?:Fact|Theory)\b')
# [Fact(Skip = "...")] / [Theory(Skip = "...")] — skipped tests never run,
# so the lack of await is irrelevant. DOTALL so we tolerate the `Skip = "..."`
# argument being split across lines.
SKIPPED_ATTR_RE = re.compile(
    r'\[(?:Fact|Theory)\([^\]]*?\bSkip\s*=',
    re.DOTALL,
)


def is_xunit_test(text: str, method_decl_start: int) -> bool | None:
    """Walk backwards from a method declaration looking for [Fact]/[Theory].

    Returns True if this is an active xUnit test, False if it isn't a test,
    or None if the test is explicitly skipped (caller should treat as exempt).

    Collects the contiguous block of attribute lines and comments above the
    method, then matches against the whole block — so multi-line attributes
    like
        [Fact(Skip =
            "long reason")]
    are handled naturally. No fixed line cap; the first ordinary statement
    above terminates the block.
    """
    block_start = method_decl_start
    line_end = method_decl_start
    while line_end > 0:
        line_start = text.rfind("\n", 0, line_end - 1) + 1
        line = text[line_start:line_end].strip()
        if not line:
            line_end = line_start
            block_start = line_start
            continue
        # Lines that belong inside the attribute block:
        is_comment = line.startswith("//") or line.startswith("/*") or line.startswith("*")
        is_attribute_or_continuation = (
            line.startswith("[")
            or line.endswith(")]")
            or line.endswith(",")
            or (line.startswith('"') and line.endswith('"'))
            or (line.startswith('"') and ")]" in line)
        )
        if is_comment or is_attribute_or_continuation:
            line_end = line_start
            block_start = line_start
            continue
        break

    block = text[block_start:method_decl_start]
    if SKIPPED_ATTR_RE.search(block):
        return None
    if TEST_ATTR_RE.search(block):
        return True
    return False

# Tokens that count as "this test actually does async work". Any of them in
# the body is enough — we don't try to be clever about whether they're
# reached, lambda-scoped, etc. The goal is to catch tests that have ZERO
# async work, not to police every nuance.
AWAIT_TOKENS = (
    re.compile(r'\bawait\b'),
    re.compile(r'\.Result\b'),
    re.compile(r'\.Wait\s*\('),
    re.compile(r'\.GetAwaiter\s*\('),
    re.compile(r'\.AsTask\s*\('),
)


def find_method_body(text: str, open_brace_pos: int) -> tuple[int, int] | None:
    """Return (start, end) of the body including the outer braces."""
    depth = 0
    i = open_brace_pos
    while i < len(text):
        c = text[i]
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                return open_brace_pos, i + 1
        i += 1
    return None


def offset_to_line(text: str, offset: int) -> int:
    return text.count("\n", 0, offset) + 1


def scan_file(path: Path) -> list[tuple[Path, int, str]]:
    text = path.read_text(errors="ignore")
    violations: list[tuple[Path, int, str]] = []

    for match in METHOD_RE.finditer(text):
        test_status = is_xunit_test(text, match.start())
        if test_status is not True:
            continue
        method_name = match.group(1)
        open_brace = match.end() - 1
        body_range = find_method_body(text, open_brace)
        if body_range is None:
            continue
        body_start, body_end = body_range
        body = text[body_start:body_end]

        # Scan the raw body. We deliberately don't strip comments or strings:
        # URL literals ("https://...") and string-embedded "//" make comment
        # stripping treacherous, and false negatives from the literal word
        # "await" appearing inside a comment or string are negligible.
        if not any(tok.search(body) for tok in AWAIT_TOKENS):
            line = offset_to_line(text, match.start())
            violations.append((path, line, method_name))

    return violations


def main() -> int:
    if not TEST_ROOT.is_dir():
        print(f"error: test root not found: {TEST_ROOT}", file=sys.stderr)
        return 2

    all_violations: list[tuple[Path, int, str]] = []
    for cs in TEST_ROOT.rglob("*.cs"):
        s = str(cs)
        if "/bin/" in s or "/obj/" in s:
            continue
        all_violations.extend(scan_file(cs))

    if not all_violations:
        print("ok: every async Task test method awaits something")
        return 0

    print(
        f"FAIL: {len(all_violations)} async Task test method(s) contain no "
        f"await / .Result / .Wait() / .GetAwaiter() / .AsTask().\n"
        f"Either await the SUT call you intended to make, or drop `async` "
        f"and return `void`.\n",
        file=sys.stderr,
    )
    for path, line, method in all_violations:
        rel = path.relative_to(TEST_ROOT.parent)
        print(f"  {rel}:{line}: {method}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
