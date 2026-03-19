#!/usr/bin/env python3
"""
Git filter-branch --msg-filter: remove Co-authored-by lines for Cursor and Claude/Anthropic
so they no longer appear as GitHub contributors from those trailers.
Reads commit message on stdin, writes filtered message on stdout.
"""
import sys

def should_drop_line(line: str) -> bool:
    low = line.lower().strip()
    if not low.startswith("co-authored-by:"):
        return False
    if "cursor" in low or "cursoragent" in low:
        return True
    if "anthropic.com" in low or "claude" in low:
        return True
    return False


def main() -> None:
    text = sys.stdin.read()
    lines = text.splitlines(keepends=False)
    out_lines = [ln for ln in lines if not should_drop_line(ln)]
    result = "\n".join(out_lines)
    # Preserve trailing newline if original had one
    if text.endswith("\n"):
        result += "\n"
    sys.stdout.write(result)


if __name__ == "__main__":
    main()
