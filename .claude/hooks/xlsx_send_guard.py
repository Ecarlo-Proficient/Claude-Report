#!/usr/bin/env python3
"""PreToolUse guard: block SendUserFile if any .xlsx/.xlsm it carries would trip
Excel's "we found a problem with some content" repair prompt. Runs the repo's
shared/xlsx_verify.py on each Excel file being sent. Harness-enforced — it does
NOT depend on the model remembering to call assert_clean.

Fail-closed on Excel: a file that the verifier flags OR that isn't a readable
xlsx is blocked. Non-Excel files and unparseable hook input pass through.
"""
import json
import os
import subprocess
import sys

# Self-locate the repo (this file lives at <repo>/.claude/hooks/), so the guard
# is portable across clones. CLAUDE_PROJECT_DIR wins when the harness sets it.
REPO = os.environ.get("CLAUDE_PROJECT_DIR") or os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
VERIFIER = os.path.join(REPO, "shared", "xlsx_verify.py")


def main():
    try:
        data = json.load(sys.stdin)
    except Exception:
        return  # can't parse -> don't block
    files = (data.get("tool_input") or {}).get("files") or []
    if isinstance(files, str):
        files = [files]
    bad = []
    for f in files:
        if not str(f).lower().endswith((".xlsx", ".xlsm")):
            continue
        try:
            r = subprocess.run([sys.executable, VERIFIER, str(f)],
                               capture_output=True, text=True, timeout=60)
            if r.returncode != 0 or "ISSUES" in r.stdout:
                bad.append(f"{f}:\n{(r.stdout or r.stderr).strip()}")
        except Exception as e:
            bad.append(f"{f}: verifier could not run ({e})")
    if bad:
        reason = ("BLOCKED: this Excel file would trip Excel's repair prompt. "
                  "Do NOT deliver it. Fix the cause (most often an Excel Table "
                  "whose ref/autoFilter points past the last row after a row "
                  "insert/delete — reset the table ref or drop the table), then "
                  "re-run shared.xlsx_verify.assert_clean before sending.\n\n"
                  + "\n".join(bad))
        print(json.dumps({"hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": "deny",
            "permissionDecisionReason": reason,
        }}))


if __name__ == "__main__":
    main()
