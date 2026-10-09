"""
sync_view.py - the progress view for sync-all's two ledger steps, in the same look as
sync-ap and sync-ar (owner 2026-09-24: "3 & 4 is giving too much info ... i need it like
the 1/2"). A progress summary, never the data: one line per step with a count and a
check-mark. The full raw output still goes to the log, and a failure writes a crash
report with everything the step printed.

    python ledger/sync_view.py mirror [--reconcile]   step 0: refresh the QBO mirror
                                                       (+ the Sunday reconcile)
    python ledger/sync_view.py reload                  step 3: every ledger loader
                                                       (runs reload_ledger.sh, plain)

The steps themselves live where they always did (refresh_mirror.py, reload_ledger.sh);
this only reads their stdout. Colors switch off when output isn't a terminal.
Log: ~/Library/Logs/Proficient/ledger-sync/run.log
"""
from __future__ import annotations

import datetime
import os
import platform
import re
import subprocess
import sys
from collections import deque
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
LOG_DIR = Path(os.environ.get("LOG_DIR") or (Path.home() / "Library/Logs/Proficient/ledger-sync"))

_TTY = sys.stdout.isatty() and os.environ.get("NO_COLOR") is None


def _c(code: str) -> str:
    return code if _TTY else ""


RESET = _c("\033[0m"); BOLD = _c("\033[1m"); DIM = _c("\033[90m")
RED = _c("\033[31m"); GREEN = _c("\033[32m"); BLUE = _c("\033[94m"); BR_CYAN = _c("\033[96m")
CHECK = f"{GREEN}✓{RESET}"
CROSS = f"{RED}✗{RESET}"

# reload_ledger.sh step label -> (icon, label shown)
ICONS = {
    "WIP master": ("📊", "WIP master"),
    "Bills": ("📥", "Bills"),
    "Invoices": ("🧾", "Invoices"),
    "Customers": ("👥", "Customers"),
    "Costs": ("🧱", "Costs"),
    "Payments": ("💳", "Payments received"),
    "Bill payments": ("💵", "Bill payments"),
    "Sub LOC": ("🔁", "Sub LOC"),
    "Health": ("🩺", "Company health"),
    "Attachments": ("📎", "Attachments"),
}

STEP_RE = re.compile(r"^-- (.+?) -> ledger --$")
EXIT_RE = re.compile(r"^\[step-exit (\d+)\]$")
ANSI_RE = re.compile(r"\x1b\[[0-9;]*m")
# "Wrote 151 projects + 151 snapshots -> ..." / "wrote 860 payments + ..." -> "151 projects"
WROTE_RE = re.compile(r"\b[Ww]rote\s+([\d,]+)\s+(.+?)(?=\s+[+·(]|\s+->|\s+→|\s+to\s|[.]?$)")
MIRROR_RE = re.compile(r"^done:.*?([\d,]+) written\s*·\s*([\d,]+) deleted")

# the loaders' own wording -> plain words for the view
NOUNS = (("billing_event rows", "invoices"), ("snapshot payload(s)", "snapshots"),
         ("attachment rows", "attachments"), ("(s)", "s"))


def _n(num: str) -> str:
    try:
        return f"{int(num.replace(',', '')):,}"
    except ValueError:
        return num


def detail_for(lines: list[str]) -> str:
    """A short count for the step line - never dollars or findings."""
    for s in lines:
        m = MIRROR_RE.search(s)
        if m:
            return f"{_n(m.group(1))} changed · {_n(m.group(2))} deleted"
    for s in lines:
        m = WROTE_RE.search(s)
        if m:
            noun = m.group(2).strip()
            for raw, plain in NOUNS:
                noun = noun.replace(raw, plain)
            return f"{_n(m.group(1))} {noun.strip()}"
    return ""


def error_for(lines: list[str]) -> str:
    """The line that says why a step failed: the exception, else the last ✗/error line."""
    for s in reversed(lines):
        if re.match(r"^[A-Za-z_.]+(Error|Exception)\b", s) or s.startswith(("✗", "ERROR", "FATAL")):
            return s[:160]
    return lines[-1][:160] if lines else "no output"


class View:
    def __init__(self, title: str):
        self.title = title
        self.buffer: deque[str] = deque(maxlen=400)
        self.failures: list[tuple[str, str]] = []

    def _w(self, s: str = "") -> None:
        sys.stdout.write(s + "\n"); sys.stdout.flush()

    def header(self) -> None:
        bar = "─" * max(2, 47 - len(self.title))       # same width as the closing line
        self._w(f"\n{BOLD}{BLUE}╭─ {self.title} {bar}╮{RESET}")

    def working(self, icon: str, label: str) -> None:
        if _TTY:
            sys.stdout.write(f"  {icon}  {label}  {DIM}…{RESET}"); sys.stdout.flush()

    def done(self, icon: str, label: str, detail: str, ok: bool, why: str = "") -> None:
        if _TTY:
            sys.stdout.write("\r\033[K")
        det = f"   {BOLD}{BR_CYAN}{detail}{RESET}" if detail and ok else ""
        self._w(f"  {icon}  {label}{det}  {CHECK if ok else CROSS}")
        if not ok:
            self._w(f"     {RED}{why}{RESET}")
            self.failures.append((label, why))

    def close(self, elapsed: str) -> None:
        self._w(f"{BOLD}{BLUE}╰──────────────────────────────────────────────────╯{RESET}  {DIM}done in {elapsed}{RESET}\n")


def run_stream(cmd: list[str], view: View, logf, on_line) -> int:
    env = dict(os.environ, ACB_PLAIN="1", ACB_DRIVES_CHECKED="1", PYTHONUNBUFFERED="1")
    proc = subprocess.Popen(cmd, cwd=str(REPO), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            text=True, bufsize=1, env=env)
    try:
        for raw in iter(proc.stdout.readline, ""):
            logf.write(raw); logf.flush()
            line = ANSI_RE.sub("", raw.rstrip("\n"))
            view.buffer.append(line)
            on_line(line.strip())
    except KeyboardInterrupt:
        proc.terminate()
        view.failures.append(("run", "interrupted (Ctrl-C)"))
    return proc.wait()


def mirror(view: View, logf, reconcile: bool) -> int:
    rc = 0
    steps = [("🔄", "Change feed", [])] + ([("🔍", "Sunday reconcile", ["--reconcile"])] if reconcile else [])
    for icon, label, extra in steps:
        seen: list[str] = []
        view.working(icon, label)
        r = run_stream([sys.executable, "-u", str(HERE / "refresh_mirror.py"), *extra], view, logf,
                       lambda s: s and seen.append(s))
        view.done(icon, label, detail_for(seen) or "up to date", r == 0, "" if r == 0 else error_for(seen))
        rc = rc or r
        if r:
            break                      # never reconcile on top of a failed refresh
    return rc


def reload(view: View, logf) -> int:
    state = {"label": None, "lines": []}

    def on_line(s: str) -> None:
        m = STEP_RE.match(s)
        if m:
            state["label"], state["lines"] = m.group(1), []
            view.working(*ICONS.get(m.group(1), ("•", m.group(1))))
            return
        m = EXIT_RE.match(s)
        if m and state["label"]:
            icon, label = ICONS.get(state["label"], ("•", state["label"]))
            ok = m.group(1) == "0"
            view.done(icon, label, detail_for(state["lines"]), ok, "" if ok else error_for(state["lines"]))
            state["label"] = None
            return
        if s:
            state["lines"].append(s)

    rc = run_stream(["/bin/bash", str(HERE / "reload_ledger.sh")], view, logf, on_line)
    if rc and not view.failures:       # died before any step (e.g. the drive check)
        view.failures.append(("startup", error_for(list(view.buffer))))
        view._w(f"  {CROSS} {RED}{view.failures[-1][1]}{RESET}")
    return rc


def crash_report(view: View, mode: str, rc: int) -> Path | None:
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        path = LOG_DIR / f"crash-{ts}.log"
        path.write_text("\n".join([
            "PROFICIENT LEDGER SYNC - CRASH REPORT",
            f"generated:  {datetime.datetime.now().isoformat(timespec='seconds')}",
            f"step:       {mode}",
            f"exit code:  {rc}",
            f"python:     {sys.version.split()[0]}",
            f"platform:   {platform.platform()}",
            "", "----- failed -----", *[f"{a}: {b}" for a, b in view.failures],
            "", "----- last output lines -----", *list(view.buffer), "",
        ]), encoding="utf-8")
        return path
    except Exception:
        return None


def main() -> int:
    mode = sys.argv[1] if len(sys.argv) > 1 else ""
    if mode not in ("mirror", "reload"):
        print("usage: sync_view.py mirror [--reconcile] | reload", file=sys.stderr)
        return 2
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    started = datetime.datetime.now()
    with open(LOG_DIR / "run.log", "a", encoding="utf-8") as logf:
        logf.write(f"\n{'=' * 72}\n  {mode.upper()} @ {started:%Y-%m-%d %H:%M:%S}\n{'=' * 72}\n")
        view = View("QBO mirror · refresh" if mode == "mirror" else "Project Ledger · reload")
        view.header()
        rc = mirror(view, logf, "--reconcile" in sys.argv) if mode == "mirror" else reload(view, logf)
        logf.write(f"\n  EXIT {rc} @ {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n")
    view.close(f"{(datetime.datetime.now() - started).total_seconds():.1f}s")
    if rc or view.failures:
        report = crash_report(view, mode, rc)
        if report:
            view._w(f"  {DIM}→ trace: {report}{RESET}\n")
    return rc or (1 if view.failures else 0)


if __name__ == "__main__":
    sys.exit(main())
