"""
sync_view.py — a live visual front-end for the bill tracker sync (sync-ap).

Runs `excel_bill_sync.py` as a subprocess and re-renders its plain stdout as
emoji phases with check-marks, bright counts, color-coded events, and a final
summary panel — the same look as the invoice sync's viewer. It does NOT modify
the sync; it only reads stdout. The raw stream is still appended to the log
file (so file logging is preserved), and a crash report is written on failure.

Usage (args pass straight through to excel_bill_sync.py):
    python3 sync_view.py
    python3 sync_view.py --dry-run

No third-party dependencies. Colors auto-disable when output isn't a terminal.
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
RUN_TARGET = HERE / "excel_bill_sync.py"
LOG_DIR = Path(os.environ.get("LOG_DIR")
               or (Path.home() / "Library/Logs/Proficient/bill-tracker"))

# ── colors (disabled when not a tty or NO_COLOR set) ──────────────────────────
_TTY = sys.stdout.isatty() and os.environ.get("NO_COLOR") is None
def _c(code: str) -> str:
    return code if _TTY else ""
RESET = _c("\033[0m"); BOLD = _c("\033[1m"); DIM = _c("\033[90m")
RED = _c("\033[31m"); GREEN = _c("\033[32m"); YELLOW = _c("\033[33m")
CYAN = _c("\033[36m"); BLUE = _c("\033[94m"); BR_CYAN = _c("\033[96m")
CHECK = f"{GREEN}✓{RESET}"

# Phase keyword → (emoji, label). First keyword found in the "→ …" line wins.
PHASES = [
    ("authenticating",          "🔐", "Auth to QBO"),
    ("fetching vendors",        "👥", "Vendors"),
    ("account + item maps",     "🗂",  "Account + item maps"),
    ("PO map",                  "📦", "Purchase orders"),
    ("payment date map",        "💳", "Payment dates"),
    ("fetching OPEN bills",     "📥", "Open bills"),
    ("fetching PAID bills",     "💵", "Paid bills"),
    ("fetching invoices",       "🧾", "Invoices"),
    ("building rows",           "🧱", "Build rows"),
    ("reading existing",        "♻️",  "Preserve lien / notes"),
    ("rotating backup",         "💾", "Backup"),
    ("collapsing rows",         "🧮", "Collapse rows"),
    ("building workbook",       "📊", "Build workbook"),
    ("validating xlsx",         "🔍", "Validate (Excel-strict)"),
]
# Phases we don't surface as steps (verbose / redundant for a clean view).
SKIP_PHASES = ("status breakdown", "by division", "--limit", "excel bill sync starting")


def _phase_for(text: str):
    low = text.lower()
    if any(k in low for k in SKIP_PHASES):
        return None
    for key, icon, label in PHASES:
        if key.lower() in low:
            return icon, label
    return "•", text.rstrip(" …")


class View:
    def __init__(self, out=sys.stdout):
        self.out = out
        self.pending = None            # (icon, label) awaiting its count line
        self.buffer = deque(maxlen=200)
        self.current_phase = "startup"
        self.failed = False
        self.fail_phase = None
        self.fail_headline = ""

    def _w(self, s: str = ""):
        self.out.write(s + "\n"); self.out.flush()

    def header(self):
        self._w(f"\n{BOLD}{BLUE}╭─ Proficient Bill Tracker · sync-ap ──────────────╮{RESET}")

    def note_raw(self, raw: str):
        self.buffer.append(raw.rstrip("\n"))

    def mark_failure(self, headline: str):
        if not self.failed:
            self.failed = True
            self.fail_phase = self.current_phase
            self.fail_headline = headline.strip()[:300]

    def begin_phase(self, text: str):
        self.flush_pending()
        p = _phase_for(text)
        self.pending = p
        if p:
            self.current_phase = p[1]

    def complete_pending(self, detail: str = ""):
        if not self.pending:
            return False
        icon, label = self.pending
        det = f"   {BOLD}{BR_CYAN}{detail[:64]}{RESET}" if detail else ""
        self._w(f"  {icon}  {label}{det}  {CHECK}".rstrip())
        self.pending = None
        return True

    def flush_pending(self):
        if self.pending:
            icon, label = self.pending
            self._w(f"  {icon}  {label}  {CHECK}")
            self.pending = None

    def step(self, icon: str, label: str, detail: str = ""):
        self.flush_pending()
        det = f"   {BOLD}{BR_CYAN}{detail}{RESET}" if detail else ""
        self._w(f"  {icon}  {label}{det}  {CHECK}".rstrip())

    def event(self, text: str, color: str = DIM, icon: str = "⤷"):
        self.flush_pending()
        self._w(f"     {color}{icon} {text}{RESET}")

    def warn(self, text: str):
        self.flush_pending()
        self._w(f"     {YELLOW}⚠ {text}{RESET}")

    def summary(self, text: str):
        self.flush_pending()
        self._w(f"  {BOLD}{CYAN}{text}{RESET}")

    def close(self, elapsed: str = ""):
        self.flush_pending()
        tail = f"  {DIM}{elapsed}{RESET}" if elapsed else ""
        self._w(f"{BOLD}{BLUE}╰──────────────────────────────────────────────────╯{RESET}{tail}\n")

    def failure_panel(self, exit_code, report_path):
        hard = exit_code not in (0, None)
        tag = f"{BOLD}{RED}✗ FAILED{RESET}" if hard else f"{BOLD}{YELLOW}⚠ COMPLETED WITH ISSUES{RESET}"
        self._w(f"  {tag} during {BOLD}{self.fail_phase or 'run'}{RESET}")
        if self.fail_headline:
            self._w(f"  {RED}{self.fail_headline}{RESET}")
        if report_path:
            self._w(f"  {DIM}→ trace: {report_path}{RESET}\n")


def process_stream(lines, view: View, logf=None):
    """Consume raw stdout lines, mirror them to the log, render them pretty."""
    for raw in lines:
        if logf:
            logf.write(raw); logf.flush()
        view.note_raw(raw)
        line = raw.rstrip("\n")
        s = line.strip()
        if not s:
            continue

        # ── failure detection ──
        if "Traceback (most recent call last)" in line:
            view.mark_failure("Unhandled exception — see crash report")
        elif s.startswith("✗"):
            view.mark_failure(s.lstrip("✗ ").strip())

        # ── phase starts ("→ …") ──
        if s.startswith("→"):
            view.begin_phase(s[1:].strip())
            continue

        # ── terminal events / summaries ──
        if "✓ saved" in s:
            view.step("✅", "Saved workbook")
            continue
        if s.startswith("✓ done") or s.startswith("✓ dry run"):
            continue  # handled by close() elapsed
        if s.startswith("Bills:") or s.startswith("MFD:"):
            view.summary(s)
            continue
        if s.startswith("audit:"):
            view.event(s, color=DIM, icon="🔎")
            continue
        if s.lower().startswith(("output:", "open:")):
            view.event(s, color=DIM, icon="📂")
            continue
        if s.startswith("⚠") or s.startswith("↻") or "failed" in s.lower():
            view.warn(s.lstrip("⚠↻ ").strip())
            continue

        # ── count / detail line under a pending phase ──
        if view.pending and (line.startswith("  ") or s.startswith("ok.")):
            view.complete_pending(s.replace("ok.", "").strip() or "ready")
            continue

        # ── 4-space breakdown rows ("status: n") — suppressed for clean view ──
        if line.startswith("    "):
            continue

        # anything else: dim passthrough
        view.event(s, color=DIM, icon="·")


def write_crash_report(view: View, command: str, exit_code, log_dir: Path):
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        path = log_dir / f"crash-{ts}.log"
        path.write_text("\n".join([
            "PROFICIENT BILL TRACKER — CRASH REPORT",
            f"generated:  {datetime.datetime.now().isoformat(timespec='seconds')}",
            f"command:    {command}",
            f"exit code:  {exit_code}",
            f"phase:      {view.fail_phase or 'unknown'}",
            f"python:     {sys.version.split()[0]}",
            f"platform:   {platform.platform()}",
            "", "----- failure headline -----",
            view.fail_headline or "(none captured — see tail below)",
            "", "----- last output lines -----", *list(view.buffer), "",
        ]), encoding="utf-8")
        return path
    except Exception:
        return None


def main() -> int:
    if not RUN_TARGET.exists():
        print(f"Cannot find {RUN_TARGET}", file=sys.stderr)
        return 2

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logf = open(LOG_DIR / "run.log", "a", encoding="utf-8")
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logf.write(f"\n{'='*72}\n  RUN @ {ts}\n{'='*72}\n")

    view = View()
    view.header()
    started = datetime.datetime.now()
    proc = subprocess.Popen(
        [sys.executable, "-u", str(RUN_TARGET), *sys.argv[1:]],
        cwd=str(HERE), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, bufsize=1,
    )
    try:
        process_stream(iter(proc.stdout.readline, ""), view, logf)
    except KeyboardInterrupt:
        proc.terminate()
        view.mark_failure("Interrupted by user (Ctrl-C)")
        view._w(f"\n  {RED}interrupted{RESET}")
    finally:
        proc.wait()
        elapsed = f"{(datetime.datetime.now() - started).total_seconds():.1f}s"
        rc = proc.returncode or 0
        logf.write(f"\n  EXIT {rc} @ {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n")
        logf.close()
        view.close(elapsed=f"done in {elapsed}")

    rc = proc.returncode or 0
    if view.failed or rc != 0:
        report = write_crash_report(view, "python3 excel_bill_sync.py " + " ".join(sys.argv[1:]), rc, LOG_DIR)
        view.failure_panel(rc, report)
    return rc


if __name__ == "__main__":
    sys.exit(main())
