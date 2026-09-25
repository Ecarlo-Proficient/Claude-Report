"""
sync_view.py — a live visual front-end for the invoice sync.

Runs `run_invoice_sync.py` as a subprocess and re-renders its log stream as
phases with check-marks, a live progress bar over the invoice loop, color-coded
events, and a final summary panel. It does NOT modify the sync — it only reads
the sync's stdout, so the sync's own file logging (sync.log) is untouched.

Usage (pass any run_invoice_sync.py args straight through):
    python3 sync_view.py
    python3 sync_view.py --dry-run

No third-party dependencies. Colors auto-disable when output isn't a terminal
(piped / redirected), in which case it falls back to plain pass-through.
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

# Where crash reports land (matches the sync's own log dir; LOG_DIR overrides).
LOG_DIR = Path(os.environ.get("LOG_DIR")
               or (Path.home() / "Library/Logs/Proficient/automation-worker"))

HERE = Path(__file__).resolve().parent
RUN_TARGET = HERE / "run_invoice_sync.py"

# ── colors (disabled when not a tty or NO_COLOR set) ──────────────────────────
_TTY = sys.stdout.isatty() and os.environ.get("NO_COLOR") is None
def _c(code: str) -> str:
    return code if _TTY else ""
RESET = _c("\033[0m"); BOLD = _c("\033[1m"); DIM = _c("\033[90m")
RED = _c("\033[31m"); GREEN = _c("\033[32m"); YELLOW = _c("\033[33m")
CYAN = _c("\033[36m"); BLUE = _c("\033[94m"); MAGENTA = _c("\033[35m")
BR_CYAN = _c("\033[96m")   # bright cyan — for counts that must stand out
CLEAR_EOL = "\033[K" if _TTY else ""

CHECK = f"{GREEN}✓{RESET}"

# ── log line shape:  "<ts> <LEVEL> <logger> — <message>" ──────────────────────
_LINE = re.compile(r"^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d\s+(\w+)\s+\S+\s+—\s+(.*)$")


def _split(line: str):
    """Return (level, message) for a logger line, or (None, raw) otherwise."""
    m = _LINE.match(line.rstrip("\n"))
    if m:
        return m.group(1), m.group(2)
    return None, line.rstrip("\n")


class View:
    """Holds render state (the live progress bar) and prints formatted lines."""

    def __init__(self, out=sys.stdout):
        self.out = out
        self.bar_total = 0
        self.bar_cur = 0
        self.bar_active = False
        self.fuzzy = 0
        self.archived_nums: list[str] = []
        # ── failure tracing ──
        self.buffer = deque(maxlen=200)   # rolling tail of raw output
        self.current_phase = "startup"    # updated as phases run
        self.failed = False               # saw an ERROR / traceback
        self.fail_phase = None            # phase active when first error hit
        self.fail_headline = ""           # first error line, for the panel

    def _w(self, s: str = "", end: str = "\n"):
        self.out.write(s + end)
        self.out.flush()

    def header(self):
        self._w(f"\n{BOLD}{BLUE}╭─ Proficient Invoice Sync ───────────────────────╮{RESET}")

    # ── failure tracing helpers ──
    def note_raw(self, raw: str):
        self.buffer.append(raw.rstrip("\n"))

    def mark_failure(self, headline: str):
        if not self.failed:
            self.failed = True
            self.fail_phase = self.current_phase
            self.fail_headline = headline.strip()[:300]

    def step(self, icon: str, label: str, detail: str = "", done: bool = False):
        self.current_phase = label
        self._end_bar()
        mark = CHECK if done else f"{CYAN}·{RESET}"
        # Counts are highlighted (bold bright-cyan) so they're easy to read,
        # not dim grey. Two spaces after the icon guarantee a visible gap
        # regardless of how wide a given emoji renders.
        det = f"   {BOLD}{BR_CYAN}{detail}{RESET}" if detail else ""
        self._w(f"  {icon}  {label}{det}  {mark}".rstrip())

    def event(self, text: str, color: str = DIM, icon: str = "⤷"):
        self._end_bar()
        self._w(f"     {color}{icon} {text}{RESET}")

    def warn(self, text: str):
        self._end_bar()
        self._w(f"     {RED}⚠ {text}{RESET}")

    # ── progress bar over the invoice upsert loop ──
    def bar(self, cur: int, total: int, width: int = 26):
        self.current_phase = "Upsert invoices"
        self.bar_total = total
        self.bar_cur = cur
        self.bar_active = True
        frac = (cur / total) if total else 0.0
        filled = int(width * frac)
        blocks = f"{GREEN}{'█' * filled}{DIM}{'░' * (width - filled)}{RESET}"
        if _TTY:
            self.out.write(f"\r{CLEAR_EOL}  🔄 Upsert  [{blocks}] {cur}/{total}")
            self.out.flush()
        else:
            self._w(f"  Upsert {cur}/{total}")

    def _end_bar(self):
        if self.bar_active and _TTY:
            # ✓ only when the loop actually finished; otherwise just break the
            # line (e.g. a warning fired mid-loop) so we don't fake completion.
            done = self.bar_total and self.bar_cur >= self.bar_total
            self.out.write(("  " + CHECK + "\n") if done else "\n")
            self.out.flush()
        self.bar_active = False

    def footer_summary(self, label: str, d: dict):
        self._end_bar()
        upd = d.get("updated", 0); new = d.get("created", 0)
        deld = d.get("archived_deleted", 0); old = d.get("archived_paid_old", 0)
        flip = d.get("flipped_to_paid", 0); err = d.get("errors", 0)
        unm = d.get("customer_unmatched", 0)
        errc = RED if err else GREEN
        self._w(
            f"  {BOLD}{label:<8}{RESET} "
            f"{upd} upd · {new} new · {flip} paid · "
            f"{YELLOW}{deld} del{RESET} · {old} aged · "
            f"{errc}{err} err{RESET}"
            + (f" · {DIM}{unm} unmatched{RESET}" if unm else "")
        )

    def close(self):
        self._end_bar()
        if self.fuzzy:
            self._w(f"  {DIM}({self.fuzzy} fuzzy customer matches){RESET}")
        self._w(f"{BOLD}{BLUE}╰─────────────────────────────────────────────────╯{RESET}\n")

    def failure_panel(self, exit_code: int, report_path):
        hard = exit_code not in (0, None)
        if hard:
            tag = f"{BOLD}{RED}✗ FAILED{RESET}"
        else:
            tag = f"{BOLD}{YELLOW}⚠ COMPLETED WITH ERRORS{RESET}"
        self._w(f"  {tag} during {BOLD}{self.fail_phase or 'run'}{RESET}")
        if self.fail_headline:
            self._w(f"  {RED}{self.fail_headline}{RESET}")
        if report_path:
            self._w(f"  {DIM}→ trace: {report_path}{RESET}\n")


def write_crash_report(view: "View", command: str, exit_code: int, log_dir: Path):
    """Write a self-contained failure report and return its path (or None)."""
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        path = log_dir / f"crash-{ts}.log"
        body = [
            "PROFICIENT INVOICE SYNC — CRASH REPORT",
            f"generated:  {datetime.datetime.now().isoformat(timespec='seconds')}",
            f"command:    {command}",
            f"exit code:  {exit_code}",
            f"phase:      {view.fail_phase or 'unknown'}",
            f"python:     {sys.version.split()[0]}",
            f"platform:   {platform.platform()}",
            "",
            "----- failure headline -----",
            view.fail_headline or "(none captured — see output tail below)",
            "",
            "----- last output lines -----",
            *list(view.buffer),
            "",
        ]
        path.write_text("\n".join(body), encoding="utf-8")
        return path
    except Exception:
        return None


# ── marker → action dispatch ──────────────────────────────────────────────────
def process_stream(lines, view: View):
    """Consume an iterable of raw stdout lines and render them. Testable."""
    summary_re = re.compile(r"(Res/Com|MFD) summary:\s*(\{.*\})")
    for raw in lines:
        view.note_raw(raw)
        level, msg = _split(raw)

        # ---- failure detection (capture phase + headline) ----
        if level in ("ERROR", "CRITICAL"):
            view.mark_failure(msg)
        elif level is None and "Traceback (most recent call last)" in raw:
            view.mark_failure("Unhandled exception — see traceback in crash report")

        # ---- the invoice loop progress bar ----
        m = re.search(r"Progress:\s*(\d+)/(\d+)", msg)
        if m:
            view.bar(int(m.group(1)), int(m.group(2)))
            continue

        # ---- phase results (✓ with a count) ----
        for pat, icon, label in (
            (r"Customer hierarchy loaded:\s*(\d[\d,]*)", "👥", "Customer hierarchy"),
            (r"Term map loaded:\s*(\d[\d,]*)", "💲", "Term map"),
            (r"Payment dates loaded:\s*(\d[\d,]*)", "💳", "Payment dates"),
            (r"QBO returned\s*(\d[\d,]*)\s*open invoices", "📥", "Open invoices"),
            (r"Exported\s*(\d[\d,]*)\s*invoices", "📊", "Excel export"),
        ):
            mm = re.search(pat, msg)
            if mm:
                view.step(icon, label, detail=mm.group(1), done=True)
                break
        else:
            # ---- phase starts (no count) ----
            if "Authenticating to QBO" in msg:
                view.step("🔐", "Auth to QBO", done=True)
            elif "Loading existing Notion invoice pages" in msg:
                view.step("🗂", "Notion caches loaded", done=True)
            elif msg.startswith("Sweep:"):
                view.step("🧹", "Flip / delete sweep", done=True)
            elif "CDC: checking" in msg:
                view.step("🗑", "CDC deletion check", done=True)
            elif msg.startswith("Cleanup:"):
                view.step("🧽", "12-month archive", done=True)
            # ---- CDC summary + archive events ----
            elif re.search(r"CDC:\s*\d+ deleted in QBO", msg):
                mm = re.search(r"—\s*(\d+) matched.*?(\d+) not tracked", msg)
                if mm:
                    view.event(f"CDC: {mm.group(1)} archived · {mm.group(2)} not in Notion",
                               color=YELLOW, icon="🗑")
            elif "CDC-ARCHIVE invoice #" in msg:
                mm = re.search(r"CDC-ARCHIVE invoice #(\S+)", msg)
                if mm:
                    view.archived_nums.append(mm.group(1))
                    view.event(f"archived #{mm.group(1)} (deleted in QBO)",
                               color=YELLOW, icon="🗑")
            elif "ARCHIVE-DELETED" in msg:
                view.event(msg.replace("[dry-run] ", ""), color=YELLOW, icon="🗑")
            elif "Fuzzy-matched" in msg or "Near-matched" in msg:
                view.fuzzy += 1
            elif "Excel file is OPEN" in msg:
                view.warn("Excel file is open — export skipped (close it and re-run)")
            elif (s := summary_re.search(msg)):
                try:
                    import ast
                    view.footer_summary(s.group(1), ast.literal_eval(s.group(2)))
                except Exception:
                    view.event(msg, color=DIM)
            elif level in ("WARNING", "ERROR", "CRITICAL"):
                view.warn(msg)
            elif level is None and msg.strip():
                # non-logger line (banners, stray output) — show dim
                view.event(msg, color=DIM, icon="·")
            # otherwise: routine INFO chatter — suppressed for a clean view


def main() -> int:
    if not RUN_TARGET.exists():
        print(f"Cannot find {RUN_TARGET}", file=sys.stderr)
        return 2
    view = View()
    view.header()
    proc = subprocess.Popen(
        [sys.executable, "-u", str(RUN_TARGET), *sys.argv[1:]],
        cwd=str(HERE),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )
    try:
        process_stream(iter(proc.stdout.readline, ""), view)
    except KeyboardInterrupt:
        proc.terminate()
        view.mark_failure("Interrupted by user (Ctrl-C)")
        view._w(f"\n  {RED}interrupted{RESET}")
    finally:
        proc.wait()
        view.close()

    rc = proc.returncode or 0
    if view.failed or rc != 0:
        command = "python3 run_invoice_sync.py " + " ".join(sys.argv[1:])
        report = write_crash_report(view, command, rc, LOG_DIR)
        view.failure_panel(rc, report)
    return rc


if __name__ == "__main__":
    sys.exit(main())
