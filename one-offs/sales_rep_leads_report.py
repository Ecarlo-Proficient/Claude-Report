#!/usr/bin/env python3
"""
sales_rep_leads_report.py - one rep's live outreach book, as a shareable page.

WHY
Estimators need to know which builders/GCs an outreach rep already has a thread
with, so nobody cold-calls an account that is mid-conversation. This renders that
book from the ledger (fed by the Notion Customer List) as a single self-contained
HTML file: the accounts in active conversation as detail cards, then a searchable
A-Z index of the whole working set.

ATTRIBUTION
A rep's working set = customer rows whose Notion "Last edited by" is that rep
(the settled convention - there is no manual Owner property). Pass the rep with
--rep exactly as Notion spells it; run --list-reps to see the options.

SAFETY
  * READ-ONLY on the ledger. Writes one HTML file and nothing else.
  * No names are hard-coded here - the rep is a runtime argument.

USAGE
  python3 one-offs/sales_rep_leads_report.py --list-reps
  python3 one-offs/sales_rep_leads_report.py --rep "<Notion display name>" --out ~/path/Report.html
"""
from __future__ import annotations

import argparse
import datetime as dt
import html
import re
import sqlite3
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from shared import paths  # noqa: E402

DEFAULT_DB = paths.get_path(
    "ACB_LEDGER_DB",
    Path.home() / "Library" / "Application Support" / "Proficient" / "ledger.sqlite3",
)

# pipeline stage -> (sort rank, css class). The label shown IS the Notion status,
# verbatim - a relabelled stage is one more thing to translate back to the record.
STAGES = {
    "Interested":  (1, "hot"),
    "Contacted":   (2, "live"),
    "Follow up":   (3, "warm"),
    "Lead":        (4, "cold"),
    "No response": (5, "dead"),
}
IN_PLAY = "Interested"

_TRAIL_DATE = re.compile(r"\s*[-–]?\s*\d{1,2}/\d{1,2}/\d{2,4}\s*$")
_LEAD_DATE = re.compile(r"\s+(?:[A-Z][a-z]{2,8})\s+\d{1,2},?\s+\d{4}\s*$")


def fmt_date(iso: str | None) -> str:
    if not iso:
        return ""
    try:
        return dt.date.fromisoformat(iso[:10]).strftime("%m/%d/%Y")
    except ValueError:
        return iso


def days_since(iso: str | None, today: dt.date) -> int | None:
    if not iso:
        return None
    try:
        return (today - dt.date.fromisoformat(iso[:10])).days
    except ValueError:
        return None


_DANGLE = re.compile(r"\b(?:for|on|at|to|by|of|scheduled?|due)$", re.I)


def clean_note(note: str) -> str:
    """Drop the date the writer trailed onto the line - it gets its own column.

    Unless the date was the object of the sentence ('Meeting Schedule for 8/25'),
    in which case stripping it strands a preposition and the line stops meaning
    anything. Then we keep the note whole."""
    n = _TRAIL_DATE.sub("", note.strip()).strip(" -–|")
    n = _LEAD_DATE.sub("", n).strip(" -–|")
    if not n or _DANGLE.search(n):
        return note.strip().strip(" -–|")
    return n


_LOWER = {"of", "and", "the", "for", "at", "by", "in", "on", "to", "de", "del"}
_FIXED = {"ST.": "St.", "MT.": "Mt.", "INC": "Inc.", "INC.": "Inc.",
          "CO.": "Co.", "LTD": "Ltd.", "LTD.": "Ltd."}
# short words that are words, not initialisms - without these, ALL SEASONS reads "ALL Seasons"
_WORDS = {"ALL", "THE", "JOE", "MAX", "BEN", "HEM", "NEW", "ONE", "TWO", "TOP", "BIG",
          "OAK", "SUN", "KEY", "BAY", "PRO", "RED", "OLD", "SON", "WAY", "ELM", "AIR", "ART"}


def nice(text: str | None) -> str:
    """Title-case a SHOUTED record without flattening acronyms (CR, JDB, LLC, M/I).

    Rules: an all-caps token of 3 chars or fewer is an acronym and survives; a longer
    all-caps token gets title-cased; an already-mixed-case token is left exactly as the
    person typed it; joiner words drop to lowercase unless they lead. Emails pass through
    untouched - 'Arh-Marketing@Arhomes.Com' is not an improvement."""
    if not text:
        return ""
    if "@" in text:
        return text
    out = []
    for i, tok in enumerate(text.split()):
        core = tok.strip(".,()")
        if tok.upper() in _FIXED:
            out.append(_FIXED[tok.upper()])
        elif not tok.isupper():
            out.append(tok)                              # already typed properly
        elif core.lower() in _LOWER:
            out.append(tok.lower() if i else tok.title())
        elif core.upper() in _WORDS or len(core) > 3:
            out.append(tok.title())
        else:
            out.append(tok)                              # acronym - leave it shouting
    return " ".join(out)


def div_tag(division: str | None) -> tuple[str, str]:
    d = (division or "").lower()
    if d.startswith("com"):
        return ("CP", "cp")
    if d.startswith("res"):
        return ("RP", "rp")
    if "multi" in d:
        return ("MFD", "mfd")
    return ("", "")


# statuses that mean somebody has actually made contact - the only ones that can
# collide. An untouched Lead is nobody's account yet, and a closed row is history.
TOUCHED = ("Contacted", "Follow up", "Interested", "No response")


def fetch(db: Path, rep: str | None):
    con = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    cols = ("customer_key, name, division, sales_status, last_contacted, primary_contact, "
            "primary_email, primary_phone, n_touches, notion_url, last_edited_by")
    if rep:
        rows = con.execute(f"SELECT {cols} FROM customer WHERE last_edited_by = ?", (rep,)).fetchall()
    else:
        marks = ",".join("?" * len(TOUCHED))
        rows = con.execute(
            f"SELECT {cols} FROM customer WHERE sales_status IN ({marks})", TOUCHED
        ).fetchall()
    keys = {r["customer_key"] for r in rows}
    touches: dict[str, list[sqlite3.Row]] = {k: [] for k in keys}
    for t in con.execute(
        "SELECT customer_key, seq, touch_date, note FROM sales_touch ORDER BY customer_key, seq"
    ):
        if t["customer_key"] in keys:
            touches[t["customer_key"]].append(t)

    # accounts that are ALREADY a live job's builder/GC - the sharpest conflict flag
    live: dict[str, list[str]] = {}
    projects = con.execute(
        "SELECT project_no, builder_or_gc FROM project WHERE builder_or_gc IS NOT NULL"
    ).fetchall()
    def norm(s: str) -> str:
        return re.sub(r"[^A-Z0-9 ]", "", (s or "").upper()).strip()
    for r in rows:
        stem = norm(r["name"])[:14]
        if len(stem) < 6:
            continue
        hits = sorted({p["project_no"] for p in projects if stem in norm(p["builder_or_gc"])})
        if hits:
            live[r["customer_key"]] = hits
    con.close()
    return rows, touches, live


def _tag(label: str, cls: str) -> str:
    return '<span class="tag %s">%s</span>' % (cls, label) if label else ""


def _log_li(date_txt: str, note: str) -> str:
    return '<li><span class="d">%s</span>%s</li>' % (date_txt or "-", note)


def render(rep: str | None, rows, touches, live, today: dt.date) -> str:
    e = html.escape
    ordered = sorted(
        rows,
        key=lambda r: (
            STAGES.get(r["sales_status"], (9, ""))[0],
            -(days_since(r["last_contacted"], today) is not None),
            r["last_contacted"] or "",
        ),
    )
    in_play = [r for r in ordered if r["sales_status"] == IN_PLAY]
    in_play.sort(key=lambda r: r["last_contacted"] or "", reverse=True)

    # ── cards for the live conversations ──
    cards = []
    for r in in_play:
        label, tag = div_tag(r["division"])
        d = days_since(r["last_contacted"], today)
        age = ("1 day ago" if d == 1 else f"{d} days ago") if d is not None else "no date logged"
        flag = ""
        if r["customer_key"] in live:
            flag = ('<span class="badge live-job">Already a live client &middot; '
                    + e(", ".join(live[r["customer_key"]])) + "</span>")
        log = "".join(
            _log_li(e(fmt_date(t["touch_date"])), e(clean_note(t["note"])))
            for t in touches.get(r["customer_key"], [])
        ) or '<li class="none">No interaction log yet.</li>'
        contact_bits = []
        if r["primary_contact"]:
            contact_bits.append(f'<span class="who">{e(nice(r["primary_contact"]))}</span>')
        if r["primary_email"]:
            contact_bits.append(f'<a href="mailto:{e(r["primary_email"])}">{e(r["primary_email"])}</a>')
        if r["primary_phone"]:
            contact_bits.append(e(r["primary_phone"]))
        name_tag = _tag(label, tag)
        cards.append(f"""
  <div class="acct">
    <div class="top">
      <span class="name">{e(nice(r["name"]))}{name_tag}</span>
      <span class="age">last touch {e(fmt_date(r["last_contacted"])) or "n/a"} <i>({age})</i></span>
    </div>
    <p class="contact">{' &middot; '.join(contact_bits) or '<span class="none">no contact on file</span>'}</p>
    {flag}
    <ul class="log">{log}</ul>
  </div>""")

    # ── the A-Z index of everything ──
    cols_html = (
        '<colgroup><col style="width:32%"><col style="width:5%"><col style="width:14%">'
        '<col style="width:11%"><col style="width:15%"><col style="width:23%"></colgroup>'
        if rep else
        '<colgroup><col style="width:26%"><col style="width:5%"><col style="width:12%">'
        '<col style="width:11%"><col style="width:15%"><col style="width:13%">'
        '<col style="width:18%"></colgroup>'
    )
    worked_th = "" if rep else "<th>Worked by</th>"

    def worked_td(r) -> str:
        return "" if rep else f'<td class="who">{e(r["last_edited_by"] or "-")}</td>'

    trs = []
    for r in sorted(rows, key=lambda r: r["name"].upper()):
        label, tag = div_tag(r["division"])
        _rank, cls = STAGES.get(r["sales_status"], (9, ""))
        stage_label = r["sales_status"] or "-"
        badge = ' <span class="mini">live client</span>' if r["customer_key"] in live else ""
        trs.append(
            f'<tr data-s="{e((r["name"] + " " + (r["primary_contact"] or "") + " " + (r["primary_email"] or "") + " " + (r["last_edited_by"] or "")).lower())}">'
            f'<td class="co">{e(nice(r["name"]))}{badge}</td>'
            f'<td>{_tag(label, tag)}</td>'
            f'<td><span class="pill {cls}">{e(stage_label)}</span></td>'
            f'<td class="num">{e(fmt_date(r["last_contacted"])) or "-"}</td>'
            f'{worked_td(r)}'
            f'<td class="who">{e(nice(r["primary_contact"])) or "-"}</td>'
            f'<td class="em">{e(r["primary_email"] or "-")}</td></tr>'
        )

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(rep) + ' &middot; Outreach Accounts' if rep else 'Accounts Being Worked'}</title>
<style>
 :root{{--ink:#1c2230;--muted:#5b6472;--line:#e4e7ee;--bg:#f6f7f9;--card:#fff;
  --hot:#2f9e44;--live:#3b5bdb;--warm:#e8a13a;--cold:#8a94a6;--dead:#e03131;
  --rp:#2f9e44;--cp:#e8590c;--mfd:#3b5bdb;}}
 *{{box-sizing:border-box}}
 body{{margin:0;background:var(--bg);color:var(--ink);
  font:14.5px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;-webkit-font-smoothing:antialiased;}}
 .wrap{{max-width:960px;margin:0 auto;padding:30px 24px 44px;}}
 header{{border-bottom:2px solid var(--ink);padding-bottom:14px;margin-bottom:16px;}}
 h1{{font-size:23px;margin:0 0 3px;letter-spacing:-.02em;}}
 .sub{{color:var(--muted);font-size:13px;margin:0;}} .sub b{{color:var(--ink);font-weight:600;}}
 h2{{font-size:13px;margin:24px 0 10px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);}}
 .acct{{background:var(--card);border:1px solid var(--line);border-radius:11px;padding:13px 15px;margin-bottom:9px;
  border-left:3px solid var(--hot);break-inside:avoid;}}
 .acct .top{{display:flex;justify-content:space-between;align-items:baseline;gap:12px;}}
 .acct .name{{font-weight:600;font-size:14.5px;}}
 .acct .age{{font-size:11.5px;color:var(--muted);white-space:nowrap;}} .acct .age i{{font-style:normal;}}
 .contact{{margin:4px 0 0;font-size:12.5px;color:var(--muted);}}
 .contact .who{{color:var(--ink);font-weight:600;}}
 .contact a{{color:var(--live);text-decoration:none;}}
 .tag{{font-size:10.5px;font-weight:800;letter-spacing:.04em;margin-left:8px;vertical-align:middle;}}
 .tag.rp{{color:var(--rp);}} .tag.cp{{color:var(--cp);}} .tag.mfd{{color:var(--mfd);}}
 .badge{{display:inline-block;margin-top:7px;font-size:11px;font-weight:700;padding:2px 8px;border-radius:5px;
  background:#fdecec;color:var(--dead);}}
 .log{{margin:7px 0 0;padding-left:15px;font-size:12.5px;color:#333;}}
 .log li{{margin:1.5px 0;}} .log .d{{color:var(--muted);display:inline-block;min-width:78px;}}
 .log .none,.none{{color:var(--muted);font-style:italic;}}
 .tools{{margin:0 0 10px;}}
 #q{{width:100%;padding:9px 12px;font-size:14px;border:1px solid var(--line);border-radius:9px;background:var(--card);
  color:var(--ink);font-family:inherit;}}
 #q:focus{{outline:none;border-color:var(--live);}}
 table{{width:100%;table-layout:fixed;border-collapse:collapse;background:var(--card);border:1px solid var(--line);border-radius:11px;
  overflow:hidden;font-size:12.5px;}}
 th{{text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);
  padding:9px 11px;border-bottom:1px solid var(--line);font-weight:700;background:#fafbfc;}}
 td{{padding:7px 11px;border-bottom:1px solid #f0f2f6;vertical-align:top;}}
 tr:last-child td{{border-bottom:none;}}
 td.co{{font-weight:600;}} td.num{{white-space:nowrap;color:var(--muted);}}
 td .tag{{margin-left:0;}}
 td.who{{color:#333;overflow-wrap:break-word;}} td.em{{color:var(--muted);font-size:11.5px;overflow-wrap:break-word;}}
 td.co,th{{hyphens:none;}}
 .pill{{font-size:10.5px;font-weight:700;padding:2px 8px;border-radius:20px;white-space:nowrap;}}
 .pill.hot{{background:#e6f4ea;color:var(--hot)}} .pill.live{{background:#e7ecfd;color:var(--live)}}
 .pill.warm{{background:#fdf1de;color:#a86a12}} .pill.cold{{background:#eef0f3;color:var(--cold)}}
 .pill.dead{{background:#fdecec;color:var(--dead)}}
 .mini{{font-size:9.5px;font-weight:700;color:var(--dead);background:#fdecec;padding:1px 5px;border-radius:4px;
  margin-left:6px;text-transform:uppercase;letter-spacing:.03em;vertical-align:middle;
  white-space:nowrap;display:inline-block;}}
 .foot{{margin-top:20px;font-size:11.5px;color:var(--muted);border-top:1px solid var(--line);padding-top:10px;}}
 @media print{{body{{background:#fff}} .wrap{{max-width:none;padding:0}} .tools{{display:none}}
  .acct,table{{border-color:#ccc}} h2{{break-after:avoid}} tr{{break-inside:avoid}}
  table{{overflow:visible;border-radius:0}} thead{{display:table-header-group}}}}
</style></head><body><div class="wrap">

<header>
  <h1>{"Outreach Accounts in Play" if rep else "Accounts Being Worked"}</h1>
  <p class="sub"><b>{e(rep) if rep else "All reps"}</b> &middot; {"sales outreach book" if rep else "every account someone has contacted"}
     &middot; as of {today.strftime('%m/%d/%Y')} &middot; {len(rows)} accounts</p>
</header>

<h2>Active now</h2>
{''.join(cards)}

<h2>Full list ({len(rows)})</h2>
<div class="tools"><input id="q" type="search" placeholder="Search company, contact or email..." autocomplete="off"></div>
<table>
 {cols_html}
 <thead><tr><th>Company</th><th>Div</th><th>Stage</th><th>Last contact</th>{worked_th}<th>Contact</th><th>Email</th></tr></thead>
 <tbody id="tb">{''.join(trs)}</tbody>
</table>

<p class="foot">Notion Customer List via the project ledger. &ldquo;Live client&rdquo; = already the builder/GC
on an active job.</p>
</div>
<script>
 var q=document.getElementById('q'),rows=[].slice.call(document.querySelectorAll('#tb tr'));
 q.addEventListener('input',function(){{var v=q.value.trim().toLowerCase();
  rows.forEach(function(r){{r.style.display=!v||r.dataset.s.indexOf(v)>-1?'':'none';}});}});
</script>
</body></html>"""


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--rep", help="Notion display name of the rep (exact)")
    ap.add_argument("--all", action="store_true",
                    help="every rep's touched accounts (the collision list), with a Worked-by column")
    ap.add_argument("--out", type=Path, help="output .html path")
    ap.add_argument("--db", type=Path, default=DEFAULT_DB)
    ap.add_argument("--list-reps", action="store_true", help="show who has a working set, then exit")
    a = ap.parse_args()

    if not a.db.exists():
        print(f"ledger not found: {a.db}", file=sys.stderr)
        return 2

    if a.list_reps:
        con = sqlite3.connect(f"file:{a.db}?mode=ro", uri=True)
        for rep, n in con.execute(
            "SELECT COALESCE(last_edited_by,'(unknown)'), COUNT(*) FROM customer "
            "GROUP BY 1 ORDER BY 2 DESC"
        ):
            print(f"{n:>5}  {rep}")
        con.close()
        return 0

    if not a.out or (not a.rep and not a.all):
        ap.error("--out plus either --rep or --all is required (or use --list-reps)")
    if a.rep and a.all:
        ap.error("--rep and --all are mutually exclusive")

    rows, touches, live = fetch(a.db, a.rep)
    if not rows:
        who = "any rep" if a.all else repr(a.rep)
        print(f"no accounts found for {who} - check --list-reps for the exact spelling", file=sys.stderr)
        return 1

    a.out.parent.mkdir(parents=True, exist_ok=True)
    a.out.write_text(render(a.rep if a.rep else None, rows, touches, live, dt.date.today()), encoding="utf-8")
    print(f"wrote {a.out}  ({len(rows)} accounts, {sum(len(v) for v in touches.values())} logged touches, "
          f"{len(live)} already-live-client flags)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
