"""
vault_graph.py - the org's knowledge graph + the system diagrams, for the Graph tab.

WHAT THIS IS
Two read-only, parsed-on-demand graphs feed the ledger's Graph tab:

  1. THE ORG MAP - every note in the AI Brain_Vault as a node, every [[wikilink]]
     between notes as an edge. This is the same view Obsidian's graph gives, built
     here so it lives inside the ledger. Parsed live from the vault markdown on each
     request (like registry_view); nothing is cached to the ledger DB, nothing is
     written back, so the vault stays the single owner of its own truth.

  2. THE SYSTEM DIAGRAMS - the mermaid flowcharts already authored in
     docs/ARCHITECTURE.md (AR / AP / WIP / Ledger / exports / money-bleeds). We do
     NOT redraw them; we import the exact nodes and arrows from the mermaid source
     so the ledger can render them in its own canvas viewer.

NAMES POLICY (binding)
  The vault is role-handle-only by policy, but the one file that maps handles to
  real people - 01_company/ROSTER.md - is gitignored and off-limits. This walker
  EXCLUDES it explicitly (plus dotfolders and Office temp files) so no name can
  ever reach the graph. If a new sensitive file is added to the vault's .gitignore,
  add it to _SKIP_NAMES here too.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared import paths  # noqa: E402

# Files that must never become nodes. ROSTER.md is the name<->handle map (gitignored,
# names live there and only there). The rest are editor/OS cruft, never notes.
_SKIP_NAMES = {"roster.md"}
_SKIP_DIRS = {".obsidian", ".git", ".trash"}

_WIKILINK = re.compile(r"!?\[\[([^\]]+)\]\]")   # [[target#head|alias]] and ![[embed]]
_H1 = re.compile(r"^#\s+(.+?)\s*$")


# ── the org map (vault notes + wikilinks) ───────────────────────────────────
def _iter_notes(base: Path):
    """Yield every .md note under the vault, skipping the sensitive/cruft set."""
    for path in base.rglob("*.md"):
        if any(part in _SKIP_DIRS or part.startswith("~$") for part in path.parts):
            continue
        if path.name.lower() in _SKIP_NAMES:
            continue
        yield path


def _title_of(path: Path) -> str:
    """A note's display label: its first H1, else the file stem. Never a name -
    the vault is role-handle-only and ROSTER is already excluded."""
    try:
        with path.open(encoding="utf-8") as fh:
            for _ in range(40):
                line = fh.readline()
                if not line:
                    break
                m = _H1.match(line.strip())
                if m:
                    return re.sub(r"[\[\]`*]", "", m.group(1)).strip() or path.stem
    except OSError:
        pass
    return path.stem


def _link_target(raw: str) -> str:
    """A [[wikilink]] body -> the bare target, dropping #heading and |alias."""
    t = raw.split("|", 1)[0].split("#", 1)[0].strip()
    t = t.replace("\\", "/")
    if t.lower().endswith(".md"):
        t = t[:-3]
    return t


def load_org_graph(root: Path | None = None) -> dict:
    """The vault as a node-link graph. Always returns a dict; missing vault -> ok=False."""
    base = Path(root) if root else paths.vault_dir()
    if not base.is_dir():
        return {"ok": False, "source": str(base), "nodes": [], "links": [],
                "error": f"vault not found at {base} (set ACB_VAULT_DIR in machine.env)"}

    notes = sorted(_iter_notes(base))
    # Two indexes for resolving a link target to a note id: by full path-id
    # ("02_processes/roles") and by bare stem ("roles"). Path-qualified links win;
    # a bare stem falls back to the stem index (Obsidian shortlink behaviour).
    by_id: dict[str, str] = {}
    by_stem: dict[str, list[str]] = {}
    meta: dict[str, dict] = {}
    for path in notes:
        nid = path.relative_to(base).with_suffix("").as_posix()
        parts = nid.split("/")
        group = parts[0] if len(parts) > 1 else "hub"
        by_id[nid.lower()] = nid
        by_stem.setdefault(path.stem.lower(), []).append(nid)
        meta[nid] = {"id": nid, "label": _title_of(path), "group": group}

    def resolve(target: str) -> str | None:
        low = target.lower()
        if low in by_id:                                     # path-qualified link wins
            return by_id[low]
        hits = by_stem.get(low.split("/")[-1])               # else fall back to the bare stem
        return hits[0] if hits else None                     # ambiguous stem -> first, deterministically

    edges: set[tuple[str, str]] = set()
    unresolved = 0
    for path in notes:
        nid = path.relative_to(base).with_suffix("").as_posix()
        try:
            text = path.read_text(encoding="utf-8")
        except OSError:
            continue
        for raw in _WIKILINK.findall(text):
            tgt = resolve(_link_target(raw))
            if tgt is None:
                unresolved += 1
                continue
            if tgt == nid:
                continue
            edges.add((nid, tgt) if nid < tgt else (tgt, nid))

    deg: dict[str, int] = {nid: 0 for nid in meta}
    for a, b in edges:
        deg[a] += 1
        deg[b] += 1
    nodes = [dict(m, deg=deg[nid]) for nid, m in meta.items()]
    nodes.sort(key=lambda n: (-n["deg"], n["id"]))
    links = [{"source": a, "target": b} for a, b in sorted(edges)]

    groups: dict[str, int] = {}
    for n in nodes:
        groups[n["group"]] = groups.get(n["group"], 0) + 1
    return {"ok": True, "source": str(base), "nodes": nodes, "links": links,
            "groups": groups, "unresolved": unresolved}


# ── the system diagrams (mermaid in docs/ARCHITECTURE.md) ────────────────────
# Node shapes, longest-closer first so the alternation is greedy-correct.
_NODE_DEF = re.compile(
    r"([A-Za-z0-9_][\w-]*)"
    r"(\[\(.*?\)\]|\(\[.*?\]\)|\[\[.*?\]\]|\(\(.*?\)\)|\{\{.*?\}\}|"
    r"\[.*?\]|\(.*?\)|\{.*?\}|>.*?\])")
_SHAPE_TRIM = ("([", ")]", "[(", ")]", "[[", "]]", "((", "))", "{{", "}}",
               "[", "]", "(", ")", "{", "}", ">")
# Any mermaid connector, arrowhead included: -->  ---  --x  -.->  ==>  ~~~  etc.
# The trailing [>xo] MUST be consumed or the arrow's target id is lost.
_LINK = re.compile(r"\s*(?:-\.{1,}->?|-\.{1,}-|[ox<]?(?:-{2,}|={2,})[>xo]?|~{2,})\s*")
_CLASS_TAG = re.compile(r":::[\w-]+")
_SENT = "\x00LINK\x00"   # edge-label placeholder; no tilde so it can't collide with ~~~


def _clean_label(raw: str) -> str:
    """A mermaid node label -> plain display text."""
    t = raw.strip()
    for a in ("([", ")]", "[(", "[[", "]]", "((", "))", "{{", "}}"):
        if t.startswith(a):
            t = t[len(a):]
        if t.endswith(a):
            t = t[:-len(a)]
    t = t.strip("[](){}>\"' ")
    t = t.replace("\\n", " ").replace("<br/>", " ").replace("<br>", " ")
    t = t.replace("#quot;", '"').replace("&quot;", '"')
    t = t.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
    t = re.sub(r"[`*]", "", t)
    return re.sub(r"\s+", " ", t).strip()


def _parse_mermaid(body: str) -> dict:
    """One mermaid flowchart block -> {direction, nodes[], edges[]}."""
    direction = "TD"
    labels: dict[str, str] = {}
    order: list[str] = []
    cluster_of: dict[str, str] = {}
    edges: list[dict] = []
    seen_edge: set[tuple[str, str]] = set()
    stack: list[str] = []

    def note_node(nid: str, label: str | None = None):
        if nid not in labels:
            labels[nid] = label or nid
            order.append(nid)
            if stack:
                cluster_of[nid] = stack[-1]
        elif label and labels[nid] == nid:
            labels[nid] = label

    for rawline in body.splitlines():
        line = rawline.strip()
        if not line or line.startswith("%%"):
            continue
        low = line.lower()
        m = re.match(r"(?:flowchart|graph)\s+([A-Za-z]{2})\b", line, re.I)
        if m:
            direction = m.group(1).upper()
            continue
        if low.startswith(("classdef", "class ", "style ", "linkstyle")):
            continue
        if low.startswith("subgraph"):
            title = line[len("subgraph"):].strip()
            sm = _NODE_DEF.search(title)
            stack.append(_clean_label(sm.group(2)) if sm else _clean_label(title) or "group")
            continue
        if low == "end":
            if stack:
                stack.pop()
            continue

        # Register every node definition on the line, then blank it to a bare id.
        def _sub(mo):
            nid = mo.group(1)
            note_node(nid, _clean_label(mo.group(2)))
            return nid
        clean = _CLASS_TAG.sub("", _NODE_DEF.sub(_sub, line))

        # Pull edge labels out of the way, replacing the whole labelled connector
        # with a sentinel: |lbl| (after an arrow), -- "lbl" -->, -- lbl -->, == lbl ==>.
        elabels: list[str] = []
        take = lambda mo: (elabels.append(mo.group(1).strip()) or f" {_SENT} ")
        clean = re.sub(r'[-=]{2,}\s*"([^"]*)"\s*[-=]{1,3}[>xo]?', take, clean)
        clean = re.sub(r'[-=]{2,}\s+([^->=|"\n]+?)\s+[-=]{1,3}[>xo]', take, clean)
        clean = re.sub(r"[-=]{1,3}[>xo]?\|([^|]*)\|", take, clean)
        parts = re.split(re.escape(_SENT) + "|" + _LINK.pattern, clean)
        ids = [p.strip() for p in parts if p and p.strip() and re.fullmatch(r"[A-Za-z0-9_][\w-]*", p.strip())]
        for a, b in zip(ids, ids[1:]):
            note_node(a)
            note_node(b)
            if (a, b) not in seen_edge and a != b:
                seen_edge.add((a, b))
                lbl = elabels.pop(0) if elabels else ""
                edges.append({"source": a, "target": b, "label": lbl})

    nodes = [{"id": nid, "label": labels[nid], "cluster": cluster_of.get(nid, "")}
             for nid in order]
    return {"direction": direction, "nodes": nodes, "edges": edges}


def load_architecture_diagrams(repo_root: Path | None = None) -> dict:
    """Import the mermaid system maps from docs/ARCHITECTURE.md, in file order."""
    root = Path(repo_root) if repo_root else Path(__file__).resolve().parent.parent
    src = root / "docs" / "ARCHITECTURE.md"
    if not src.is_file():
        return {"ok": False, "source": str(src), "diagrams": [],
                "error": f"architecture map not found at {src}"}

    diagrams: list[dict] = []
    heading, in_block, buf = "", False, []
    used_keys: set[str] = set()
    for line in src.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if in_block:
            if stripped.startswith("```"):
                parsed = _parse_mermaid("\n".join(buf))
                if parsed["nodes"]:
                    key = re.sub(r"[^a-z0-9]+", "-", heading.lower()).strip("-") or f"diagram-{len(diagrams)}"
                    while key in used_keys:
                        key += "x"
                    used_keys.add(key)
                    diagrams.append(dict(parsed, key=key, title=heading or key))
                in_block, buf = False, []
            else:
                buf.append(line)
            continue
        if stripped.startswith("#"):
            heading = stripped.lstrip("#").strip()
        elif stripped.startswith("```mermaid"):
            in_block = True

    return {"ok": True, "source": str(src), "diagrams": diagrams}


def load_all() -> dict:
    """Everything the Graph tab needs in one payload."""
    return {"org": load_org_graph(), "diagrams": load_architecture_diagrams()}


if __name__ == "__main__":
    org = load_org_graph()
    if not org["ok"]:
        print(org["error"])
        raise SystemExit(1)
    print(f"vault  : {org['source']}")
    print(f"org map: {len(org['nodes'])} notes, {len(org['links'])} links, "
          f"{org['unresolved']} unresolved")
    print("groups :", ", ".join(f"{g}={n}" for g, n in sorted(org["groups"].items())))
    top = sorted(org["nodes"], key=lambda n: -n["deg"])[:8]
    print("hubs   :", ", ".join(f"{n['label']}({n['deg']})" for n in top))
    arch = load_architecture_diagrams()
    print(f"\ndiagrams: {arch['source']}")
    if arch["ok"]:
        for d in arch["diagrams"]:
            print(f"  [{d['direction']}] {len(d['nodes']):2d}n {len(d['edges']):2d}e  "
                  f"{d['key']}  ::  {d['title']}")
    else:
        print("  " + arch["error"])
