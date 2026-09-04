# synology/ — file-tree audit

`synology_audit.py` / `synology_tree.py` scan the NAS shares. **Always pass
`--exclude /Volumes/Proinfo/Items/`** (sensitive — see the repo CLAUDE.md).

Outputs are data dumps and are **never tracked** (see `.gitignore`). The
reference map now lives OUTSIDE the repo at
`~/Library/Logs/Proficient/synology/synology_reference_map.md` — moved
2026-07-30 because it carries staff folder names and the full company layout,
which don't belong on GitHub.
