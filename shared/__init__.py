"""shared/ — the ONLY importable common code in this repo.

Rule: tools never import tools. The moment a second tool needs a file,
that file moves here. Tool folders may import from shared/, never from
each other.

Entry scripts (one folder deep) bootstrap with:

    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from shared import qbo_vault as kc
    from shared import paths
"""
