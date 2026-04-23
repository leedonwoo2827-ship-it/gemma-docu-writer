from __future__ import annotations

import zipfile
from pathlib import Path


def unpack(pptx_path: Path, dest: Path) -> Path:
    dest = Path(dest)
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pptx_path) as z:
        z.extractall(dest)
    return dest


def pack(src: Path, out_pptx: Path) -> Path:
    src = Path(src)
    out_pptx = Path(out_pptx)
    out_pptx.parent.mkdir(parents=True, exist_ok=True)
    # [Content_Types].xml must be the first entry for strict PPTX consumers.
    ordered: list[Path] = []
    ct = src / "[Content_Types].xml"
    if ct.exists():
        ordered.append(ct)
    for p in sorted(src.rglob("*")):
        if p.is_file() and p != ct:
            ordered.append(p)
    with zipfile.ZipFile(out_pptx, "w", zipfile.ZIP_DEFLATED) as z:
        for p in ordered:
            z.write(p, p.relative_to(src).as_posix())
    return out_pptx
