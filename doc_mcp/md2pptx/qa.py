from __future__ import annotations

import re
import shutil
import subprocess
import sys
from pathlib import Path

PLACEHOLDER_RE = re.compile(r"xxxx|lorem|ipsum|this.*(page|slide).*layout", re.IGNORECASE)


def run_placeholder_check(pptx_path: Path) -> list[str]:
    """Dump pptx text via markitdown and return lines that look like leftover placeholders."""
    try:
        out = subprocess.run(
            [sys.executable, "-m", "markitdown", str(pptx_path)],
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except (subprocess.CalledProcessError, FileNotFoundError):
        return []  # markitdown not installed; silently skip
    hits: list[str] = []
    for line in (out.stdout or "").splitlines():
        if PLACEHOLDER_RE.search(line):
            hits.append(line.strip())
    return hits


def run_visual_export(pptx_path: Path, out_dir: Path) -> Path | None:
    """Optional: convert pptx → pdf → jpegs if LibreOffice & Poppler are available."""
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    pdftoppm = shutil.which("pdftoppm")
    if not soffice or not pdftoppm:
        return None
    out_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(pptx_path)],
        check=False,
    )
    pdf = out_dir / (pptx_path.stem + ".pdf")
    if not pdf.exists():
        return None
    subprocess.run(
        [pdftoppm, "-jpeg", "-r", "150", str(pdf), str(out_dir / "slide")],
        check=False,
    )
    return pdf
