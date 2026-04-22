"""kordoc 서브프로세스 호출. 없으면 파일 확장자별 폴백 (PDF/HWPX/MD)."""
from __future__ import annotations

import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _kordoc_cli() -> str | None:
    for name in ("kordoc", "kordoc.cmd"):
        path = shutil.which(name)
        if path:
            return path
    local = Path(__file__).resolve().parents[2] / "mcp" / "kordoc" / "dist" / "cli.js"
    if local.exists():
        node = shutil.which("node")
        if node:
            return f"{node}:{local}"
    return None


def _hwpx_to_md_fallback(source: str, out_md: str) -> str:
    """kordoc 없이 HWPX를 직접 파싱해 MD로 추출. 헤딩 패턴(1./가./A./(1))을 #/##/###/#### 로 매핑."""
    from hwp_mcp.hwpx_vision.lib.hwpx_template import (
        _is_heading,
        _heading_level,
        _paragraph_text,
        _is_inside_table,
        NS,
    )
    from lxml import etree
    import tempfile

    lines: list[str] = [f"# {Path(source).stem}", ""]
    with tempfile.TemporaryDirectory() as tmp:
        with zipfile.ZipFile(source, "r") as z:
            z.extractall(tmp)
        contents = Path(tmp) / "Contents"
        section_xmls = sorted(contents.glob("section*.xml"))
        for sx in section_xmls:
            tree = etree.parse(str(sx))
            root = tree.getroot()
            for p in root.xpath(".//hp:p", namespaces=NS):
                if _is_inside_table(p):
                    continue
                txt = _paragraph_text(p)
                if not txt:
                    continue
                if _is_heading(txt):
                    lvl = _heading_level(txt)
                    prefix = "#" * min(6, max(2, lvl + 1))
                    lines.append(f"{prefix} {txt}")
                    lines.append("")
                else:
                    lines.append(txt)
            lines.append("")
    Path(out_md).write_text("\n".join(lines), encoding="utf-8")
    return out_md


def convert_to_md(source: str, out_md: str) -> str:
    """HWP/HWPX/PDF/DOCX 등을 MD로 변환. kordoc 있으면 우선, 없으면 확장자별 폴백."""
    cli = _kordoc_cli()
    if cli:
        node_path = shutil.which("node")
        if node_path and cli.startswith(node_path):
            _, script = cli.split(":", 1)
            cmd = [node_path, script, source, "-o", out_md]
        else:
            cmd = [cli, source, "-o", out_md]
        subprocess.run(cmd, check=True, timeout=120)
        return out_md

    src = Path(source)
    ext = src.suffix.lower()

    if ext == ".pdf":
        import fitz
        doc = fitz.open(str(src))
        try:
            lines: list[str] = [f"# {src.stem}\n"]
            for i, page in enumerate(doc):
                lines.append(f"\n## p.{i + 1}\n")
                lines.append(page.get_text("text"))
        finally:
            doc.close()
        Path(out_md).write_text("\n".join(lines), encoding="utf-8")
        return out_md

    if ext == ".hwpx":
        return _hwpx_to_md_fallback(source, out_md)

    if ext == ".md":
        Path(out_md).write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
        return out_md

    raise RuntimeError(
        f"kordoc CLI를 찾을 수 없고 {ext} 폴백이 없습니다. "
        "HWP는 한/글에서 PDF로 저장 후 올려 주세요."
    )
