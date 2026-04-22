from __future__ import annotations

from pathlib import Path
from typing import Any

from ..lib.pptx_template import list_slides, inject_to_template

import sys
ROOT = Path(__file__).resolve().parents[3]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from doc_mcp.hwpx_vision.lib.md_clean import clean_markdown


def list_pptx_slides(template_pptx: str) -> list[dict[str, Any]]:
    if not Path(template_pptx).exists():
        raise FileNotFoundError(template_pptx)
    return list_slides(template_pptx)


def inject_pptx_from_map(
    template_pptx: str,
    slide_to_body: dict[str, str],
    output_pptx: str,
) -> dict[str, Any]:
    if not Path(template_pptx).exists():
        raise FileNotFoundError(template_pptx)
    Path(output_pptx).parent.mkdir(parents=True, exist_ok=True)
    cleaned = {title: clean_markdown(body) for title, body in slide_to_body.items()}
    return inject_to_template(template_pptx, cleaned, output_pptx)
