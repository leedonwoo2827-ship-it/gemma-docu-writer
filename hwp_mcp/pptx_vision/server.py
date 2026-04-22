from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from .tools.template_inject import (
    list_pptx_slides as _list_slides,
    inject_pptx_from_map as _inject,
)


app = FastMCP("pptx_vision")


@app.tool()
def list_slides(template_pptx: str) -> list[dict[str, Any]]:
    """템플릿 PPTX의 슬라이드별 메타 정보(제목, 본문 도형 수, 표 구조)를 반환한다."""
    return _list_slides(template_pptx)


@app.tool()
def inject_to_template(
    template_pptx: str,
    slide_to_body: dict[str, str],
    output_pptx: str,
) -> dict[str, Any]:
    """슬라이드 제목별로 본문 텍스트를 주입해 새 PPTX를 저장한다 (이미지/레이아웃/애니메이션 유지)."""
    return _inject(template_pptx, slide_to_body, output_pptx)


if __name__ == "__main__":
    app.run()
