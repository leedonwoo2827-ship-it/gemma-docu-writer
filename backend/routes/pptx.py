"""
PPTX 라우터 v5 — md2pptx-template 엔진.

단일 엔드포인트: POST /api/pptx/convert
  - MD + 양식 PPTX → 결과 PPTX (결정론적, LLM/API 없음)
  - 디자인 byte-level 보존
  - 테이블은 헤더 Jaccard 매칭
  - 미매칭 슬라이드는 기본 삭제 (keep_unused 로 유지 가능)
"""
from __future__ import annotations

import sys
import time as _time
from pathlib import Path
from typing import Any

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from doc_mcp.md2pptx.cli import convert as md2pptx_convert


router = APIRouter(prefix="/api/pptx", tags=["pptx"])


class ConvertBody(BaseModel):
    template_pptx: str
    md_path: str
    output_pptx: str | None = None
    dry_run: bool = False
    keep_unused: bool = False


@router.post("/convert")
def pptx_convert(body: ConvertBody) -> dict[str, Any]:
    tpl = Path(body.template_pptx)
    md = Path(body.md_path)
    if not tpl.exists():
        raise HTTPException(404, f"양식 PPTX 없음: {tpl}")
    if not md.exists():
        raise HTTPException(404, f"MD 없음: {md}")
    if tpl.suffix.lower() != ".pptx":
        raise HTTPException(400, f"양식 파일이 PPTX 아님: {tpl.name}")
    if md.suffix.lower() != ".md":
        raise HTTPException(400, f"MD 파일이 .md 아님: {md.name}")

    # 출력 경로 결정
    if body.output_pptx:
        out = Path(body.output_pptx)
    else:
        ts = _time.strftime("%Y%m%d_%H%M%S")
        out = md.parent / f"{md.stem}_result_{ts}.pptx"

    try:
        result = md2pptx_convert(
            template=str(tpl),
            md=str(md),
            out=str(out),
            dry_run=body.dry_run,
            keep_unused=body.keep_unused,
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(500, f"변환 실패: {type(e).__name__}: {e}")

    result_size = 0
    if result.get("output_path") and Path(result["output_path"]).exists():
        result_size = Path(result["output_path"]).stat().st_size

    return {
        "output_path": result.get("output_path", ""),
        "bytes": result_size,
        "slides_count": result.get("slides_count", 0),
        "slides_final": result.get("slides_final", []),
        "slides_dropped": result.get("slides_dropped", []),
        "headings_matched": result.get("headings_matched", []),
        "tables_matched": result.get("tables_matched", []),
        "tables_unmatched": result.get("tables_unmatched", []),
        "plan_text": result.get("plan_text", ""),
        "dry_run": result.get("dry_run", False),
    }
