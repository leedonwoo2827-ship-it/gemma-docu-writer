from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from backend.services.composer import compose_with_template_headings
from backend.services.section_composer import compose_section

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from hwp_mcp.pptx_vision.tools.template_inject import (
    list_pptx_slides,
    inject_pptx_from_map,
)
from hwp_mcp.hwpx_vision.lib.md_sections import (
    parse_md_sections,
    match_to_template_headings,
)


router = APIRouter(prefix="/api/pptx", tags=["pptx"])


def _load_sources(md_paths: list[str]) -> list[tuple[str, str]]:
    sources: list[tuple[str, str]] = []
    for p in md_paths:
        pp = Path(p)
        if not pp.exists():
            raise HTTPException(404, f"MD 없음: {p}")
        sources.append((pp.stem, pp.read_text(encoding="utf-8")))
    if not sources:
        raise HTTPException(400, "최소 1개 MD 필요")
    return sources


class SlidesBody(BaseModel):
    template_pptx: str


@router.post("/template/headings")
def template_headings(body: SlidesBody) -> dict[str, Any]:
    if not Path(body.template_pptx).exists():
        raise HTTPException(404, "템플릿 PPTX 없음")
    try:
        slides = list_pptx_slides(body.template_pptx)
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(500, f"슬라이드 추출 실패: {type(e).__name__}: {e}")
    # HWPX 와 같은 인터페이스로 headings 형태 통일
    headings = [
        {
            "heading": s["title"] or f"슬라이드 {s['index'] + 1}",
            "level": 1,
            "body_paragraphs": s["body_shapes"] + sum(rc[0] * rc[1] for rc in s.get("table_rows_cols", [])),
            "_slide_index": s["index"],
            "_has_table": s["has_table"],
        }
        for s in slides
    ]
    return {"headings": headings, "slides_raw": slides}


class DraftMdBody(BaseModel):
    template_pptx: str
    output_md: str
    source_md_paths: list[str]


@router.post("/template/draft-md")
async def template_draft_md(body: DraftMdBody):
    if not Path(body.template_pptx).exists():
        raise HTTPException(404, "템플릿 없음")
    sources = _load_sources(body.source_md_paths)

    try:
        slides = list_pptx_slides(body.template_pptx)
    except Exception as e:
        raise HTTPException(500, f"슬라이드 추출 실패: {e}")

    seen: set[str] = set()
    filtered: list[dict] = []
    for s in slides:
        title = s["title"] or f"슬라이드 {s['index'] + 1}"
        if title in seen:
            continue
        seen.add(title)
        body_para = s["body_shapes"] + sum(rc[0] * rc[1] for rc in s.get("table_rows_cols", []))
        filtered.append({"heading": title, "level": 1, "body_paragraphs": body_para})

    out = Path(body.output_md)
    out.parent.mkdir(parents=True, exist_ok=True)

    async def stream():
        collected: list[str] = []
        try:
            yield f"event: start\ndata: {len(filtered)}\n\n"
            async for chunk in compose_with_template_headings(filtered, sources):
                collected.append(chunk)
                safe = chunk.replace("\r", "").replace("\n", "\\n")
                yield f"data: {safe}\n\n"
            out.write_text("".join(collected), encoding="utf-8")
            yield f"event: done\ndata: {out}\n\n"
        except Exception as e:
            import traceback
            traceback.print_exc()
            yield f"event: error\ndata: {type(e).__name__}: {e}\n\n"

    return StreamingResponse(stream(), media_type="text/event-stream")


class InjectFromMdBody(BaseModel):
    template_pptx: str
    md_path: str
    output_pptx: str


@router.post("/template/inject-from-md")
def template_inject_from_md(body: InjectFromMdBody) -> dict[str, Any]:
    if not Path(body.template_pptx).exists():
        raise HTTPException(404, "템플릿 없음")
    if not Path(body.md_path).exists():
        raise HTTPException(404, "MD 없음")

    md_text = Path(body.md_path).read_text(encoding="utf-8")
    md_sections = parse_md_sections(md_text)

    try:
        slides = list_pptx_slides(body.template_pptx)
    except Exception as e:
        raise HTTPException(500, f"슬라이드 추출 실패: {e}")

    slide_titles = [s["title"] or f"슬라이드 {s['index'] + 1}" for s in slides]
    slide_to_body = match_to_template_headings(md_sections, slide_titles)

    if not slide_to_body:
        raise HTTPException(
            400,
            f"MD 헤딩({len(md_sections)})과 슬라이드 제목({len(slide_titles)}) 매칭 실패",
        )

    try:
        result = inject_pptx_from_map(body.template_pptx, slide_to_body, body.output_pptx)
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(500, f"주입 실패: {e}")
    return {**result, "md_sections_total": len(md_sections)}


class TemplateInjectBody(BaseModel):
    template_pptx: str
    output_pptx: str
    source_md_paths: list[str]


@router.post("/template/inject")
async def template_inject(body: TemplateInjectBody):
    if not Path(body.template_pptx).exists():
        raise HTTPException(404, "템플릿 없음")
    sources = _load_sources(body.source_md_paths)

    try:
        slides = list_pptx_slides(body.template_pptx)
    except Exception as e:
        raise HTTPException(500, f"슬라이드 추출 실패: {e}")

    seen: set[str] = set()
    targets: list[str] = []
    for s in slides:
        title = s["title"] or f"슬라이드 {s['index'] + 1}"
        if title in seen:
            continue
        seen.add(title)
        targets.append(title)

    async def stream():
        slide_map: dict[str, str] = {}
        try:
            yield f"event: start\ndata: {len(targets)}\n\n"
            for i, title in enumerate(targets, 1):
                yield f"event: section_begin\ndata: {i}/{len(targets)}::{title}\n\n"
                body_text = await compose_section(title, sources)
                slide_map[title] = body_text
                preview = body_text[:80].replace("\n", " ")
                yield f"event: section_done\ndata: {i}/{len(targets)}::{title}::{preview}\n\n"
            yield f"event: injecting\ndata: {body.output_pptx}\n\n"
            result = inject_pptx_from_map(body.template_pptx, slide_map, body.output_pptx)
            yield f"event: done\ndata: {result['path']}|{result['bytes']}|{result['slides_replaced']}\n\n"
        except Exception as e:
            import traceback
            traceback.print_exc()
            yield f"event: error\ndata: {type(e).__name__}: {e}\n\n"

    return StreamingResponse(stream(), media_type="text/event-stream")
