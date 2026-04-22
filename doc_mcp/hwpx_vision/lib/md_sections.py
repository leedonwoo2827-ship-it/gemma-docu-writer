"""Markdown을 헤딩 단위로 분해해 {heading_text: body_text} 로 반환."""
from __future__ import annotations

import re


HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*$")


def parse_md_sections(md_text: str) -> dict[str, str]:
    """
    Markdown을 헤딩(#, ##, ###...)으로 나눠 섹션 맵을 만든다.
    키는 heading 텍스트 그대로 (예: "1. 파견 개요", "가. 목적 수행").
    중복 헤딩은 마지막 값이 우선.
    """
    sections: dict[str, str] = {}
    current_heading: str | None = None
    current_body: list[str] = []

    for raw in md_text.split("\n"):
        m = HEADING_RE.match(raw)
        if m:
            if current_heading is not None:
                sections[current_heading] = "\n".join(current_body).strip()
            current_heading = m.group(2).strip()
            current_body = []
        else:
            if current_heading is not None:
                current_body.append(raw)

    if current_heading is not None:
        sections[current_heading] = "\n".join(current_body).strip()

    return sections


def match_to_template_headings(
    md_sections: dict[str, str],
    template_headings: list[str],
) -> dict[str, str]:
    """
    템플릿 헤딩 목록에 MD 섹션을 매칭한다.
    1차: 정확 일치
    2차: 공백/구두점 제거 후 일치
    3차: MD 헤딩에 템플릿 헤딩이 포함되거나 그 역
    매칭 못한 템플릿 헤딩은 포함 안 됨.
    """
    def norm(s: str) -> str:
        return re.sub(r"[\s.,:;~\-()]+", "", s).lower()

    md_by_norm = {norm(k): (k, v) for k, v in md_sections.items()}
    result: dict[str, str] = {}
    for th in template_headings:
        if th in md_sections:
            result[th] = md_sections[th]
            continue
        key = norm(th)
        if key in md_by_norm:
            result[th] = md_by_norm[key][1]
            continue
        for md_key, md_val in md_sections.items():
            if th in md_key or md_key in th:
                result[th] = md_val
                break
    return result
