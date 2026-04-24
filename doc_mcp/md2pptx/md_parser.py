from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from markdown_it import MarkdownIt


@dataclass
class Table:
    headers: list[str]
    rows: list[list[str]]
    preceding_heading: str = ""

    @property
    def ncols(self) -> int:
        return len(self.headers)

    @property
    def nrows(self) -> int:
        return len(self.rows)


@dataclass
class BodyBlock:
    """H2 섹션 아래의 줄글·bullet·ordered-list 묶음.
    양식의 body text shape 에 주입할 후보."""
    heading: str                   # 소속 H2 (없으면 빈 문자열)
    kind: str                       # "prose" | "bullets" | "ordered"
    text: str                       # 단락 원문 또는 bullet join 된 결과
    bullets: list[str] = field(default_factory=list)   # list 항목들 (kind=bullets/ordered 일 때)


@dataclass
class Document:
    title: str = ""
    subtitle: str = ""
    footer: str = ""
    headings: list[str] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    body_blocks: list[BodyBlock] = field(default_factory=list)


def _text_of(tokens: list, start: int, stop: int) -> str:
    parts: list[str] = []
    for t in tokens[start:stop]:
        if t.type == "inline":
            for child in t.children or []:
                if child.type == "text":
                    parts.append(child.content)
                elif child.type == "code_inline":
                    parts.append(child.content)
                elif child.type == "softbreak" or child.type == "hardbreak":
                    parts.append("\n")
                elif child.type == "html_inline" and child.content.lower().startswith("<br"):
                    parts.append("\n")
        elif t.type == "text":
            parts.append(t.content)
    return "".join(parts).strip()


def _collect_inline(token) -> str:
    """Render an inline token to plain text with <br>→newline."""
    parts: list[str] = []
    for child in token.children or []:
        tag = child.type
        if tag == "text":
            parts.append(child.content)
        elif tag == "code_inline":
            parts.append(child.content)
        elif tag in ("softbreak", "hardbreak"):
            parts.append("\n")
        elif tag == "html_inline" and child.content.lower().startswith("<br"):
            parts.append("\n")
        elif tag in ("em_open", "em_close", "strong_open", "strong_close"):
            continue
    return "".join(parts)


def parse_md(md_path: Path | str) -> Document:
    md_path = Path(md_path)
    text = md_path.read_text(encoding="utf-8")
    md = MarkdownIt("commonmark", {"html": True}).enable("table")
    tokens = md.parse(text)

    doc = Document()
    last_heading = ""
    i = 0
    n = len(tokens)

    while i < n:
        tok = tokens[i]

        # Headings
        if tok.type == "heading_open":
            level = int(tok.tag[1])
            inline = tokens[i + 1]
            content = _collect_inline(inline).strip()
            if level == 1 and not doc.title:
                doc.title = content
            else:
                doc.headings.append(content)
                last_heading = content
            i += 3
            continue

        # Paragraph — italic wrapped → subtitle/footer; 그 외 본문 단락 → body_block(prose)
        if tok.type == "paragraph_open":
            inline = tokens[i + 1]
            raw = _collect_inline(inline).strip()
            kids = inline.children or []
            wrapped_em = (
                len(kids) >= 2
                and kids[0].type == "em_open"
                and kids[-1].type == "em_close"
            )
            if wrapped_em:
                if not doc.subtitle and not doc.tables:
                    doc.subtitle = raw
                else:
                    doc.footer = raw
            elif raw:
                # 일반 본문 단락 → 해당 heading 섹션의 body_block 에 추가
                doc.body_blocks.append(
                    BodyBlock(heading=last_heading, kind="prose", text=raw)
                )
            i += 3
            continue

        # Bullet list → body_block(bullets)
        if tok.type == "bullet_list_open":
            bullets, j = _collect_list_items(tokens, i)
            if bullets:
                doc.body_blocks.append(
                    BodyBlock(
                        heading=last_heading,
                        kind="bullets",
                        text="\n".join(f"• {b}" for b in bullets),
                        bullets=bullets,
                    )
                )
            i = j
            continue

        # Ordered list → body_block(ordered)
        if tok.type == "ordered_list_open":
            bullets, j = _collect_list_items(tokens, i)
            if bullets:
                doc.body_blocks.append(
                    BodyBlock(
                        heading=last_heading,
                        kind="ordered",
                        text="\n".join(f"{n_+1}. {b}" for n_, b in enumerate(bullets)),
                        bullets=bullets,
                    )
                )
            i = j
            continue

        # Tables
        if tok.type == "table_open":
            headers, rows, j = _parse_table(tokens, i)
            doc.tables.append(Table(headers=headers, rows=rows, preceding_heading=last_heading))
            i = j
            continue

        i += 1

    return doc


def _collect_list_items(tokens: list, start: int) -> tuple[list[str], int]:
    """bullet_list_open / ordered_list_open 부터 대응 close 까지 item 텍스트 수집.
    중첩 list 는 첫 레벨 flat join 하되 prefix 처리 없이 원문 유지.
    """
    items: list[str] = []
    i = start + 1
    depth = 1
    current_item_parts: list[str] = []
    in_item = False
    while i < len(tokens):
        t = tokens[i]
        if t.type in ("bullet_list_open", "ordered_list_open"):
            depth += 1
        elif t.type in ("bullet_list_close", "ordered_list_close"):
            depth -= 1
            if depth == 0:
                return items, i + 1
        elif t.type == "list_item_open":
            in_item = True
            current_item_parts = []
        elif t.type == "list_item_close":
            joined = " ".join(p.strip() for p in current_item_parts if p.strip())
            if joined:
                items.append(joined)
            in_item = False
            current_item_parts = []
        elif t.type == "paragraph_open" and in_item:
            inline = tokens[i + 1]
            current_item_parts.append(_collect_inline(inline))
            i += 3
            continue
        elif t.type == "inline" and in_item:
            current_item_parts.append(_collect_inline(t))
        i += 1
    return items, i


def _parse_table(tokens: list, start: int) -> tuple[list[str], list[list[str]], int]:
    headers: list[str] = []
    rows: list[list[str]] = []
    current_row: list[str] = []
    in_header = False
    in_body = False
    i = start + 1
    while i < len(tokens):
        t = tokens[i]
        if t.type == "table_close":
            return headers, rows, i + 1
        if t.type == "thead_open":
            in_header = True
        elif t.type == "thead_close":
            in_header = False
        elif t.type == "tbody_open":
            in_body = True
        elif t.type == "tbody_close":
            in_body = False
        elif t.type == "tr_open":
            current_row = []
        elif t.type == "tr_close":
            if in_header:
                headers = current_row
            elif in_body:
                rows.append(current_row)
            current_row = []
        elif t.type in ("th_open", "td_open"):
            inline = tokens[i + 1]
            cell = _collect_inline(inline).strip()
            current_row.append(cell)
            i += 2  # skip inline + close
        i += 1
    return headers, rows, i
