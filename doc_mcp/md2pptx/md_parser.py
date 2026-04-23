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
class Document:
    title: str = ""
    subtitle: str = ""
    footer: str = ""
    headings: list[str] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)


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

        # Paragraph (capture top-level italics as subtitle/footer)
        if tok.type == "paragraph_open":
            inline = tokens[i + 1]
            raw = _collect_inline(inline).strip()
            # If entire paragraph wrapped in em_open/em_close → italic → subtitle/footer.
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
            i += 3
            continue

        # Tables
        if tok.type == "table_open":
            headers, rows, j = _parse_table(tokens, i)
            doc.tables.append(Table(headers=headers, rows=rows, preceding_heading=last_heading))
            i = j
            continue

        i += 1

    return doc


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
