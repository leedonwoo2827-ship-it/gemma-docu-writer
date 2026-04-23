from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}

SLIDE_NUM_RE = re.compile(r"slide(\d+)\.xml$")


@dataclass
class TextSlot:
    slide_idx: int          # 1-based (slide1.xml → 1)
    slide_path: Path
    sp_path: str            # lxml getpath() for later round-trip
    text: str               # current text in the shape (for role heuristics)
    sp_elem: etree._Element = field(repr=False, default=None)


@dataclass
class TableSlot:
    slide_idx: int
    slide_path: Path
    tbl_path: str
    headers: list[str]
    n_cols: int
    n_rows: int
    tbl_elem: etree._Element = field(repr=False, default=None)


@dataclass
class SlotCatalog:
    text_slots: list[TextSlot] = field(default_factory=list)
    table_slots: list[TableSlot] = field(default_factory=list)
    slide_trees: dict[int, etree._ElementTree] = field(default_factory=dict, repr=False)
    slide_paths: dict[int, Path] = field(default_factory=dict, repr=False)


def _cell_text(tc: etree._Element) -> str:
    return "".join(t.text or "" for t in tc.findall(".//a:t", NS)).strip()


def _sp_text(sp: etree._Element) -> str:
    return "".join(t.text or "" for t in sp.findall(".//a:t", NS)).strip()


def scan_unpacked(unpacked_dir: Path) -> SlotCatalog:
    unpacked_dir = Path(unpacked_dir)
    slides_dir = unpacked_dir / "ppt" / "slides"
    catalog = SlotCatalog()
    if not slides_dir.exists():
        return catalog

    slide_files = sorted(
        slides_dir.glob("slide*.xml"),
        key=lambda p: int(SLIDE_NUM_RE.search(p.name).group(1)),
    )
    for slide_path in slide_files:
        m = SLIDE_NUM_RE.search(slide_path.name)
        if not m:
            continue
        idx = int(m.group(1))
        tree = etree.parse(str(slide_path))
        catalog.slide_trees[idx] = tree
        catalog.slide_paths[idx] = slide_path
        root = tree.getroot()

        # Tables: <a:graphicFrame> → <a:graphic> → <a:graphicData> → <a:tbl>
        for tbl in root.iter(f"{{{A}}}tbl"):
            trs = tbl.findall(f"{{{A}}}tr")
            if not trs:
                continue
            header_row = trs[0]
            headers = [_cell_text(tc) for tc in header_row.findall(f"{{{A}}}tc")]
            n_cols = len(headers)
            n_rows = len(trs)
            catalog.table_slots.append(
                TableSlot(
                    slide_idx=idx,
                    slide_path=slide_path,
                    tbl_path=tree.getpath(tbl),
                    headers=headers,
                    n_cols=n_cols,
                    n_rows=n_rows,
                    tbl_elem=tbl,
                )
            )

        # Text shapes: <p:sp> with a <p:txBody> containing any <a:t>
        for sp in root.iter(f"{{{P}}}sp"):
            if sp.find(".//a:t", NS) is None:
                continue
            txt = _sp_text(sp)
            catalog.text_slots.append(
                TextSlot(
                    slide_idx=idx,
                    slide_path=slide_path,
                    sp_path=tree.getpath(sp),
                    text=txt,
                    sp_elem=sp,
                )
            )

    return catalog
