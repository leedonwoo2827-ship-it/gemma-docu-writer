from __future__ import annotations

import copy
from typing import Optional

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"a": A}


def _first_run(txBody: etree._Element) -> Optional[etree._Element]:
    for p in txBody.findall(f"{{{A}}}p"):
        r = p.find(f"{{{A}}}r")
        if r is not None:
            return r
    return None


def _first_paragraph(txBody: etree._Element) -> Optional[etree._Element]:
    return txBody.find(f"{{{A}}}p")


def _new_run_from(template_run: etree._Element, text: str) -> etree._Element:
    """Clone a run's rPr, but replace its text with the given value."""
    new_r = etree.SubElement(template_run.getparent(), f"{{{A}}}r")
    new_r.getparent().remove(new_r)  # detached; we'll parent it explicitly
    rPr = template_run.find(f"{{{A}}}rPr")
    if rPr is not None:
        new_r.append(copy.deepcopy(rPr))
    t = etree.SubElement(new_r, f"{{{A}}}t")
    t.text = text
    return new_r


def _set_txBody_text(txBody: etree._Element, text: str) -> None:
    """Replace the text content of a <a:txBody> while preserving the first run's rPr.

    Multi-line input (text containing '\n') is split into separate <a:p> paragraphs.
    All original paragraphs/runs are cleared; the first original run is used as the
    style template for every new paragraph.
    """
    first_p = _first_paragraph(txBody)
    template_run = _first_run(txBody) if first_p is not None else None
    # Capture paragraph-level pPr from the first paragraph, if any.
    pPr_template = first_p.find(f"{{{A}}}pPr") if first_p is not None else None

    # Remove all existing <a:p> paragraphs from txBody.
    for p in list(txBody.findall(f"{{{A}}}p")):
        txBody.remove(p)

    lines = text.split("\n") if text else [""]
    for line in lines:
        p = etree.SubElement(txBody, f"{{{A}}}p")
        if pPr_template is not None:
            p.append(copy.deepcopy(pPr_template))
        if line:
            if template_run is not None:
                rPr = template_run.find(f"{{{A}}}rPr")
                r = etree.SubElement(p, f"{{{A}}}r")
                if rPr is not None:
                    r.append(copy.deepcopy(rPr))
                t = etree.SubElement(r, f"{{{A}}}t")
                t.text = line
            else:
                r = etree.SubElement(p, f"{{{A}}}r")
                t = etree.SubElement(r, f"{{{A}}}t")
                t.text = line
        else:
            # Empty line → keep paragraph with endParaRPr only (valid OOXML).
            if template_run is not None:
                rPr = template_run.find(f"{{{A}}}rPr")
                if rPr is not None:
                    endParaRPr = etree.SubElement(p, f"{{{A}}}endParaRPr")
                    for k, v in rPr.attrib.items():
                        endParaRPr.set(k, v)


def set_sp_text(sp: etree._Element, text: str) -> None:
    """Replace the text content of a <p:sp> shape, preserving run styling."""
    # Find the txBody under p:sp (can be p:txBody or a:txBody depending on part).
    txBody = sp.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}txBody")
    if txBody is None:
        txBody = sp.find(f".//{{{A}}}txBody")
    if txBody is None:
        return
    _set_txBody_text(txBody, text)


def set_cell_text(tc: etree._Element, text: str) -> None:
    """Replace the text content of a <a:tc> table cell, preserving run styling."""
    txBody = tc.find(f"{{{A}}}txBody")
    if txBody is None:
        return
    _set_txBody_text(txBody, text)


def fill_table(
    tbl: etree._Element,
    md_rows: list[list[str]],
    col_map: list[Optional[int]],
) -> None:
    """Fill a <a:tbl> with rows from MD.

    Assumptions:
      - Row 0 is the header; not modified.
      - Row 1 (if present) is the 'body template' — its XML is cloned to extend.
      - If there are fewer MD rows than existing body rows, excess body rows are removed.
      - If there are more, the last body row's XML is deep-cloned.

    col_map[i] gives the template column index for MD column i (or None to skip).
    """
    trs = tbl.findall(f"{{{A}}}tr")
    if len(trs) < 2:
        # Degenerate: no body template row to clone. Abort gracefully.
        return

    body_template = trs[1]
    n_body_needed = len(md_rows)
    n_body_existing = len(trs) - 1

    # Ensure exactly n_body_needed body rows exist.
    if n_body_needed > n_body_existing:
        # Clone body_template for extra rows.
        for _ in range(n_body_needed - n_body_existing):
            clone = copy.deepcopy(body_template)
            tbl.append(clone)
    elif n_body_needed < n_body_existing:
        # Remove excess body rows (from the end).
        for extra in trs[1 + n_body_needed :]:
            tbl.remove(extra)

    # Re-fetch rows (now the exact length we need).
    trs = tbl.findall(f"{{{A}}}tr")
    body_rows = trs[1:]

    # Clear all body cells first so leftover content from the template
    # (e.g., unmapped decorative columns) does not leak into the output.
    for tr in body_rows:
        for tc in tr.findall(f"{{{A}}}tc"):
            set_cell_text(tc, "")

    for mi_row, (md_row, tr) in enumerate(zip(md_rows, body_rows)):
        tcs = tr.findall(f"{{{A}}}tc")
        for mi_col, md_value in enumerate(md_row):
            if mi_col >= len(col_map):
                continue
            tpl_col = col_map[mi_col]
            if tpl_col is None or tpl_col >= len(tcs):
                continue
            set_cell_text(tcs[tpl_col], md_value)
