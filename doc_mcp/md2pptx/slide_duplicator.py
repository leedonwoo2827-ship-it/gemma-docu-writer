from __future__ import annotations

import re
import shutil
from pathlib import Path

from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

SLIDE_NUM_RE = re.compile(r"slide(\d+)\.xml$")


def _parse(path: Path) -> etree._ElementTree:
    return etree.parse(str(path))


def _write(tree: etree._ElementTree, path: Path) -> None:
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


def _next_slide_index(slides_dir: Path) -> int:
    used: set[int] = set()
    for f in slides_dir.glob("slide*.xml"):
        m = SLIDE_NUM_RE.search(f.name)
        if m:
            used.add(int(m.group(1)))
    i = 1
    while i in used:
        i += 1
    return i


def _next_rid(rels_root: etree._Element) -> str:
    used: set[int] = set()
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        rid = rel.get("Id") or ""
        if rid.startswith("rId") and rid[3:].isdigit():
            used.add(int(rid[3:]))
    i = 1
    while i in used:
        i += 1
    return f"rId{i}"


def _next_sld_id(sldIdLst: etree._Element) -> int:
    used: set[int] = set()
    for sldId in sldIdLst.findall(f"{{{P}}}sldId"):
        try:
            used.add(int(sldId.get("id") or "0"))
        except ValueError:
            pass
    i = 256
    while i in used:
        i += 1
    return i


def duplicate_slide(unpacked: Path, source_idx: int) -> int:
    """Duplicate an existing slide and register it with the presentation.

    Returns the new slide's numeric index (e.g., 14). The new slide is appended
    to the end of <p:sldIdLst>; use reorder_slides() to move it.
    """
    unpacked = Path(unpacked)
    slides_dir = unpacked / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"
    pres_xml = unpacked / "ppt" / "presentation.xml"
    pres_rels = unpacked / "ppt" / "_rels" / "presentation.xml.rels"
    ct_xml = unpacked / "[Content_Types].xml"

    src_slide = slides_dir / f"slide{source_idx}.xml"
    src_rels = rels_dir / f"slide{source_idx}.xml.rels"
    if not src_slide.exists():
        raise FileNotFoundError(src_slide)

    new_idx = _next_slide_index(slides_dir)
    dst_slide = slides_dir / f"slide{new_idx}.xml"
    dst_rels = rels_dir / f"slide{new_idx}.xml.rels"

    shutil.copyfile(src_slide, dst_slide)
    if src_rels.exists():
        rels_dir.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(src_rels, dst_rels)
        # Strip the notesSlide relationship from the copy — two slides should
        # not share the same notes part (the notes itself only back-references
        # the original slide, so the link would be mismatched anyway).
        try:
            tree = _parse(dst_rels)
            root = tree.getroot()
            changed = False
            for rel in list(root.findall(f"{{{REL_NS}}}Relationship")):
                if (rel.get("Type") or "").endswith("/notesSlide"):
                    root.remove(rel)
                    changed = True
            if changed:
                _write(tree, dst_rels)
        except etree.XMLSyntaxError:
            pass

    # Register in presentation.xml.rels
    rels_tree = _parse(pres_rels)
    rels_root = rels_tree.getroot()
    new_rid = _next_rid(rels_root)
    rel = etree.SubElement(rels_root, f"{{{REL_NS}}}Relationship")
    rel.set("Id", new_rid)
    rel.set(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    )
    rel.set("Target", f"slides/slide{new_idx}.xml")
    _write(rels_tree, pres_rels)

    # Register in [Content_Types].xml as an Override
    ct_tree = _parse(ct_xml)
    ct_root = ct_tree.getroot()
    override = etree.SubElement(ct_root, f"{{{CT}}}Override")
    override.set("PartName", f"/ppt/slides/slide{new_idx}.xml")
    override.set(
        "ContentType",
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    )
    _write(ct_tree, ct_xml)

    # Register in presentation.xml sldIdLst (appended to the end)
    pres_tree = _parse(pres_xml)
    sldIdLst = pres_tree.getroot().find(f"{{{P}}}sldIdLst")
    if sldIdLst is None:
        raise RuntimeError("presentation.xml has no sldIdLst")
    new_sldId = etree.SubElement(sldIdLst, f"{{{P}}}sldId")
    new_sldId.set("id", str(_next_sld_id(sldIdLst)))
    new_sldId.set(f"{{{R}}}id", new_rid)
    _write(pres_tree, pres_xml)

    return new_idx


def reorder_slides(unpacked: Path, ordered_slide_indices: list[int]) -> None:
    """Reorder <p:sldId> entries to match ordered_slide_indices.

    Any slide index not present in the list keeps its position at the end (stable).
    """
    unpacked = Path(unpacked)
    pres_xml = unpacked / "ppt" / "presentation.xml"
    pres_rels = unpacked / "ppt" / "_rels" / "presentation.xml.rels"

    rels_tree = _parse(pres_rels)
    rid_to_idx: dict[str, int] = {}
    for rel in rels_tree.getroot().findall(f"{{{REL_NS}}}Relationship"):
        rtype = rel.get("Type") or ""
        if not rtype.endswith("/slide"):
            continue
        target = rel.get("Target") or ""
        m = SLIDE_NUM_RE.search(target)
        if m:
            rid_to_idx[rel.get("Id") or ""] = int(m.group(1))

    pres_tree = _parse(pres_xml)
    sldIdLst = pres_tree.getroot().find(f"{{{P}}}sldIdLst")
    if sldIdLst is None:
        return
    # Read current order.
    entries: list[tuple[int, etree._Element]] = []
    for sldId in list(sldIdLst.findall(f"{{{P}}}sldId")):
        rid = sldId.get(f"{{{R}}}id") or ""
        idx = rid_to_idx.get(rid, -1)
        entries.append((idx, sldId))
        sldIdLst.remove(sldId)

    # Partition: items in ordered list first (in that order), then the rest stable.
    by_idx = {idx: elem for idx, elem in entries}
    placed: set[int] = set()
    for idx in ordered_slide_indices:
        if idx in by_idx and idx not in placed:
            sldIdLst.append(by_idx[idx])
            placed.add(idx)
    for idx, elem in entries:
        if idx not in placed:
            sldIdLst.append(elem)
            placed.add(idx)

    _write(pres_tree, pres_xml)
