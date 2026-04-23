from __future__ import annotations

import re
from pathlib import Path

from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": P, "r": R, "ct": CT, "rel": REL_NS}

SLIDE_NUM_RE = re.compile(r"slide(\d+)\.xml$")


def _parse(path: Path) -> etree._ElementTree:
    return etree.parse(str(path))


def _write(tree: etree._ElementTree, path: Path) -> None:
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


def drop_slides(unpacked: Path, keep_slide_indices: set[int]) -> list[int]:
    """Remove slides NOT in keep_slide_indices.

    Safely updates:
      - ppt/presentation.xml   (<p:sldIdLst>)
      - ppt/_rels/presentation.xml.rels
      - [Content_Types].xml    (<Override>)
      - delete ppt/slides/slide{N}.xml and its .rels

    Returns a sorted list of the slide indices that were actually dropped.
    """
    unpacked = Path(unpacked)
    pres_xml = unpacked / "ppt" / "presentation.xml"
    pres_rels = unpacked / "ppt" / "_rels" / "presentation.xml.rels"
    ct_xml = unpacked / "[Content_Types].xml"
    slides_dir = unpacked / "ppt" / "slides"

    if not pres_xml.exists() or not pres_rels.exists() or not ct_xml.exists():
        return []

    pres_tree = _parse(pres_xml)
    rels_tree = _parse(pres_rels)
    ct_tree = _parse(ct_xml)

    # Build rId → slide path (and slide index) map from the presentation rels.
    rid_to_path: dict[str, str] = {}
    rid_to_idx: dict[str, int] = {}
    for rel in rels_tree.getroot().findall(f"{{{REL_NS}}}Relationship"):
        target = rel.get("Target") or ""
        rtype = rel.get("Type") or ""
        rid = rel.get("Id") or ""
        if rtype.endswith("/slide"):
            rid_to_path[rid] = target
            m = SLIDE_NUM_RE.search(target)
            if m:
                rid_to_idx[rid] = int(m.group(1))

    # Figure out which rIds to drop.
    drop_rids: list[str] = []
    dropped_indices: list[int] = []
    for rid, idx in rid_to_idx.items():
        if idx not in keep_slide_indices:
            drop_rids.append(rid)
            dropped_indices.append(idx)

    if not drop_rids:
        return []

    # 1) Remove <p:sldId> entries pointing to dropped rIds.
    sldIdLst = pres_tree.getroot().find(f"{{{P}}}sldIdLst")
    if sldIdLst is not None:
        for sldId in list(sldIdLst.findall(f"{{{P}}}sldId")):
            rid_attr = sldId.get(f"{{{R}}}id")
            if rid_attr in drop_rids:
                sldIdLst.remove(sldId)

    # 2) Remove matching Relationship entries.
    rels_root = rels_tree.getroot()
    for rel in list(rels_root.findall(f"{{{REL_NS}}}Relationship")):
        if rel.get("Id") in drop_rids:
            rels_root.remove(rel)

    # 3) Remove matching <Override> entries from [Content_Types].xml and
    # delete the slide XML + its .rels on disk.
    drop_paths = {
        f"/ppt/{rid_to_path[rid].lstrip('/')}" for rid in drop_rids
    }
    # Targets in relationships are usually relative like "slides/slide4.xml";
    # Content_Types overrides use absolute "/ppt/slides/slide4.xml". Normalize.
    normalized = set()
    for p in drop_paths:
        # collapse double slashes, ensure leading slash
        normalized.add("/" + p.strip("/").replace("ppt//", "ppt/"))
    drop_paths = normalized

    ct_root = ct_tree.getroot()
    for override in list(ct_root.findall(f"{{{CT}}}Override")):
        if override.get("PartName") in drop_paths:
            ct_root.remove(override)

    # Gather the notesSlide targets referenced by each dropped slide's rels,
    # so we can clean those up too (otherwise they become orphans that point
    # at deleted slides — PowerPoint rejects such packages).
    orphan_notes: set[Path] = set()
    for rid in drop_rids:
        idx = rid_to_idx[rid]
        sf = slides_dir / f"slide{idx}.xml"
        rf = slides_dir / "_rels" / f"slide{idx}.xml.rels"
        if rf.exists():
            try:
                srels = _parse(rf)
                for rel in srels.getroot().findall(f"{{{REL_NS}}}Relationship"):
                    if (rel.get("Type") or "").endswith("/notesSlide"):
                        target = rel.get("Target") or ""
                        # OOXML convention: the Target in a .rels file is
                        # relative to the PART (the slide file), not to the
                        # .rels file itself. e.g. "../notesSlides/notesSlide3.xml"
                        # from ppt/slides/slide5.xml → ppt/notesSlides/notesSlide3.xml
                        notes_path = (sf.parent / target).resolve()
                        orphan_notes.add(notes_path)
            except etree.XMLSyntaxError:
                pass
        if sf.exists():
            sf.unlink()
        if rf.exists():
            rf.unlink()

    # Delete orphan notesSlide files + their .rels + their Content_Types Override.
    notes_drop_paths: set[str] = set()
    for np in orphan_notes:
        if not np.exists():
            continue
        rel_np = np.with_name(np.name).parent / "_rels" / (np.name + ".rels")
        try:
            np.unlink()
        except OSError:
            pass
        try:
            if rel_np.exists():
                rel_np.unlink()
        except OSError:
            pass
        # PartName in [Content_Types].xml is absolute from package root.
        try:
            rel_to_pkg = np.relative_to(unpacked).as_posix()
            notes_drop_paths.add("/" + rel_to_pkg)
        except ValueError:
            pass

    ct_root = ct_tree.getroot()
    if notes_drop_paths:
        for override in list(ct_root.findall(f"{{{CT}}}Override")):
            if override.get("PartName") in notes_drop_paths:
                ct_root.remove(override)

    _write(pres_tree, pres_xml)
    _write(rels_tree, pres_rels)
    _write(ct_tree, ct_xml)

    return sorted(dropped_indices)
