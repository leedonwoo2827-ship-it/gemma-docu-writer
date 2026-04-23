from __future__ import annotations

import argparse
import io
import json
import shutil
import sys
import tempfile
from pathlib import Path

# Ensure UTF-8 stdout/stderr so non-ASCII paths/text (e.g. Korean, em-dash) don't crash
# on legacy Windows consoles (cp949).
for _stream in ("stdout", "stderr"):
    s = getattr(sys, _stream, None)
    if s is not None and hasattr(s, "buffer") and getattr(s, "encoding", "").lower() != "utf-8":
        setattr(sys, _stream, io.TextIOWrapper(s.buffer, encoding="utf-8", errors="replace"))

from . import __version__
from .editor import fill_table, set_sp_text
from .mapper import Plan, _find_section_exemplar, _find_title_slot_on_slide, build_plan, format_plan
from .md_parser import Document, parse_md
from .pack import pack, unpack
from .qa import run_placeholder_check, run_visual_export
from .slide_duplicator import duplicate_slide, reorder_slides
from .slide_remover import drop_slides
from .slide_scanner import scan_unpacked


def _compute_order(doc: Document, plan: Plan) -> list[int]:
    """Figure out the final slide order: title → sections (with tables) → footer."""
    order: list[int] = []

    title_slide = next((a.slot.slide_idx for a in plan.titles if a.role == "title"), None)
    subtitle_slide = next((a.slot.slide_idx for a in plan.titles if a.role == "subtitle"), None)
    footer_slide = next((a.slot.slide_idx for a in plan.titles if a.role == "footer"), None)

    if title_slide is not None:
        order.append(title_slide)
    # Subtitle usually lives on the same slide as title; don't double-count.
    if subtitle_slide is not None and subtitle_slide not in order:
        order.append(subtitle_slide)

    # Map each matched table to the heading text it falls under.
    table_slide_by_heading: dict[str, int] = {}
    for ta in plan.tables:
        ph = doc.tables[ta.md_table_idx].preceding_heading
        table_slide_by_heading.setdefault(ph, ta.slot.slide_idx)

    # Walk headings in MD order.
    for hi, heading in enumerate(doc.headings):
        if hi < len(plan.headings):
            order.append(plan.headings[hi].source_slide_idx)
        tbl = table_slide_by_heading.get(heading)
        if tbl is not None and tbl not in order:
            order.append(tbl)

    if footer_slide is not None and footer_slide not in order:
        order.append(footer_slide)

    # Dedup preserving order.
    seen: set[int] = set()
    deduped: list[int] = []
    for idx in order:
        if idx not in seen:
            deduped.append(idx)
            seen.add(idx)
    return deduped


def _apply_plan(plan: Plan, catalog, doc: Document) -> None:
    # Text assignments (title/subtitle/footer).
    for asn in plan.titles:
        set_sp_text(asn.slot.sp_elem, asn.text)

    # Table assignments.
    for asn in plan.tables:
        md_table = doc.tables[asn.md_table_idx]
        fill_table(asn.slot.tbl_elem, md_table.rows, asn.col_map)

    # Heading assignments: edit the title text slot on each target slide.
    used_on_each_slide: set[int] = set()
    # Don't clobber title/subtitle/footer slots we've already written.
    for asn in plan.titles:
        try:
            used_on_each_slide.add(catalog.text_slots.index(asn.slot))
        except ValueError:
            pass
    for h in plan.headings:
        pair = _find_title_slot_on_slide(catalog, h.source_slide_idx, used_on_each_slide)
        if pair is None:
            continue
        idx, slot = pair
        set_sp_text(slot.sp_elem, h.text)
        used_on_each_slide.add(idx)


def _write_slides(slide_trees: dict, slide_paths: dict) -> None:
    for idx, tree in slide_trees.items():
        path = slide_paths[idx]
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="md2pptx",
        description="Deterministic Markdown → PPTX converter that preserves a template's design.",
    )
    parser.add_argument("template", type=Path, help="Template .pptx path")
    parser.add_argument("md", type=Path, help="Input markdown path")
    parser.add_argument("out", type=Path, help="Output .pptx path")
    parser.add_argument("--map", dest="mapping", type=Path, default=None,
                        help="Optional mapping JSON to override automatic assignments")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print the mapping plan and exit without writing output")
    parser.add_argument("--qa", action="store_true",
                        help="Run placeholder + visual QA checks on the output")
    parser.add_argument("--keep-unused", action="store_true",
                        help="Keep template slides that no MD content maps to "
                             "(default: drop them)")
    parser.add_argument("--version", action="version", version=f"md2pptx {__version__}")
    args = parser.parse_args(argv)

    if not args.template.exists():
        print(f"error: template not found: {args.template}", file=sys.stderr)
        return 2
    if not args.md.exists():
        print(f"error: markdown not found: {args.md}", file=sys.stderr)
        return 2

    doc = parse_md(args.md)

    work = Path(tempfile.mkdtemp(prefix="md2pptx_"))
    try:
        unpack(args.template, work)

        # Pre-scan to find the section-divider exemplar, then pre-duplicate
        # one slide per extra H2 before building the real plan.
        pre_catalog = scan_unpacked(work)
        exemplar = _find_section_exemplar(pre_catalog)
        heading_slide_indices: list[int] = []
        if exemplar is not None:
            for i, _ in enumerate(doc.headings):
                if i == 0:
                    heading_slide_indices.append(exemplar)
                else:
                    new_idx = duplicate_slide(work, exemplar)
                    heading_slide_indices.append(new_idx)

        # Re-scan after duplication so catalog has fresh references.
        catalog = scan_unpacked(work)
        plan = build_plan(doc, catalog)
        # Overwrite the heading plan's source_slide_idx with the actual (pre-duplicated) indices.
        for hp, actual_idx in zip(plan.headings, heading_slide_indices):
            hp.source_slide_idx = actual_idx
            hp.is_duplicate = False  # from the apply step's POV, all target slides exist now

        print(format_plan(plan, doc))

        if args.dry_run:
            return 0

        if args.mapping:
            # Reserved for future overrides; read and discard for now to validate JSON.
            json.loads(args.mapping.read_text(encoding="utf-8"))

        _apply_plan(plan, catalog, doc)
        _write_slides(catalog.slide_trees, catalog.slide_paths)

        order = _compute_order(doc, plan)

        if not args.keep_unused:
            keep = set(order) | plan.used_slide_indices()
            if not keep:
                keep = {1}
            dropped = drop_slides(work, keep)
            if dropped:
                print(f"dropped {len(dropped)} unused slide(s): {dropped}")

        if order:
            reorder_slides(work, order)
            print(f"final slide order: {order}")

        pack(work, args.out)
        print(f"\nwrote: {args.out}")

        if plan.unmatched_tables:
            print(
                f"warning: {len(plan.unmatched_tables)} MD table(s) had no matching template slot: "
                f"{plan.unmatched_tables}",
                file=sys.stderr,
            )

        if args.qa:
            hits = run_placeholder_check(args.out)
            if hits:
                print("QA: possible placeholder leftovers:", file=sys.stderr)
                for h in hits:
                    print(f"  {h}", file=sys.stderr)
                return 1
            print("QA: no placeholder leftovers detected.")
            qa_dir = args.out.parent / f"{args.out.stem}_qa"
            pdf = run_visual_export(args.out, qa_dir)
            if pdf:
                print(f"QA: visual export → {qa_dir}/")
            else:
                print("QA: LibreOffice/Poppler not found, skipped visual export.")

        return 0
    finally:
        shutil.rmtree(work, ignore_errors=True)


def convert(
    template: str,
    md: str,
    out: str,
    dry_run: bool = False,
    keep_unused: bool = False,
) -> dict:
    """
    FastAPI 등 프로그래매틱 호출용 래퍼.
    main() 의 argparse 경로를 우회하고 내부 파이프라인을 직접 실행하여
    결과 메타데이터를 dict 로 반환.

    Returns: {
        "output_path": str,
        "slides_final": list[int],           # 최종 포함된 원본 슬라이드 번호 (재정렬 전)
        "slides_count": int,
        "slides_dropped": list[int],
        "headings_matched": list[str],
        "tables_matched": list[dict],        # [{md_idx, template_slide, md_headers, tpl_headers}]
        "tables_unmatched": list[int],       # 매칭 실패한 MD 테이블 인덱스
        "plan_text": str,                    # format_plan() 출력 (사람 읽기용)
    }
    """
    tpl = Path(template)
    md_p = Path(md)
    out_p = Path(out)
    if not tpl.exists():
        raise FileNotFoundError(f"template not found: {tpl}")
    if not md_p.exists():
        raise FileNotFoundError(f"markdown not found: {md_p}")

    doc = parse_md(md_p)

    work = Path(tempfile.mkdtemp(prefix="md2pptx_"))
    try:
        unpack(tpl, work)

        pre_catalog = scan_unpacked(work)
        exemplar = _find_section_exemplar(pre_catalog)
        heading_slide_indices: list[int] = []
        if exemplar is not None:
            for i, _ in enumerate(doc.headings):
                if i == 0:
                    heading_slide_indices.append(exemplar)
                else:
                    new_idx = duplicate_slide(work, exemplar)
                    heading_slide_indices.append(new_idx)

        catalog = scan_unpacked(work)
        plan = build_plan(doc, catalog)
        for hp, actual_idx in zip(plan.headings, heading_slide_indices):
            hp.source_slide_idx = actual_idx
            hp.is_duplicate = False

        plan_text = format_plan(plan, doc)

        if dry_run:
            return {
                "output_path": "",
                "slides_final": [],
                "slides_count": 0,
                "slides_dropped": [],
                "headings_matched": [h.text for h in plan.headings],
                "tables_matched": [
                    {
                        "md_idx": t.md_table_idx,
                        "template_slide": t.slot.slide_idx,
                        "md_headers": doc.tables[t.md_table_idx].headers,
                        "score": t.score,
                    }
                    for t in plan.tables
                ],
                "tables_unmatched": plan.unmatched_tables,
                "plan_text": plan_text,
                "dry_run": True,
            }

        _apply_plan(plan, catalog, doc)
        _write_slides(catalog.slide_trees, catalog.slide_paths)

        order = _compute_order(doc, plan)

        dropped: list[int] = []
        if not keep_unused:
            keep = set(order) | plan.used_slide_indices()
            if not keep:
                keep = {1}
            dropped = drop_slides(work, keep)

        if order:
            reorder_slides(work, order)

        out_p.parent.mkdir(parents=True, exist_ok=True)
        pack(work, out_p)

        return {
            "output_path": str(out_p),
            "slides_final": order,
            "slides_count": len(order) if order else 0,
            "slides_dropped": dropped,
            "headings_matched": [h.text for h in plan.headings],
            "tables_matched": [
                {
                    "md_idx": t.md_table_idx,
                    "template_slide": t.slot.slide_idx,
                    "md_headers": doc.tables[t.md_table_idx].headers,
                    "score": t.score,
                }
                for t in plan.tables
            ],
            "tables_unmatched": plan.unmatched_tables,
            "plan_text": plan_text,
            "dry_run": False,
        }
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
