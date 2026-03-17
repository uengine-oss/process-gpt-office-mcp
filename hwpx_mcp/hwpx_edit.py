import logging
import os
import tempfile
from xml.etree import ElementTree as ET

from .core.html_edit import extract_fills_and_ids
from .core.filler import apply_fills
from .core.parser import parse_section
from .core.xml_utils import tag as core_tag
from .io.file import extract_hwpx, find_section_files, repack_hwpx


logger = logging.getLogger("process-gpt-office-mcp")


def _collect_tables_in_order(root: ET.Element) -> list[ET.Element]:
    tables: list[ET.Element] = []
    for child in list(root):
        if core_tag(child) != "p":
            continue
        for run in list(child):
            if core_tag(run) != "run":
                continue
            for gc in list(run):
                if core_tag(gc) == "tbl":
                    tables.append(gc)
    return tables


def _is_empty_paragraph(p_elem: ET.Element) -> bool:
    for ch in p_elem.iter():
        if core_tag(ch) == "tbl":
            return False
        if core_tag(ch) == "t" and (ch.text or "").strip():
            return False
    return True


def _remove_adjacent_empty_paragraphs(parent: ET.Element, start_index: int) -> int:
    removed = 0
    children = list(parent)
    left = start_index - 1
    while left >= 0:
        node = children[left]
        if core_tag(node) == "p" and _is_empty_paragraph(node):
            if node in list(parent):
                parent.remove(node)
                removed += 1
            left -= 1
            continue
        break
    right = start_index
    children = list(parent)
    while right < len(children):
        node = children[right]
        if core_tag(node) == "p" and _is_empty_paragraph(node):
            if node in list(parent):
                parent.remove(node)
                removed += 1
            right += 1
            continue
        break
    return removed


def apply_html_edits_to_hwpx(
    hwpx_path: str,
    output_path: str,
    edited_html: str,
) -> str:
    if not edited_html:
        raise ValueError("edited_html is required")

    edits, present_ids = extract_fills_and_ids(edited_html)
    if not present_ids:
        raise ValueError("edited_html에 data-id가 없습니다.")
    logger.info(
        "html_edit: fills=%d present_ids=%d",
        len(edits),
        len(present_ids),
    )

    with tempfile.TemporaryDirectory() as tmp_dir:
        extract_dir = os.path.join(tmp_dir, "hwpx")
        compress_info, file_order = extract_hwpx(hwpx_path, extract_dir)
        section_files = find_section_files(extract_dir)
        if not section_files:
            raise RuntimeError("No section files found in HWPX")

        global_offset = 0
        for sf in section_files:
            nodes, tree, parent_map, t_ns = parse_section(sf, style_maps=None)
            if not nodes:
                continue
            fills: list[dict] = []
            table_nodes: dict[int, list[int]] = {}
            for node in nodes:
                if node.id is None:
                    continue
                global_id = global_offset + node.id
                if getattr(node, "type", "") == "table_cell" and node.table_idx >= 0:
                    table_nodes.setdefault(node.table_idx, []).append(global_id)
                if global_id in edits:
                    fills.append({"id": node.id, "new_text": edits[global_id]})
            remove_table_indices: set[int] = set()
            for tbl_idx, ids in table_nodes.items():
                if ids and not any(node_id in present_ids for node_id in ids):
                    remove_table_indices.add(tbl_idx)
            logger.info(
                "section=%s nodes=%d tables=%d fills=%d remove_tables=%s offset=%d",
                os.path.basename(sf),
                len(nodes),
                len(table_nodes),
                len(fills),
                sorted(remove_table_indices),
                global_offset,
            )
            if remove_table_indices:
                root = tree.getroot()
                parent_map_local = {c: p for p in root.iter() for c in p}
                tables_in_order = _collect_tables_in_order(root)
                removed = 0
                for tbl_idx in sorted(remove_table_indices, reverse=True):
                    if tbl_idx < 0 or tbl_idx >= len(tables_in_order):
                        continue
                    tbl = tables_in_order[tbl_idx]
                    parent = parent_map_local.get(tbl)
                    if parent is not None and tbl in list(parent):
                        parent.remove(tbl)
                        removed += 1
                        parent_p = parent_map_local.get(parent)
                        if parent_p is not None and core_tag(parent_p) == "p":
                            grand = parent_map_local.get(parent_p)
                            if grand is not None and parent_p in list(grand):
                                idx = list(grand).index(parent_p)
                                grand.remove(parent_p)
                                removed += 1
                                removed += _remove_adjacent_empty_paragraphs(grand, idx)
                if removed:
                    logger.info(
                        "section=%s removed_tables=%d by table_idx",
                        os.path.basename(sf),
                        removed,
                    )
            if fills or remove_table_indices:
                apply_fills(
                    nodes,
                    fills,
                    tree,
                    sf,
                    parent_map,
                    instruction_ids=set(),
                    remove_table_indices=set(),
                    t_ns=t_ns,
                )
            global_offset += len(nodes)

        repack_hwpx(
            extract_dir,
            output_path,
            original_compress_info=compress_info,
            original_file_order=file_order,
        )

    return output_path
