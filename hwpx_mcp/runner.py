import asyncio
import logging
import os
import re
import tempfile
import time
from typing import Iterable

from .agent.agent import agent_analyze_chunk, agent_fill_chunk, agent_chunk_plan
from .config import (
    IMAGE_GENERATION_ENABLED,
    IMAGE_MIN_HEIGHT_MM,
    IMAGE_MIN_WIDTH_MM,
    MAX_CONCURRENT_LLM,
)
from .core.chunker import chunk_nodes, chunk_nodes_by_plan
from .core.filler import apply_fills
from .core.parser import parse_section, scan_header_charpr
from .core.style_mapper import load_style_maps, log_style_summary, StyleMaps
from .core.table_analyzer import build_table_summaries
from .io.file import extract_hwpx, find_section_files, repack_hwpx


logger = logging.getLogger("process-gpt-office-mcp")


def _merge_chunk_results(results: Iterable[dict]) -> list[dict]:
    fills: list[dict] = []
    for res in results:
        if isinstance(res, dict):
            chunk_fills = res.get("fills")
            if isinstance(chunk_fills, list):
                fills.extend(chunk_fills)
    return fills


def _merge_table_chunks(chunks: list[list]) -> list[list]:
    if not chunks:
        return chunks

    parent = list(range(len(chunks)))

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    table_to_chunks: dict[int, list[int]] = {}
    for idx, chunk in enumerate(chunks):
        table_idxs = {n.table_idx for n in chunk if getattr(n, "type", "") == "table_cell" and n.table_idx >= 0}
        for t in table_idxs:
            table_to_chunks.setdefault(t, []).append(idx)

    for indices in table_to_chunks.values():
        if len(indices) < 2:
            continue
        base = indices[0]
        for other in indices[1:]:
            union(base, other)

    groups: dict[int, list[int]] = {}
    for i in range(len(chunks)):
        groups.setdefault(find(i), []).append(i)

    if all(len(g) == 1 for g in groups.values()):
        return chunks

    merged_chunks: dict[int, list] = {}
    drop: set[int] = set()
    for g in groups.values():
        g_sorted = sorted(g)
        keep = g_sorted[0]
        combined: list = []
        for idx in g_sorted:
            combined.extend(chunks[idx])
            if idx != keep:
                drop.add(idx)
        merged_chunks[keep] = combined

    new_chunks: list[list] = []
    for i, ch in enumerate(chunks):
        if i in drop:
            continue
        new_chunks.append(merged_chunks.get(i, ch))
    logger.info("Table chunk merge: %d -> %d", len(chunks), len(new_chunks))
    return new_chunks


_IMAGE_MARKER_RE = re.compile(r"\[IMAGE(?::([^\]]+))?\]")


def _cell_range(col: int, span: int) -> tuple[int, int]:
    span = max(1, span)
    return col, col + span - 1


def _ranges_overlap(a: tuple[int, int], b: tuple[int, int]) -> bool:
    return not (a[1] < b[0] or b[1] < a[0])


def _estimate_max_chars(node) -> int:
    width_mm = getattr(node, "cell_width_mm", 0) or 0
    if width_mm <= 0:
        return 30
    est = int(round(width_mm / 3.5))
    return max(10, min(40, est))


def _build_summary_detail_pairs(nodes: list) -> dict[int, int]:
    table_nodes: dict[int, list] = {}
    for n in nodes:
        if getattr(n, "type", "") == "table_cell" and n.table_idx >= 0:
            table_nodes.setdefault(n.table_idx, []).append(n)

    pairs: dict[int, int] = {}
    for tbl_nodes in table_nodes.values():
        for cell in tbl_nodes:
            h = getattr(cell, "cell_height_mm", 0) or 0
            if h <= 0 or h > 10:
                continue
            base_row = cell.row + max(1, cell.cell_row_span)
            col_range = _cell_range(cell.col, cell.cell_col_span)
            candidates = [
                n for n in tbl_nodes
                if n.row >= base_row
                and _ranges_overlap(col_range, _cell_range(n.col, n.cell_col_span))
            ]
            if not candidates:
                continue
            candidates.sort(key=lambda n: (n.row, n.col))
            detail = candidates[0]
            detail_h = getattr(detail, "cell_height_mm", 0) or 0
            if detail_h < 12:
                continue
            pairs[cell.id] = detail.id
    return pairs


def _inject_role_pairs(analysis: dict, nodes: list) -> dict:
    if not isinstance(analysis, dict):
        return analysis
    items = analysis.get("nodes")
    if not isinstance(items, list) or not items:
        return analysis
    node_map = {n.id: n for n in nodes if hasattr(n, "id")}
    id_to_item = {item.get("id"): item for item in items if isinstance(item, dict)}
    pairs = _build_summary_detail_pairs(nodes)
    for header_id, detail_id in pairs.items():
        header_item = id_to_item.get(header_id)
        detail_item = id_to_item.get(detail_id)
        if header_item is not None:
            header_item.setdefault("role", "summary_header")
            header_item.setdefault("detail_node", detail_id)
            if "max_chars" not in header_item:
                header_node = node_map.get(header_id)
                if header_node is not None:
                    header_item["max_chars"] = _estimate_max_chars(header_node)
        if detail_item is not None:
            detail_item.setdefault("role", "detail_body")
            detail_item.setdefault("header_node", header_id)
    return analysis


def _extract_image_markers(text: str) -> tuple[str, list[dict]]:
    markers: list[dict] = []

    def _replace(match: re.Match) -> str:
        content = (match.group(1) or "").strip()
        prompt = ""
        caption = ""
        ratio = ""
        if content:
            parts = [p.strip() for p in content.split("|")]
            prompt = parts[0] if len(parts) > 0 else ""
            caption = parts[1] if len(parts) > 1 else ""
            ratio = parts[2] if len(parts) > 2 else ""
        markers.append({"prompt": prompt, "caption": caption, "ratio": ratio})
        return ""

    cleaned = _IMAGE_MARKER_RE.sub(_replace, text)
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip()
    return cleaned, markers


async def process_hwpx_file(
    hwpx_path: str,
    output_path: str,
    *,
    report_topic: str,
    report_description: str = "",
    reference_text: str = "",
    max_concurrent_llm: int = MAX_CONCURRENT_LLM,
) -> str:
    if not report_topic:
        raise ValueError("report_topic is required")

    sem = asyncio.Semaphore(max_concurrent_llm)

    with tempfile.TemporaryDirectory() as tmp_dir:
        extract_dir = os.path.join(tmp_dir, "hwpx")
        compress_info, file_order = extract_hwpx(hwpx_path, extract_dir)
        section_files = find_section_files(extract_dir)

        if not section_files:
            top_entries = []
            contents_entries = []
            try:
                top_entries = os.listdir(extract_dir)
                contents_dir = os.path.join(extract_dir, "Contents")
                if os.path.isdir(contents_dir):
                    contents_entries = os.listdir(contents_dir)
            except FileNotFoundError:
                top_entries = []
            raise RuntimeError(
                "No section files found in HWPX "
                f"(top_entries={top_entries}, contents_entries={contents_entries})"
            )

        header_path = os.path.join(extract_dir, "Contents", "header.xml")
        _, charpr_warnings = scan_header_charpr(header_path)
        if charpr_warnings:
            logger.info("Header warnings: %s", ", ".join(charpr_warnings[:5]))

        style_maps: StyleMaps | None = None
        if os.path.exists(header_path):
            style_maps = load_style_maps(header_path)

        logger.info("HWPX sections=%d", len(section_files))
        for sf in section_files:
            t_section = time.perf_counter()
            nodes, tree, parent_map, t_ns = parse_section(sf, style_maps=style_maps)
            if not nodes:
                continue
            log_style_summary(nodes)
            table_summaries = build_table_summaries(tree, nodes)
            plan = await agent_chunk_plan(nodes, table_summaries=table_summaries)
            chunks, errors = chunk_nodes_by_plan(nodes, plan)
            if errors:
                logger.info("Chunk plan failed: %s", "; ".join(errors[:5]))
                chunks = chunk_nodes(nodes, max_nodes=len(nodes))
            else:
                chunks = _merge_table_chunks(chunks)
            logger.info(
                "Section %s nodes=%d chunks=%d",
                os.path.basename(sf),
                len(nodes),
                len(chunks),
            )

            category_map: dict[int, tuple[str, bool]] = {}
            action_map: dict[int, str] = {}
            tables_to_remove: set[int] = set()
            delete_node_ids: set[int] = set()
            delete_table_indices: set[int] = set()
            delete_if_no_image: set[int] = set()
            image_placeholder_map: dict[int, dict] = {}

            async def _process_chunk(chunk: list) -> dict:
                async with sem:
                    analysis = await agent_analyze_chunk(
                        chunk,
                        report_description=report_description,
                        table_summaries=table_summaries,
                    )
                analysis = _inject_role_pairs(analysis, chunk)
                async with sem:
                    filled = await agent_fill_chunk(
                        analysis,
                        chunk,
                        report_topic=report_topic,
                        report_description=report_description,
                        reference_text=reference_text,
                    )
                if isinstance(analysis, dict):
                    for item in analysis.get("nodes", []):
                        nid = item.get("id")
                        if nid is None:
                            continue
                        category_map[nid] = (
                            (item.get("category") or "").strip(),
                            bool(item.get("skip_fill", False)),
                        )
                        action_map[nid] = (item.get("action") or "").strip().lower()
                        if bool(item.get("delete")):
                            delete_node_ids.add(nid)
                            node = next((n for n in chunk if n.id == nid), None)
                            if node is not None and getattr(node, "type", "") == "table_cell":
                                if node.table_idx >= 0:
                                    delete_table_indices.add(node.table_idx)
                        if bool(item.get("delete_if_no_image")):
                            delete_if_no_image.add(nid)
                        category = (item.get("category") or "").strip().lower()
                        if category == "image_placeholder" or action_map[nid] == "insert_image":
                            image_placeholder_map[nid] = {
                                "prompt": (item.get("image_prompt") or "").strip(),
                                "caption": (item.get("image_caption") or "").strip(),
                                "ratio": (item.get("image_ratio") or "").strip(),
                            }
                    for t_idx in analysis.get("tables_to_remove", []) or []:
                        if isinstance(t_idx, int):
                            tables_to_remove.add(t_idx)
                return filled

            results = await asyncio.gather(*[_process_chunk(c) for c in chunks])
            all_raw = [
                f for fills in results
                for f in (fills.get("fills") if isinstance(fills, dict) else [])
                if isinstance(f, dict) and "id" in f and "new_text" in f
            ]
            all_fills: list[dict] = []
            FILLABLE = ("fill", "placeholder", "checkbox")
            ACTION_FILLABLE = ("write", "replace")
            for f in all_raw:
                cat, skip = category_map.get(f["id"], ("", False))
                action = action_map.get(f["id"], "")
                if action and action not in ACTION_FILLABLE:
                    continue
                if cat not in FILLABLE or skip:
                    continue
                all_fills.append(f)

            image_insert_enabled = IMAGE_GENERATION_ENABLED
            node_by_id = {n.id: n for n in nodes if n.id is not None}
            image_inserts: list[dict] = []
            marker_total = 0
            skipped_small = 0

            def _image_size_ok(node) -> bool:
                if getattr(node, "type", "") != "table_cell":
                    return True
                w = getattr(node, "cell_width_mm", 0) or 0
                h = getattr(node, "cell_height_mm", 0) or 0
                if w > 0 and w < IMAGE_MIN_WIDTH_MM:
                    return False
                if h > 0 and h < IMAGE_MIN_HEIGHT_MM:
                    return False
                return True

            for f in all_fills:
                text = f.get("new_text") or ""
                if "[IMAGE" not in text:
                    continue
                cleaned, markers = _extract_image_markers(text)
                if not markers:
                    continue
                f["new_text"] = cleaned
                if not image_insert_enabled:
                    marker_total += len(markers)
                    continue
                node = node_by_id.get(f.get("id"))
                if node is None:
                    continue
                if not _image_size_ok(node):
                    skipped_small += len(markers)
                    marker_total += len(markers)
                    continue
                fallback_text = (cleaned or node.text or node.raw_text or report_topic).strip()
                for m in markers:
                    prompt = (m.get("prompt") or fallback_text).strip()
                    caption = (m.get("caption") or "").strip()
                    ratio = (m.get("ratio") or "").strip()
                    if ratio not in ("16:9", "4:3"):
                        ratio = "16:9"
                    if not prompt:
                        continue
                    image_inserts.append(
                        {"node": node, "prompt": prompt, "caption": caption, "ratio": ratio}
                    )
                marker_total += len(markers)

            if image_insert_enabled:
                for nid, meta in image_placeholder_map.items():
                    node = node_by_id.get(nid)
                    if node is None:
                        continue
                    if not _image_size_ok(node):
                        skipped_small += 1
                        continue
                    prompt = meta.get("prompt") or (node.text or node.raw_text or report_topic)
                    prompt = (prompt or "").strip()
                    if not prompt:
                        continue
                    caption = (meta.get("caption") or "").strip()
                    ratio = (meta.get("ratio") or "").strip()
                    if ratio not in ("16:9", "4:3"):
                        ratio = "16:9"
                    image_inserts.append(
                        {"node": node, "prompt": prompt, "caption": caption, "ratio": ratio}
                    )

            if not image_insert_enabled:
                if marker_total:
                    logger.info("Image markers removed (count=%d)", marker_total)
            else:
                if marker_total:
                    logger.info(
                        "Image markers detected=%d, to_generate=%d, skipped_small=%d",
                        marker_total,
                        len(image_inserts),
                        skipped_small,
                    )
                else:
                    if image_inserts:
                        logger.info(
                            "Image placeholders to_generate=%d, skipped_small=%d",
                            len(image_inserts),
                            skipped_small,
                        )
                    else:
                        logger.info("Image markers detected=0")
                if not image_inserts:
                    logger.info("No image markers to insert")

            if delete_if_no_image:
                if not image_insert_enabled:
                    delete_node_ids.update(delete_if_no_image)
                else:
                    no_image_ids = {
                        nid for nid in delete_if_no_image
                        if nid not in {item.get("node").id for item in image_inserts if item.get("node")}
                    }
                    delete_node_ids.update(no_image_ids)
            if delete_node_ids:
                for nid in list(delete_node_ids):
                    node = node_by_id.get(nid)
                    if node is not None and getattr(node, "type", "") == "table_cell":
                        if node.table_idx >= 0:
                            delete_table_indices.add(node.table_idx)

            apply_fills(
                nodes,
                all_fills,
                tree,
                sf,
                parent_map,
                instruction_ids={
                    nid for nid, (cat, _skip) in category_map.items() if cat == "instruction"
                },
                remove_table_indices=tables_to_remove.union(delete_table_indices),
                remove_node_ids=delete_node_ids,
                t_ns=t_ns,
                image_inserts=image_inserts,
            )
            logger.info("Section %s done in %.2fs", os.path.basename(sf), time.perf_counter() - t_section)

        repack_hwpx(
            extract_dir,
            output_path,
            original_compress_info=compress_info,
            original_file_order=file_order,
        )

    return output_path
