import asyncio
import logging
import os
import re
import tempfile
import time
from typing import Iterable

from pathlib import Path

from ...agent.agent import (
    agent_analyze_chunk,
    agent_fill_chunk,
    agent_generate_rag_queries,
    agent_judge_image_reference,
    agent_select_reference_images,
    agent_select_source_chunks,
)
from ...memento import search_memento_multi_query, search_memento_images_multi_query, sources_to_reference_text
from ...config import (
    DEBUG_OUTPUT_DIR,
    DEBUG_OUTPUT_ENABLED,
    IMAGE_GENERATION_ENABLED,
    IMAGE_MIN_HEIGHT_MM,
    IMAGE_MIN_WIDTH_MM,
    IMAGE_REFERENCE_ENABLED,
    LLM_VISION_ENABLED,
    MAX_CONCURRENT_LLM,
    VISION_CHUNK_PLAN_ENABLED,
)
from ...core.chunker import chunk_nodes_semantic
from ...core.filler import apply_fills
from ...core.html_screenshots import init_capture, close_capture, screenshot_chunk
from ...core.parser import parse_section, scan_header_charpr
from ...core.style_mapper import load_style_maps, log_style_summary, StyleMaps
from ...core.table_analyzer import build_table_summaries
from ...domains import detect_domain
from ...io.file import extract_hwpx, find_section_files, repack_hwpx
from .hwpx_to_html import hwpx_to_html


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


def _attach_headings_to_tables(chunks: list[list]) -> list[list]:
    """청크 끝에 남은 본문 헤딩 노드를 다음 청크(표 시작)로 이동시킨다.

    예: chunk A 끝이 ["요 약 서-1"] 본문 노드이고,
        chunk B가 표 셀로 시작하면 → 해당 노드를 B 앞으로 이동.

    단, 실제 텍스트가 있는 헤딩/제목만 이동한다.
    빈칸, 서식 마커(□, ㅇ, 가.) 등은 이동하지 않는다.
    """
    if len(chunks) < 2:
        return chunks

    for i in range(len(chunks) - 1):
        cur = chunks[i]
        nxt = chunks[i + 1]
        if not cur or not nxt:
            continue

        # 다음 청크가 표로 시작하는지 확인
        if not (getattr(nxt[0], "type", "") == "table_cell"):
            continue

        # 현재 청크 끝에서 본문(body_text) 노드를 역순으로 수집
        # 단, 실제 제목 텍스트가 있는 노드만 (빈칸/마커 제외)
        tail_nodes: list = []
        for node in reversed(cur):
            if getattr(node, "type", "") != "body_text":
                break
            text = (node.text or node.raw_text or "").strip()
            # 빈칸이거나 서식 마커(□, ㅇ, 가. 등)면 이동 대상 아님 → 중단
            if not text or text in ("□", "ㅇ") or (len(text) <= 2 and text.endswith(".")):
                break
            tail_nodes.append(node)

        if not tail_nodes:
            continue

        tail_nodes.reverse()

        for node in tail_nodes:
            cur.remove(node)
        chunks[i + 1] = tail_nodes + nxt
        logger.info(
            "[헤딩병합] 청크 %d 끝 본문 %d개 → 청크 %d 앞으로 이동: %s",
            i, len(tail_nodes), i + 1,
            [n.text[:20] if n.text else "<빈>" for n in tail_nodes],
        )

    # 비어버린 청크 제거
    return [c for c in chunks if c]


_MAX_CHUNK_NODES = 30  # 청크당 최대 노드 수


_MIN_CHUNK_NODES = 10  # 이보다 작은 서브청크는 인접 청크에 병합


def _split_large_chunks(chunks: list[list], max_nodes: int = _MAX_CHUNK_NODES) -> list[list]:
    """큰 청크를 본문 섹션 구분자(요약서, 페이지 제목 등) 기준으로 분할한다.

    분할 후 너무 작은 서브청크(제목만, 빈칸만 등)는 다음 서브청크에 병합하여
    제목과 표가 분리되지 않도록 한다.
    """
    result: list[list] = []
    for chunk in chunks:
        if len(chunk) <= max_nodes:
            result.append(chunk)
            continue

        # 분할 지점 찾기: 본문→표 전환점
        split_points: list[int] = []
        for i, node in enumerate(chunk):
            if i == 0:
                continue
            if (getattr(node, "type", "") == "table_cell"
                    and getattr(chunk[i - 1], "type", "") != "table_cell"):
                prev_tbl = getattr(chunk[i - 1], "table_idx", -1)
                curr_tbl = getattr(node, "table_idx", -1)
                if prev_tbl != curr_tbl:
                    split_points.append(i)

        if not split_points:
            result.append(chunk)
            continue

        # 분할 실행
        sub_chunks: list[list] = []
        prev = 0
        for sp in split_points:
            part = chunk[prev:sp]
            if part:
                sub_chunks.append(part)
            prev = sp
        tail = chunk[prev:]
        if tail:
            sub_chunks.append(tail)

        # 너무 작은 서브청크는 다음 서브청크 앞에 병합 (제목+표 유지)
        merged: list[list] = []
        for sc in sub_chunks:
            if merged and len(merged[-1]) < _MIN_CHUNK_NODES:
                merged[-1].extend(sc)
            else:
                merged.append(sc)
        # 마지막이 너무 작으면 이전에 병합
        if len(merged) >= 2 and len(merged[-1]) < _MIN_CHUNK_NODES:
            merged[-2].extend(merged.pop())

        result.extend(merged)
        logger.info("[청크분할] %d개 노드 → %d개 서브청크로 분할", len(chunk), len(merged))

    return result


def _merge_small_chunks(chunks: list[list], min_nodes: int = _MIN_CHUNK_NODES) -> list[list]:
    """너무 작은 청크를 인접 청크에 병합한다."""
    if len(chunks) <= 1:
        return chunks
    merged: list[list] = []
    for chunk in chunks:
        if merged and len(merged[-1]) < min_nodes:
            merged[-1].extend(chunk)
        else:
            merged.append(chunk)
    # 마지막이 너무 작으면 이전에 병합
    if len(merged) >= 2 and len(merged[-1]) < min_nodes:
        merged[-2].extend(merged.pop())
    if len(merged) != len(chunks):
        logger.info("[소청크병합] %d → %d 청크 (min=%d)", len(chunks), len(merged), min_nodes)
    return merged


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
    tenant_id: str = "",
    max_concurrent_llm: int = MAX_CONCURRENT_LLM,
    selected_file_names: list[str] | None = None,
    image_generation_enabled: bool | None = None,
    image_reference_enabled: bool | None = None,
    source_chunks: list[dict] | None = None,
) -> str:
    if not report_topic:
        raise ValueError("report_topic is required")

    if source_chunks:
        logger.info("[소스참고] 소스 참고자료 청크 %d개 수신", len(source_chunks))
    else:
        logger.info("[소스참고] 소스 참고자료 없음 — 기존 RAG/reference_text 방식 사용")

    sem = asyncio.Semaphore(max_concurrent_llm)
    screenshot_sem = asyncio.Semaphore(3)  # Playwright 동시 페이지 수 제한

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

        # 디버그: chunks 폴더 초기화
        if DEBUG_OUTPUT_ENABLED:
            import shutil
            debug_chunks_dir = Path(__file__).resolve().parent.parent.parent / DEBUG_OUTPUT_DIR / "chunks"
            if debug_chunks_dir.exists():
                shutil.rmtree(debug_chunks_dir)
            debug_chunks_dir.mkdir(parents=True, exist_ok=True)
            logger.info("[디버그] chunks 폴더 초기화 → %s", debug_chunks_dir)

        # 비전 분석용: 전체 HWPX → HTML → Playwright 브라우저 준비
        full_capture = None
        if VISION_CHUNK_PLAN_ENABLED and LLM_VISION_ENABLED:
            try:
                logger.info("[비전] HWPX → HTML 변환 시작 (inject_ids=True)")
                html_preview_path = Path(tmp_dir) / "preview_vision.html"
                hwpx_to_html(Path(hwpx_path), html_preview_path, use_lineseg=False, inject_ids=True, split_pages=False)
                html_text = html_preview_path.read_text(encoding="utf-8")
                # 디버그: 전체 HTML 저장
                if DEBUG_OUTPUT_ENABLED:
                    debug_dir = Path(__file__).resolve().parent.parent.parent / DEBUG_OUTPUT_DIR
                    debug_dir.mkdir(parents=True, exist_ok=True)
                    (debug_dir / "preview_vision.html").write_text(html_text, encoding="utf-8")
                    logger.info("[비전] 디버그 HTML 저장 → %s/preview_vision.html", debug_dir)
                logger.info("[비전] Playwright 브라우저 초기화")
                full_capture = await init_capture(html_text)
                logger.info("[비전] 브라우저 준비 완료")
            except Exception as exc:
                logger.warning("[비전] 캡처 초기화 실패 (텍스트 전용 폴백): %s", exc, exc_info=True)
                full_capture = None

        # ── 도메인 감지: 첫 번째 섹션의 노드로 문서 유형 판별 (LLM 기반) ──
        domain_type = "generic"
        domain_guide = ""
        first_section_nodes = None
        for sf_tmp in section_files:
            first_section_nodes, _, _, _ = parse_section(sf_tmp, style_maps=style_maps)
            if first_section_nodes:
                break
        if first_section_nodes:
            domain_type, domain_guide = await detect_domain(first_section_nodes)
        logger.info("[도메인] 감지 결과: type=%s, guide_len=%d", domain_type, len(domain_guide))

        logger.info("HWPX sections=%d", len(section_files))
        for sf in section_files:
            t_section = time.perf_counter()
            nodes, tree, parent_map, t_ns = parse_section(sf, style_maps=style_maps)
            if not nodes:
                continue
            log_style_summary(nodes)
            table_summaries = build_table_summaries(tree, nodes)
            chunks = chunk_nodes_semantic(nodes)
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
            # 이미지 참조 중복 방지: 이미 선택된 image_id를 청크 간 공유
            used_image_ids: set[str] = set()
            _used_image_lock = asyncio.Lock()

            async def _process_chunk(chunk: list, chunk_idx: int = 0,
                                     prev_chunk: list | None = None,
                                     next_chunk: list | None = None) -> dict:
                # 청크별 독립 스크린샷 (해당 노드만 추출하여 렌더링)
                chunk_image_b64 = ""
                if full_capture and full_capture.full_html:
                    try:
                        chunk_node_ids = [n.id for n in chunk]
                        async with screenshot_sem:
                            chunk_image_b64 = await screenshot_chunk(full_capture, chunk_node_ids)
                        if chunk_image_b64:
                            logger.info("[비전] 청크 %d 스크린샷 완료 (%.1f KB)", chunk_idx, len(chunk_image_b64) * 3 / 4 / 1024)
                            if DEBUG_OUTPUT_ENABLED:
                                import base64 as _b64
                                debug_img_dir = Path(__file__).resolve().parent.parent.parent / DEBUG_OUTPUT_DIR / "chunks"
                                debug_img_dir.mkdir(parents=True, exist_ok=True)
                                (debug_img_dir / f"chunk_{chunk_idx:03d}.png").write_bytes(_b64.b64decode(chunk_image_b64))
                    except Exception as exc:
                        logger.warning("[비전] 청크 %d 스크린샷 실패: %s", chunk_idx, exc)
                        chunk_image_b64 = ""

                logger.info("[청크처리] 청크 %d: agent_analyze_chunk 호출 시작 (nodes=%d)", chunk_idx, len(chunk))
                async with sem:
                    analysis = await agent_analyze_chunk(
                        chunk,
                        report_description=report_description,
                        table_summaries=table_summaries,
                        chunk_image_b64=chunk_image_b64,
                        prev_chunk=prev_chunk,
                        domain_type=domain_type,
                        next_chunk=next_chunk,
                    )
                logger.info("[청크처리] 청크 %d: agent_analyze_chunk 완료", chunk_idx)
                analysis = _inject_role_pairs(analysis, chunk)

                # ── RAG 쿼리 생성 + 이미지 판단을 병렬 실행 ──
                chunk_reference = reference_text
                chunk_ref_images: list[dict] = []
                _img_ref_cfg = image_reference_enabled if image_reference_enabled is not None else IMAGE_REFERENCE_ENABLED
                image_ref_enabled = _img_ref_cfg and bool((tenant_id or "").strip())
                has_tenant = bool((tenant_id or "").strip())

                async def _do_rag() -> str:
                    """RAG 쿼리 생성 → memento 검색 → reference_text 반환."""
                    if not has_tenant:
                        return reference_text
                    try:
                        async with sem:
                            rag_queries = await agent_generate_rag_queries(
                                chunk, analysis,
                                report_topic=report_topic,
                                report_description=report_description,
                            )
                        if rag_queries:
                            logger.info("[청크RAG] 쿼리 %d개 생성: %s", len(rag_queries), rag_queries)
                            rag_sources = await search_memento_multi_query(
                                rag_queries, tenant_id.strip(), top_k=5,
                                file_names=selected_file_names,
                            )
                            if rag_sources:
                                rag_text = sources_to_reference_text(rag_sources)
                                result = rag_text + "\n\n---\n\n" + reference_text if reference_text else rag_text
                                logger.info("[청크RAG] 참고자료 %d개 추가 (%.1f KB)", len(rag_sources), len(rag_text) / 1024)
                                return result
                    except Exception as exc:
                        logger.warning("[청크RAG] 실패 (기존 reference_text 사용): %s", exc)
                    return reference_text

                async def _do_image_ref() -> list[dict]:
                    """이미지 참조 판단 → 검색 → 선택 → 삽입대상 결정."""
                    ref_images: list[dict] = []
                    if not image_ref_enabled:
                        return ref_images
                    t_imgref = time.perf_counter()
                    try:
                        # STEP 1: AI가 이 청크에 기존 이미지 첨부가 필요한지 판단
                        logger.info("[이미지참조][청크%d] STEP1 — 이미지 첨부 필요 여부 판단 중...", chunk_idx)
                        async with sem:
                            img_judge = await agent_judge_image_reference(
                                chunk, analysis,
                                report_topic=report_topic,
                                report_description=report_description,
                            )
                        need = img_judge.get("need_images", False)
                        queries = img_judge.get("queries") or []
                        reason = img_judge.get("reason", "")

                        if not need:
                            logger.info("[이미지참조][청크%d] STEP1 결과 — 불필요 (reason: %s) %.1fs",
                                        chunk_idx, reason, time.perf_counter() - t_imgref)
                        elif not queries:
                            logger.warning("[이미지참조][청크%d] STEP1 결과 — 필요하나 쿼리 생성 실패 (reason: %s)",
                                           chunk_idx, reason)
                        else:
                            logger.info("[이미지참조][청크%d] STEP1 결과 — 필요 (reason: %s), 쿼리 %d개: %s",
                                        chunk_idx, reason, len(queries), queries)

                            # STEP 2: memento에서 캡션 기반 이미지 검색
                            logger.info("[이미지참조][청크%d] STEP2 — memento 이미지 검색 중...", chunk_idx)
                            candidates = await search_memento_images_multi_query(
                                queries, tenant_id.strip(), top_k=5,
                            )
                            if not candidates:
                                logger.info("[이미지참조][청크%d] STEP2 결과 — 검색 결과 없음", chunk_idx)
                            else:
                                logger.info(
                                    "[이미지참조][청크%d] STEP2 결과 — 후보 %d개:\n%s",
                                    chunk_idx, len(candidates),
                                    "\n".join(
                                        f"  [{i}] 폴더={c.get('drive_folder_name', '?')} | 출처={c.get('source_file_name', '?')} | 캡션={c.get('caption', '')[:50]}"
                                        for i, c in enumerate(candidates[:5])
                                    ),
                                )

                                # STEP 3: AI가 후보 중 적절한 이미지 선택 + 삽입 노드 결정
                                logger.info("[이미지참조][청크%d] STEP3 — 후보 중 적절한 이미지 선택 중...", chunk_idx)
                                async with sem:
                                    selected = await agent_select_reference_images(
                                        candidates, chunk, analysis,
                                        report_topic=report_topic,
                                        chunk_image_b64=chunk_image_b64,
                                    )
                                if not selected:
                                    logger.info("[이미지참조][청크%d] STEP3 결과 — 적절한 이미지 없음 (후보 모두 탈락)",
                                                chunk_idx)
                                else:
                                    # STEP 4: LLM이 지정한 target_node_id로 삽입 대상 결정
                                    node_map_local = {n.id: n for n in chunk}
                                    async with _used_image_lock:
                                        for img in selected:
                                            img_id = img.get("image_id") or img.get("image_url") or ""
                                            if img_id and img_id in used_image_ids:
                                                logger.info("[이미지참조][청크%d] 중복 제외: %s", chunk_idx, img_id[:40])
                                                continue
                                            # LLM이 지정한 노드 사용
                                            target_nid = img.get("target_node_id")
                                            target_node = node_map_local.get(target_nid) if target_nid is not None else None
                                            if target_node is None:
                                                logger.warning("[이미지참조][청크%d] target_node_id=%s 찾지 못함 — 건너뜀",
                                                               chunk_idx, target_nid)
                                                continue
                                            ref_images.append({
                                                "node": target_node,
                                                "image_url": img["image_url"],
                                                "caption": img.get("caption", ""),
                                            })
                                            if img_id:
                                                used_image_ids.add(img_id)
                                    if ref_images:
                                        target_ids = list({img["node"].id for img in ref_images})
                                        logger.info(
                                            "[이미지참조][청크%d] STEP4 완료 — %d개 이미지 → node %s 삽입 예정",
                                            chunk_idx, len(ref_images), target_ids,
                                        )
                                    else:
                                        logger.info("[이미지참조][청크%d] STEP4 — 유효한 삽입 대상 없음", chunk_idx)

                        logger.info("[이미지참조][청크%d] 완료 — 소요 %.1fs, 삽입예정 %d개",
                                    chunk_idx, time.perf_counter() - t_imgref, len(ref_images))
                    except Exception as exc:
                        logger.warning("[이미지참조][청크%d] 예외 발생 — %s (%.1fs)",
                                       chunk_idx, exc, time.perf_counter() - t_imgref)
                    return ref_images

                # RAG + 이미지참조를 병렬 실행
                chunk_reference, chunk_ref_images = await asyncio.gather(
                    _do_rag(), _do_image_ref()
                )

                # ── ① tables_to_remove 먼저 확정 (fill 전에 처리) ──
                if isinstance(analysis, dict):
                    chunk_actual_table_idxs = {
                        n.table_idx for n in chunk
                        if getattr(n, "type", "") == "table_cell" and n.table_idx >= 0
                    }
                    chunk_remove_idxs: set[int] = set()
                    for t_idx in analysis.get("tables_to_remove", []) or []:
                        if not isinstance(t_idx, int):
                            continue
                        if t_idx in chunk_actual_table_idxs:
                            chunk_remove_idxs.add(t_idx)
                            tables_to_remove.add(t_idx)
                        else:
                            logger.warning(
                                "[tables_to_remove] table_idx %d 무시 — 청크 내 실제 table_idx %s에 없음",
                                t_idx, sorted(chunk_actual_table_idxs),
                            )

                    # ── ② 삭제 대상 테이블의 노드 → skip_fill=true 강제 오버라이드 ──
                    if chunk_remove_idxs:
                        remove_node_ids_in_chunk = {
                            n.id for n in chunk
                            if getattr(n, "type", "") == "table_cell"
                            and n.table_idx in chunk_remove_idxs
                        }
                        overridden = 0
                        for item in analysis.get("nodes", []):
                            if item.get("id") in remove_node_ids_in_chunk:
                                if not item.get("skip_fill", False):
                                    overridden += 1
                                item["skip_fill"] = True
                                item["action"] = "keep"
                        if overridden:
                            logger.info(
                                "[tables_to_remove] 청크 %d: table_idx %s 삭제 결정 → %d개 노드 skip_fill 강제 전환",
                                chunk_idx, sorted(chunk_remove_idxs), overridden,
                            )

                # ── ②-b 소스 참고자료 청크 선택 (source_chunks가 있을 때) ──
                if source_chunks:
                    try:
                        async with sem:
                            selected_indices = await agent_select_source_chunks(
                                chunk, analysis, source_chunks,
                                report_topic=report_topic,
                                report_description=report_description,
                            )
                        if selected_indices:
                            source_texts = []
                            for idx in selected_indices:
                                sc = source_chunks[idx]
                                fn = sc.get("file_name", "")
                                orig = sc.get("original_text", "")
                                source_texts.append(f"[출처: {fn}]\n{orig}")
                            source_reference = "\n\n---\n\n".join(source_texts)
                            if chunk_reference:
                                chunk_reference = source_reference + "\n\n---\n\n" + chunk_reference
                            else:
                                chunk_reference = source_reference
                            logger.info(
                                "[소스참고] 청크 %d: 소스 청크 %d개 선택 → reference_text에 추가 (%.1f KB)",
                                chunk_idx, len(selected_indices), len(source_reference) / 1024,
                            )
                        else:
                            logger.info("[소스참고] 청크 %d: 관련 소스 청크 없음", chunk_idx)
                    except Exception as exc:
                        logger.warning("[소스참고] 청크 %d: 선택 실패 — %s", chunk_idx, exc)

                # ── ③ 오버라이드 반영된 analysis로 fill 생성 ──
                async with sem:
                    filled = await agent_fill_chunk(
                        analysis,
                        chunk,
                        report_topic=report_topic,
                        report_description=report_description,
                        reference_text=chunk_reference,
                        domain_guide=domain_guide,
                    )

                # ── ④ category_map / action_map 수집 ──
                # 안전장치: LLM이 data-id 대신 0-based 순번을 반환한 경우 보정
                if isinstance(analysis, dict) and analysis.get("nodes"):
                    chunk_real_ids = [n.id for n in chunk if n.id is not None]
                    resp_ids = [item.get("id") for item in analysis["nodes"] if item.get("id") is not None]
                    # 응답 id가 청크의 실제 id와 하나도 겹치지 않고, 0-based 순번처럼 보이면 매핑
                    if resp_ids and chunk_real_ids:
                        resp_id_set = {int(x) for x in resp_ids if isinstance(x, (int, float)) or (isinstance(x, str) and x.isdigit())}
                        real_id_set = set(chunk_real_ids)
                        if not (resp_id_set & real_id_set) and min(resp_id_set) == 0:
                            # 0-based 순번 → 실제 id 매핑
                            id_remap = {i: chunk_real_ids[i] for i in range(len(chunk_real_ids))}
                            remapped = 0
                            for item in analysis["nodes"]:
                                old_id = item.get("id")
                                if isinstance(old_id, (int, float)):
                                    old_id = int(old_id)
                                elif isinstance(old_id, str) and old_id.isdigit():
                                    old_id = int(old_id)
                                else:
                                    continue
                                if old_id in id_remap:
                                    item["id"] = id_remap[old_id]
                                    remapped += 1
                            if remapped:
                                logger.warning(
                                    "[id보정] 청크 %d: LLM이 0-based 순번 반환 → data-id로 매핑 (%d개)",
                                    chunk_idx, remapped,
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
                # 참조 이미지 정보를 filled에 포함
                if isinstance(filled, dict):
                    filled["_reference_images"] = chunk_ref_images
                return filled

            logger.info("[청크처리] asyncio.gather 시작: %d개 청크", len(chunks))
            results = await asyncio.gather(*[
                _process_chunk(
                    c, i,
                    prev_chunk=chunks[i - 1] if i > 0 else None,
                    next_chunk=chunks[i + 1] if i < len(chunks) - 1 else None,
                )
                for i, c in enumerate(chunks)
            ])
            all_raw = []
            all_node_ids = {n.id for n in nodes if n.id is not None}
            for chunk_fills, chunk_nodes_list in zip(results, chunks):
                fills_list = chunk_fills.get("fills") if isinstance(chunk_fills, dict) else []
                if not fills_list:
                    continue
                # id 정수 변환
                for f in fills_list:
                    if not isinstance(f, dict) or "id" not in f or "new_text" not in f:
                        continue
                    try:
                        f["id"] = int(f["id"])
                    except (ValueError, TypeError):
                        continue
                # 0-based 순번 보정 (analyze와 동일 로직)
                chunk_real_ids = [n.id for n in chunk_nodes_list if n.id is not None]
                fill_ids = {f["id"] for f in fills_list if isinstance(f, dict) and isinstance(f.get("id"), int)}
                if fill_ids and chunk_real_ids and not (fill_ids & all_node_ids) and min(fill_ids) == 0:
                    id_remap = {i: chunk_real_ids[i] for i in range(len(chunk_real_ids))}
                    for f in fills_list:
                        if isinstance(f, dict) and isinstance(f.get("id"), int) and f["id"] in id_remap:
                            f["id"] = id_remap[f["id"]]
                    logger.warning("[fill id보정] 0-based 순번 → data-id로 매핑")
                for f in fills_list:
                    if isinstance(f, dict) and isinstance(f.get("id"), int) and "new_text" in f:
                        all_raw.append(f)
            # 참조 이미지 수집
            all_reference_images: list[dict] = []
            for fills in results:
                if isinstance(fills, dict):
                    all_reference_images.extend(fills.get("_reference_images") or [])
            if all_reference_images:
                logger.info(
                    "[이미지참조] ━━ 전체 수집 완료: %d개 참조 이미지 (대상 노드: %s)",
                    len(all_reference_images),
                    list({img.get("node").id for img in all_reference_images if img.get("node")}),
                )
            elif (image_reference_enabled if image_reference_enabled is not None else IMAGE_REFERENCE_ENABLED):
                logger.info("[이미지참조] ━━ 전체 수집 완료: 참조 이미지 없음")

            # tables_to_remove에 포함된 표의 모든 노드는 강제 skip
            remove_node_ids_by_table: set[int] = set()
            if tables_to_remove:
                for n in nodes:
                    if getattr(n, "type", "") == "table_cell" and n.table_idx in tables_to_remove:
                        remove_node_ids_by_table.add(n.id)
                if remove_node_ids_by_table:
                    logger.info(
                        "[후처리] tables_to_remove %s → %d개 노드 fill 차단",
                        tables_to_remove, len(remove_node_ids_by_table),
                    )

            all_fills: list[dict] = []
            FILLABLE = ("fill", "placeholder")
            ACTION_FILLABLE = ("write", "replace")

            # 안전장치: analyze가 완전 실패하여 category_map이 비어있으면
            # fill 결과를 그대로 통과시킴 (tables_to_remove만 체크)
            analyze_failed = not category_map
            if analyze_failed and all_raw:
                logger.warning(
                    "[fill필터] ⚠ category_map 비어있음 (analyze 실패). "
                    "fill 결과 %d개를 필터 없이 통과시킵니다.",
                    len(all_raw),
                )

            for f in all_raw:
                if f["id"] in remove_node_ids_by_table:
                    logger.info("[fill필터] node %d: tables_to_remove 차단", f["id"])
                    continue
                if analyze_failed:
                    # analyze 실패 시 fill 결과 그대로 통과
                    all_fills.append(f)
                    continue
                cat, skip = category_map.get(f["id"], ("", False))
                action = action_map.get(f["id"], "")
                if action and action not in ACTION_FILLABLE:
                    logger.info("[fill필터] node %d: action=%s 차단", f["id"], action)
                    continue
                if cat not in FILLABLE or skip:
                    logger.info("[fill필터] node %d: cat=%s skip=%s 차단", f["id"], cat, skip)
                    continue
                all_fills.append(f)
            logger.info("[fill필터] all_raw=%d → all_fills=%d (통과 ID: %s)",
                        len(all_raw), len(all_fills), [f["id"] for f in all_fills])

            image_insert_enabled = image_generation_enabled if image_generation_enabled is not None else IMAGE_GENERATION_ENABLED
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
                header_path=header_path,
            )

            # 참조 이미지 삽입 (Gemini 생성과 별개)
            if all_reference_images:
                from ...images import apply_reference_images_to_section

                logger.info(
                    "[이미지참조] ━━ Section %s — %d개 참조 이미지 다운로드 & 삽입 시작",
                    os.path.basename(sf), len(all_reference_images),
                )
                t_ref_insert = time.perf_counter()
                ref_inserted = apply_reference_images_to_section(
                    tree, sf, parent_map, all_reference_images,
                    t_ns=t_ns, write_back=True,
                )
                logger.info(
                    "[이미지참조] ━━ Section %s — 삽입 완료: %d/%d건 성공 (%.1fs)",
                    os.path.basename(sf), ref_inserted, len(all_reference_images),
                    time.perf_counter() - t_ref_insert,
                )

            logger.info("Section %s done in %.2fs", os.path.basename(sf), time.perf_counter() - t_section)

        # Playwright 브라우저 종료
        if full_capture:
            try:
                await close_capture(full_capture)
            except Exception:
                pass

        repack_hwpx(
            extract_dir,
            output_path,
            original_compress_info=compress_info,
            original_file_order=file_order,
        )

    # 디버그: 완성본 hwpx + HTML 저장
    if DEBUG_OUTPUT_ENABLED:
        import shutil
        debug_dir = Path(__file__).resolve().parent.parent.parent / DEBUG_OUTPUT_DIR
        result_dir = debug_dir / "result"
        if result_dir.exists():
            shutil.rmtree(result_dir)
        result_dir.mkdir(parents=True, exist_ok=True)
        # hwpx 복사
        result_hwpx = result_dir / Path(output_path).name
        shutil.copy2(output_path, result_hwpx)
        # HTML 변환본 생성
        result_html = result_dir / (Path(output_path).stem + ".html")
        try:
            hwpx_to_html(Path(output_path), result_html, use_lineseg=False, inject_ids=True)
            logger.info("[디버그] 완성본 저장 → %s (hwpx + html)", result_dir)
        except Exception as exc:
            logger.warning("[디버그] 완성본 HTML 변환 실패: %s", exc)

    return output_path
